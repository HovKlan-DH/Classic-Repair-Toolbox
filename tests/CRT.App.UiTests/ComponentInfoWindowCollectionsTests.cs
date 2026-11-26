using System.Security.Cryptography;
using Avalonia.Controls;
using Avalonia.Interactivity;
using CRT;
using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Collect captures into" in the component info window (ComponentInfoWindow.Collections.cs): the board's collections to pick
// from once there is an oscilloscope, a new one made for the board, each capture kept in the one picked - with the
// settings read back - and taken out again from the banner.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentInfoWindowCollectionsTests : IDisposable
{
    private const string BoardKey = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();

    private readonly CollectionStore thisStore;

    private readonly List<string> thisOpened = [];

    public ComponentInfoWindowCollectionsTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
        UserSettings.EnableNetworkConnectedOscilloscopeTab = true;
        this.thisStore = new CollectionStore(this.thisWorkspace.Path_("collections"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
        this.thisWorkspace.Dispose();
    }

    private CollectionContext Context() =>
        new(
            this.thisStore,
            BoardKey,
            "Commodore 64",
            "250407",
            "PAL",
            () => Task.FromResult<ScopeSettingsReading?>(new ScopeSettingsReading(50e-6, 2, 1)),
            this.thisOpened.Add);

    // A popup of U8, its oscilloscope session seen.
    private ComponentInfoWindow Popup(bool sessionSeen = true)
    {
        var window = new ComponentInfoWindow();

        window.SetComponent("U8", "U8", [new ComponentEntry { BoardLabel = "U8", Category = "IC" }], [], [], [], "PAL",
            Path.GetTempPath(), [], string.Empty, string.Empty);
        window.UseCollections(this.Context());
        window.UpdateOscilloscopeSessionTitleState(sessionSeen, sessionSeen);

        return window;
    }

    private static List<string> Choices(ComponentInfoWindow window) =>
        window.FindControl<ComboBox>("CollectionComboBox")!.Items.Cast<object>().Select(item => item.ToString()!).ToList();

    [Fact]
    public void The_row_shows_once_there_is_an_oscilloscope_to_capture_from()
    {
        UiTest.Run(() =>
        {
            Assert.False(this.Popup(sessionSeen: false).FindControl<Border>("CollectionRow")!.IsVisible);

            ComponentInfoWindow window = this.Popup();

            Assert.True(window.FindControl<Border>("CollectionRow")!.IsVisible);
            Assert.Equal([ComponentInfoWindow.NoCollectionChoice, ComponentInfoWindow.NewCollectionChoice], Choices(window));
            Assert.False(window.FindControl<Button>("ShowCollectionButton")!.IsEnabled);
        });
    }

    [Fact]
    public void A_new_collection_is_made_for_the_board_named_after_the_component_and_stays_picked()
    {
        UiTest.Run(() =>
        {
            ComponentInfoWindow window = this.Popup();
            var box = window.FindControl<ComboBox>("CollectionComboBox")!;

            box.SelectedItem = box.Items.Cast<ComponentInfoWindow.CollectionChoice>().Single(choice => choice.IsNew);

            LocalCollection made = Assert.Single(this.thisStore.OfBoard("Commodore 64", "250407"));

            Assert.StartsWith("U8 captures, ", made.Title, StringComparison.Ordinal);
            Assert.Equal("PAL", made.Region);
            Assert.Equal(made.LocalId, UserSettings.CollectionOfBoard(BoardKey));
            Assert.Equal($"{made.Title} (0 captures)", box.SelectedItem?.ToString());

            // Another popup of the board - another component, another day - fills the same one.
            Assert.Equal($"{made.Title} (0 captures)", this.Popup().FindControl<ComboBox>("CollectionComboBox")!.SelectedItem?.ToString());

            var show = window.FindControl<Button>("ShowCollectionButton")!;

            Assert.True(show.IsEnabled);
            show.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Assert.Equal([made.LocalId], this.thisOpened);
        });
    }

    [Fact]
    public async Task A_capture_goes_into_the_collection_picked_with_its_settings_and_comes_out_again()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, "No picture", DateTimeOffset.UtcNow)!;
        UserSettings.SetCollectionOfBoard(BoardKey, collection.LocalId);

        string saved = this.thisWorkspace.Path_("U8_12_RAS.png");
        string baseline = this.thisWorkspace.Path_("baseline.png");

        File.WriteAllBytes(saved, [1, 2, 3]);
        File.WriteAllBytes(baseline, [4, 5, 6]);

        var shown = new ComponentImageEntry { BoardLabel = "U8", Pin = "12", Name = "/RAS", File = "Scope baseline/U8_12_RAS.png", TimeDiv = "100us" };

        await UiTest.RunAsync(async () =>
        {
            ComponentInfoWindow window = this.Popup();

            await window.CollectCaptureAsync(shown, baseline, saved);

            LocalCapture kept = Assert.Single(this.thisStore.Find(collection.LocalId)!.Captures);

            Assert.Equal(("U8", "12", "50us", "2V", "1V"), (kept.ComponentLabel, kept.Pin, kept.TimePerDivision, kept.VoltsPerDivision, kept.TriggerLevel));
            Assert.Equal(Convert.ToHexStringLower(SHA256.HashData([4, 5, 6])), kept.BaselineSha256);
            Assert.Equal("100us", kept.BaselineTimePerDivision);
            Assert.Equal("Added to [No picture]", window.CollectedTextForTests);
            Assert.Equal("No picture (1 capture)", window.FindControl<ComboBox>("CollectionComboBox")!.SelectedItem?.ToString());

            var takeOut = window.FindControl<Button>("TakeOutOfCollectionButton")!;

            Assert.True(takeOut.IsVisible);
            takeOut.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Empty(this.thisStore.Find(collection.LocalId)!.Captures);
            Assert.Equal("Taken out of [No picture]", window.CollectedTextForTests);
            Assert.False(takeOut.IsVisible);
        });
    }

    [Fact]
    public async Task With_no_collection_picked_a_capture_goes_nowhere_else()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, null, DateTimeOffset.UtcNow)!;
        string saved = this.thisWorkspace.Path_("capture.png");
        File.WriteAllBytes(saved, [1]);

        await UiTest.RunAsync(async () =>
        {
            ComponentInfoWindow window = this.Popup();

            await window.CollectCaptureAsync(new ComponentImageEntry { BoardLabel = "U8", Pin = "1" }, null, saved);

            Assert.Empty(this.thisStore.Find(collection.LocalId)!.Captures);
            Assert.Null(window.CollectedTextForTests);
        });
    }

    // A collection deleted elsewhere is no pick: the drop-down falls back to "None".
    [Fact]
    public void A_collection_picked_and_since_deleted_is_no_longer_picked()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, "Gone", DateTimeOffset.UtcNow)!;
        UserSettings.SetCollectionOfBoard(BoardKey, collection.LocalId);
        this.thisStore.Delete(collection.LocalId);

        UiTest.Run(() =>
        {
            ComponentInfoWindow window = this.Popup();

            Assert.Equal(ComponentInfoWindow.NoCollectionChoice, window.FindControl<ComboBox>("CollectionComboBox")!.SelectedItem?.ToString());
        });
    }
}
