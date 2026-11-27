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

    // An EMPTY file, never a missing one: LoadFrom keeps the settings in memory when there is no file, and the oscilloscope
    // tab and the collections picked would reach the next test (code review, 2026-11-27).
    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
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

    // The choices are shown again at every push of the oscilloscope's state and as the comment check answers - never under an
    // open list, whose item clicked would be another: they wait for it to close (code review, 2026-11-27).
    [Fact]
    public void The_choices_wait_while_the_list_is_open()
    {
        UiTest.Run(() =>
        {
            ComponentInfoWindow window = this.Popup();
            var box = window.FindControl<ComboBox>("CollectionComboBox")!;

            window.Show();

            try
            {
                box.IsDropDownOpen = true;
                Assert.True(box.IsDropDownOpen);

                LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, "No picture", DateTimeOffset.UtcNow)!;
                window.ShowCollectionChoices();

                Assert.Equal([ComponentInfoWindow.NoCollectionChoice, ComponentInfoWindow.NewCollectionChoice], Choices(window));

                box.IsDropDownOpen = false;

                Assert.Equal([ComponentInfoWindow.NoCollectionChoice, $"{made.Title} (0 captures)", ComponentInfoWindow.NewCollectionChoice], Choices(window));
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A pick is made from the open list: the arrow keys on the closed drop-down step nothing - each pass over "New collection"
    // made an empty collection, picked for the board (code review, 2026-11-27).
    [Fact]
    public void The_arrow_keys_on_the_closed_list_make_no_collection()
    {
        UiTest.Run(() =>
        {
            ComponentInfoWindow window = this.Popup();
            var box = window.FindControl<ComboBox>("CollectionComboBox")!;

            window.Show();

            try
            {
                // With the oscilloscope's keyboard control off - on, the popup takes Up and Down for the trigger level first.
                window.FindControl<CheckBox>("NumpadOscilloscopeSwitch")!.IsChecked = false;
                Assert.True(box.Focus());

                for (int i = 0; i < 3; i++)
                    Avalonia.Headless.HeadlessWindowExtensions.KeyPress(window, Avalonia.Input.Key.Down, Avalonia.Input.RawInputModifiers.None, Avalonia.Input.PhysicalKey.ArrowDown, keySymbol: null);

                Assert.False(box.IsDropDownOpen);
                Assert.Empty(this.thisStore.OfBoard("Commodore 64", "250407"));
                Assert.Equal(ComponentInfoWindow.NoCollectionChoice, box.SelectedItem?.ToString());
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public async Task A_capture_goes_into_the_collection_picked_with_its_settings_and_comes_out_again()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, "No picture", DateTimeOffset.UtcNow)!;
        UserSettings.SetCollectionOfBoard(BoardKey, collection.LocalId);

        string saved = this.thisWorkspace.Path_("U8_12_RAS.png");
        byte[] pinout = CRT.Data.Tests.TestPng.Rgb(20, 12, seed: 7);
        string pinoutSha = Convert.ToHexStringLower(SHA256.HashData(pinout));

        File.WriteAllBytes(saved, CRT.Data.Tests.TestPng.Rgb(16, 10));

        var shown = new ComponentImageEntry { BoardLabel = "U8", Pin = "12", Name = "/RAS", File = "Scope baseline/U8_12_RAS.png", TimeDiv = "100us" };

        await UiTest.RunAsync(async () =>
        {
            ComponentInfoWindow window = this.Popup();

            // The drawn pinout, pictured - the headless platform draws nothing of its own (owner request, 2026-11-27).
            window.PinoutPictureOverrideForTests = () => pinout;

            window.CaptureOnScreenForTests(saved);
            await window.CollectCaptureAsync(shown, saved);

            LocalCapture kept = Assert.Single(this.thisStore.Find(collection.LocalId)!.Captures);

            Assert.Equal(("U8", "12", "50us", "2V", "1V"), (kept.ComponentLabel, kept.Pin, kept.TimePerDivision, kept.VoltsPerDivision, kept.TriggerLevel));
            Assert.Equal(pinoutSha, kept.PinoutSha256);
            Assert.True(File.Exists(this.thisStore.CapturePath(collection.LocalId, pinoutSha)));
            Assert.Equal("Added to [No picture]", window.CollectedTextForTests);
            Assert.Equal("No picture (1 capture)", window.FindControl<ComboBox>("CollectionComboBox")!.SelectedItem?.ToString());

            var takeOut = window.FindControl<Button>("TakeOutOfCollectionButton")!;

            Assert.True(takeOut.IsVisible);
            takeOut.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Empty(this.thisStore.Find(collection.LocalId)!.Captures);
            Assert.Equal("Taken out of [No picture]", window.CollectedTextForTests);
            Assert.False(takeOut.IsVisible);

            // Its pinout drawing went with it - no other capture shows it.
            Assert.False(File.Exists(this.thisStore.CapturePath(collection.LocalId, pinoutSha)));
        });
    }

    // A new collection is of the region THIS popup shows - its own buttons switch it, and the context's is what Main showed as
    // the popup opened (code review, 2026-11-27: an NTSC popup's collection read "(PAL)" on its page).
    [Fact]
    public void A_new_collection_is_of_the_region_the_popup_shows()
    {
        UiTest.Run(() =>
        {
            ComponentInfoWindow window = this.Popup();
            window.FindControl<Button>("NtscRegionButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            var box = window.FindControl<ComboBox>("CollectionComboBox")!;
            box.SelectedItem = box.Items.Cast<ComponentInfoWindow.CollectionChoice>().Single(choice => choice.IsNew);

            Assert.Equal("NTSC", Assert.Single(this.thisStore.OfBoard("Commodore 64", "250407")).Region);
        });
    }

    // The pick is the board's: every popup of it is told, so none shows a pick that is no longer where its captures go (code
    // review, 2026-11-27).
    [Fact]
    public void A_pick_is_told_to_the_other_popups()
    {
        var told = new List<string?>();

        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();
            window.SetComponent("U8", "U8", [new ComponentEntry { BoardLabel = "U8", Category = "IC" }], [], [], [], "PAL",
                Path.GetTempPath(), [], string.Empty, string.Empty);
            window.UseCollections(this.Context() with { CollectionChanged = told.Add });
            window.UpdateOscilloscopeSessionTitleState(true, true);

            var box = window.FindControl<ComboBox>("CollectionComboBox")!;
            box.SelectedItem = box.Items.Cast<ComponentInfoWindow.CollectionChoice>().Single(choice => choice.IsNew);

            // Told with no collection's content changed - an open collection window is left as it is (code review, 2026-11-27).
            Assert.Equal([(string?)null], told);
        });
    }

    // Two ground pins at the same settings give the very same screen - one capture. "Take out" is offered for neither the
    // second, nor it would take the first out (code review, 2026-11-27).
    [Fact]
    public async Task The_same_screen_twice_is_said_to_be_in_already_and_offers_no_take_out()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, "No picture", DateTimeOffset.UtcNow)!;
        UserSettings.SetCollectionOfBoard(BoardKey, collection.LocalId);

        string first = this.thisWorkspace.Path_("U8_7_GND.png");
        string second = this.thisWorkspace.Path_("U8_14_GND.png");
        byte[] screen = CRT.Data.Tests.TestPng.Rgb(16, 10, seed: 3);

        File.WriteAllBytes(first, screen);
        File.WriteAllBytes(second, screen);

        await UiTest.RunAsync(async () =>
        {
            ComponentInfoWindow window = this.Popup();

            window.CaptureOnScreenForTests(first);
            await window.CollectCaptureAsync(new ComponentImageEntry { BoardLabel = "U8", Pin = "7" }, first);
            window.CaptureOnScreenForTests(second);
            await window.CollectCaptureAsync(new ComponentImageEntry { BoardLabel = "U8", Pin = "14" }, second);

            Assert.Equal("7", Assert.Single(this.thisStore.Find(collection.LocalId)!.Captures).Pin);
            Assert.Equal("The very same screen is in [No picture] already", window.CollectedTextForTests);
            Assert.False(window.FindControl<Button>("TakeOutOfCollectionButton")!.IsVisible);
        });
    }

    // What the server would refuse is said at the capture (CollectionPng), and kept out.
    [Fact]
    public async Task A_screen_no_collection_takes_is_said_and_kept_out()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, "No picture", DateTimeOffset.UtcNow)!;
        UserSettings.SetCollectionOfBoard(BoardKey, collection.LocalId);

        string saved = this.thisWorkspace.Path_("U8_12_RAS.png");
        File.WriteAllBytes(saved, [1, 2, 3]);

        await UiTest.RunAsync(async () =>
        {
            ComponentInfoWindow window = this.Popup();

            window.CaptureOnScreenForTests(saved);
            await window.CollectCaptureAsync(new ComponentImageEntry { BoardLabel = "U8", Pin = "12" }, saved);

            Assert.Empty(this.thisStore.Find(collection.LocalId)!.Captures);
            Assert.Equal($"Not added to [No picture]: {CollectionPng.NotAPicture}", window.CollectedTextForTests);
        });
    }

    // Moved to another image while the oscilloscope's settings were read, the capture goes into the collection all the same -
    // and the banner stays off the image now shown, whose "Take out" would take this capture out (code review, 2026-11-27).
    [Fact]
    public async Task A_capture_moved_on_from_puts_no_banner_under_the_image_shown()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, "No picture", DateTimeOffset.UtcNow)!;
        UserSettings.SetCollectionOfBoard(BoardKey, collection.LocalId);

        string saved = this.thisWorkspace.Path_("U8_5_D0.png");
        File.WriteAllBytes(saved, CRT.Data.Tests.TestPng.Rgb(16, 10, seed: 9));

        await UiTest.RunAsync(async () =>
        {
            ComponentInfoWindow? window = null;
            var told = new List<string?>();

            window = new ComponentInfoWindow();
            window.SetComponent("U8", "U8", [new ComponentEntry { BoardLabel = "U8", Category = "IC" }], [], [], [], "PAL",
                Path.GetTempPath(), [], string.Empty, string.Empty);
            window.UseCollections(this.Context() with
            {
                // The settings read while the next image is picked - its banner cleared, as a picked thumbnail clears it.
                ReadScopeSettings = () =>
                {
                    window.ClearTemporaryCapturedOscilloscopeImageForTests();
                    return Task.FromResult<ScopeSettingsReading?>(new ScopeSettingsReading(50e-6, 2, 1));
                },
                CollectionChanged = told.Add
            });
            window.UpdateOscilloscopeSessionTitleState(true, true);

            window.CaptureOnScreenForTests(saved);
            await window.CollectCaptureAsync(new ComponentImageEntry { BoardLabel = "U8", Pin = "5" }, saved);

            Assert.Equal("5", Assert.Single(this.thisStore.Find(collection.LocalId)!.Captures).Pin);
            Assert.Null(window.CollectedTextForTests);
            Assert.False(window.FindControl<Button>("TakeOutOfCollectionButton")!.IsVisible);
            Assert.Equal([collection.LocalId], told);
        });
    }

    [Fact]
    public async Task With_no_collection_picked_a_capture_goes_nowhere_else()
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", null, null, DateTimeOffset.UtcNow)!;
        string saved = this.thisWorkspace.Path_("capture.png");
        File.WriteAllBytes(saved, CRT.Data.Tests.TestPng.Rgb(8, 8));

        await UiTest.RunAsync(async () =>
        {
            ComponentInfoWindow window = this.Popup();

            window.CaptureOnScreenForTests(saved);
            await window.CollectCaptureAsync(new ComponentImageEntry { BoardLabel = "U8", Pin = "1" }, saved);

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
