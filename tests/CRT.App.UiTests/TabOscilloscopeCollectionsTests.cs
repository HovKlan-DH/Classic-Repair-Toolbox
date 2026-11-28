using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.Collections;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "Oscilloscope" tab's two tabs where the log stood alone (owner request, 2026-11-27): "Communication log" first, and
// "Collections" - every collection on this computer, its card's buttons asking the host for the collection's window. No
// oscilloscope anywhere.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabOscilloscopeCollectionsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private readonly CollectionStore thisStore;

    private readonly List<string> thisOpened = [];

    private readonly List<string> thisAskedToDelete = [];

    private readonly List<string> thisLinks = [];

    public TabOscilloscopeCollectionsTests()
    {
        this.thisStore = new CollectionStore(this.thisWorkspace.Path_("collections"));
    }

    public void Dispose() => this.thisWorkspace.Dispose();

    private LocalCollection Collection(string title, DateTimeOffset made)
    {
        LocalCollection collection = this.thisStore.Create("Commodore 64", "250407", "PAL", title, made)!;

        return this.thisStore.AddCapture(
            collection.LocalId,
            CRT.Data.Tests.TestPng.Rgb(16, 10),
            new LocalCapture { ComponentLabel = "U7", Pin = "12", CapturedUtc = made }).Collection!;
    }

    // Shared under its key, its page up.
    private LocalCollection Shared(LocalCollection collection)
    {
        LocalCollection shared = collection with
        {
            OwnerKey = new string('c', Handlers.DataHandling.CollectionRules.OwnerKeyLength),
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p",
            SharedUtc = collection.CreatedUtc,
            DeleteAtUtc = collection.CreatedUtc.AddYears(1)
        };

        Assert.True(this.thisStore.Save(shared));
        return shared;
    }

    private OscilloscopeCollectionsHost Host() =>
        new(
            () => this.thisStore,
            this.thisOpened.Add,
            this.thisAskedToDelete.Add,
            this.thisLinks.Add);

    private (Window Window, TabOscilloscope Tab) Shown()
    {
        var tab = new TabOscilloscope();
        tab.UseCollections(this.Host());

        var window = new Window { Content = tab, Width = 1200, Height = 800 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, tab);
    }

    private static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    // The buttons of the cards shown, by their words - the list made and laid out first.
    private static List<Button> CardButtons(Window window, string words)
    {
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();

        return window.GetVisualDescendants().OfType<Button>().Where(button => button.Content as string == words && button.IsEffectivelyVisible).ToList();
    }

    [Fact]
    public void The_communication_log_is_shown_first_and_Collections_shows_every_collection()
    {
        this.Collection("Older", DateTimeOffset.UtcNow.AddDays(-2));
        this.Collection("Newest", DateTimeOffset.UtcNow);

        UiTest.Run(() =>
        {
            (Window window, TabOscilloscope tab) = this.Shown();

            try
            {
                Assert.True(tab.OutputTextBox.IsVisible);
                Assert.False(tab.CollectionsPanel.IsVisible);
                Assert.Contains("Selected", tab.CommunicationLogTabButton.Classes);

                Click(tab.CollectionsTabButton);

                Assert.True(tab.CollectionsPanel.IsVisible);
                Assert.False(tab.OutputTextBox.IsVisible);
                Assert.Contains("Selected", tab.CollectionsTabButton.Classes);
                Assert.DoesNotContain("Selected", tab.CommunicationLogTabButton.Classes);

                Assert.Equal(["Newest", "Older"], tab.CollectionRowsForTests.Select(row => row.Item.Name));
                Assert.Equal("2 collections on this computer", tab.CollectionsHeadingText.Text);
                Assert.True(tab.CollectionsIntroText.IsVisible);
                Click(tab.CommunicationLogTabButton);

                Assert.True(tab.OutputTextBox.IsVisible);
                Assert.False(tab.CollectionsPanel.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public void With_no_collection_the_list_says_where_one_is_made()
    {
        UiTest.Run(() =>
        {
            (Window window, TabOscilloscope tab) = this.Shown();

            try
            {
                tab.ShowOutputTab(collections: true);

                Assert.Empty(tab.CollectionRowsForTests);
                Assert.Equal(CollectionListing.NoneYet, tab.CollectionsHeadingText.Text);
                Assert.False(tab.CollectionsIntroText.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // "Open collection" and "Delete collection" open the collection's window - asking, for the delete, as the window's own
    // delete is the one that takes the page down. "Open shared page" only on a card that is shared (owner request, 2026-11-28).
    [Fact]
    public void A_card_s_buttons_ask_for_its_collection_s_window_and_page()
    {
        LocalCollection kept = this.Collection("Kept here", DateTimeOffset.UtcNow.AddDays(-1));
        LocalCollection shared = this.Shared(this.Collection("Shared", DateTimeOffset.UtcNow));

        UiTest.Run(() =>
        {
            (Window window, TabOscilloscope tab) = this.Shown();

            try
            {
                tab.ShowOutputTab(collections: true);

                List<Button> show = CardButtons(window, "Open collection");
                List<Button> delete = CardButtons(window, "Delete collection");
                List<Button> page = CardButtons(window, "Open shared page");

                Assert.Equal(2, show.Count);
                Assert.Equal(2, delete.Count);
                Assert.Single(page);

                // The newest first: the shared one.
                Click(show[1]);
                Click(delete[0]);
                Click(page[0]);

                Assert.Equal([kept.LocalId], this.thisOpened);
                Assert.Equal([shared.LocalId], this.thisAskedToDelete);
                Assert.Equal([shared.PageUrl!], this.thisLinks);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A collection changed elsewhere - its window, a component info window - is shown as it is now, the cards of the others
    // kept as they were (a card made again under a press throws the click away).
    [Fact]
    public void A_change_shows_the_collection_as_it_is_now_and_keeps_the_other_cards()
    {
        LocalCollection first = this.Collection("First", DateTimeOffset.UtcNow.AddDays(-1));
        LocalCollection second = this.Collection("Second", DateTimeOffset.UtcNow);

        UiTest.Run(() =>
        {
            (Window window, TabOscilloscope tab) = this.Shown();

            try
            {
                tab.ShowOutputTab(collections: true);

                CollectionListRow firstRow = tab.CollectionRowsForTests.Single(row => row.Item.LocalId == first.LocalId);

                Assert.True(this.thisStore.Save(this.thisStore.Find(second.LocalId)! with { Title = "Renamed" }));
                tab.ShowCollections();

                Assert.Same(firstRow, tab.CollectionRowsForTests.Single(row => row.Item.LocalId == first.LocalId));
                Assert.Equal("Renamed", tab.CollectionRowsForTests.Single(row => row.Item.LocalId == second.LocalId).Item.Name);

                Assert.True(this.thisStore.Delete(first.LocalId));
                tab.ShowCollections();

                Assert.Equal(["Renamed"], tab.CollectionRowsForTests.Select(row => row.Item.Name));
                Assert.Equal("1 collection on this computer", tab.CollectionsHeadingText.Text);
            }
            finally
            {
                window.Close();
            }
        });
    }
}
