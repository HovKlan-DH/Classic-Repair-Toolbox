using System.Net;
using System.Text;
using System.Text.Json;
using Avalonia.Controls;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A collection's window (CollectionWindow): edits kept on this computer, "Share" making the collection's KEY and keeping the
// page it made (no account - owner decision, 2026-11-28: "Anybody, no account"), "Update share", a page the server no longer
// has forgotten, "Copy link", and "Delete collection" asking first - on a fake server.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class CollectionWindowTests : IDisposable
{
    private static readonly string Key = new('c', CollectionRules.OwnerKeyLength);

    private readonly TempWorkspace thisWorkspace = new();

    private readonly CollectionStore thisStore;

    private readonly List<(HttpMethod Method, string Path)> thisRequests = [];

    private readonly List<string> thisCopied = [];

    private Func<HttpRequestMessage, HttpResponseMessage> thisAnswer = _ => new HttpResponseMessage(HttpStatusCode.InternalServerError);

    public CollectionWindowTests()
    {
        this.thisStore = new CollectionStore(this.thisWorkspace.Path_("collections"));
    }

    public void Dispose() => this.thisWorkspace.Dispose();

    private static HttpResponseMessage Json(HttpStatusCode status, object body) =>
        new(status) { Content = new StringContent(JsonSerializer.Serialize(body, ReviewApiContract.WireSettings), Encoding.UTF8, "application/json") };

    private static CollectionAnswer Page(string id = "k7mq3x9vbn2p", DateTimeOffset? deleteAt = null) =>
        new(id, $"https://x/collections/{id}", DateTimeOffset.UtcNow, deleteAt ?? DateTimeOffset.UtcNow.AddYears(1), 1);

    private LocalCollection Collection()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", "PAL", "No picture", DateTimeOffset.UtcNow)!;

        return this.thisStore.AddCapture(made.LocalId, CRT.Data.Tests.TestPng.Rgb(16, 10), new LocalCapture { ComponentLabel = "U7", Pin = "12", CapturedUtc = DateTimeOffset.UtcNow }).Collection!;
    }

    // Held, a request that is no GET waits for it to be let go - an answer on its way while the test clicks again.
    private TaskCompletionSource? thisHold;

    // The keys requests were sent under (CollectionRoutes.OwnerKeyHeader), one per request.
    private readonly List<string?> thisKeysSent = [];

    // The bodies of the requests sent, as JSON.
    private readonly List<string> thisBodies = [];

    private CollectionWindow Window(
        LocalCollection collection,
        Func<bool>? worklogIsOn = null,
        Func<LocalCollection, Avalonia.Controls.Window, Task<string?>>? attach = null,
        Func<WindowPlace?>? place = null,
        Action<WindowPlace>? rememberPlace = null,
        Func<CollectionLinkFormat>? linkFormat = null,
        Action<CollectionLinkFormat>? rememberLinkFormat = null)
    {
        var client = new CollectionClient("https://classic-repair-toolbox.dk", new HttpClient(new AnsweringHttpHandler(async (request, _) =>
        {
            this.thisRequests.Add((request.Method, request.RequestUri!.AbsolutePath));
            this.thisKeysSent.Add(request.Headers.TryGetValues(CollectionRoutes.OwnerKeyHeader, out IEnumerable<string>? keys) ? keys.Single() : null);

            if (request.Content is { } content)
                this.thisBodies.Add(await content.ReadAsStringAsync());

            if (this.thisHold is { } hold && request.Method != HttpMethod.Get)
                await hold.Task;

            return this.thisAnswer(request);
        })));

        var host = new CollectionWindowHost(
            this.thisStore,
            client,
            OpenLink: _ => { },
            CopyText: text =>
            {
                this.thisCopied.Add(text);
                return Task.CompletedTask;
            },
            Changed: () => { },
            WorklogIsOn: worklogIsOn,
            AttachToWorklog: attach,
            Place: place,
            RememberPlace: rememberPlace,
            LinkFormat: linkFormat,
            RememberLinkFormat: rememberLinkFormat);

        var window = new CollectionWindow();
        window.Initialize(host, collection.LocalId);

        return window;
    }

    // A collection shared under its key, as shared - its page up, nothing changed since.
    private LocalCollection SharedCollection()
    {
        LocalCollection collection = this.Collection();
        DateTimeOffset deleteAt = DateTimeOffset.UtcNow.AddYears(1);
        LocalCollection kept = this.thisStore.Find(collection.LocalId)! with
        {
            OwnerKey = Key,
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p",
            DeleteAtUtc = deleteAt,
            SharedUtc = DateTimeOffset.UtcNow
        };

        this.thisStore.Save(kept with { SharedContentHash = CollectionSharing.ContentHash(CollectionSharing.RequestOf(kept, deleteAt)) });

        return this.thisStore.Find(collection.LocalId)!;
    }

    // A collection shared WITH AN ACCOUNT before 2026-11-28: a page on the server, no key here.
    private LocalCollection SharedWithAnAccountBefore()
    {
        LocalCollection collection = this.Collection();

        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with
        {
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://x/collections/k7mq3x9vbn2p",
            DeleteAtUtc = DateTimeOffset.UtcNow.AddYears(1),
            SharedContentHash = "old"
        });

        return this.thisStore.Find(collection.LocalId)!;
    }

    [Fact]
    public void The_window_shows_the_collection_and_keeps_an_edit_on_this_computer()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(collection);

            Assert.Equal("No picture", window.TitleBox.Text);
            Assert.Equal("Commodore 64 - 250407 (PAL)", window.BoardText.Text);
            Assert.Equal("1 year", window.LifetimeComboBox.SelectedItem);
            Assert.Single(window.RowsForTests);
            Assert.Equal(CollectionWindow.NotSharedYet, window.SharedStateText.Text);
            Assert.Equal("Share", window.ShareButton.Content);
            Assert.False(window.StopSharingButton.IsVisible);

            window.TitleBox.Text = "Found the bad trace";
            window.ForumLinksBox.Text = "https://www.lemon64.com/forum/viewtopic.php?t=1\n\n";
            window.RowsForTests[0].Caption = "Pin 12 stays low";

            Assert.True(window.SaveEdits());

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Equal("Found the bad trace", kept.Title);
            Assert.Equal(["https://www.lemon64.com/forum/viewtopic.php?t=1"], kept.ForumLinks);
            Assert.Equal("Pin 12 stays low", kept.Captures[0].Caption);
        });
    }

    // ###########################################################################################
    // *** "SHARE" NEEDS NO ACCOUNT: THE COLLECTION'S KEY IS MADE AND KEPT HERE (owner decision, 2026-11-28). *** Sent with the
    // share, the page made kept - and the window says the page is this computer's alone (owner request, the same day: "It must
    // also be clear that the collection is then tied to the computer").
    // ###########################################################################################
    [Fact]
    public async Task Share_makes_the_collection_s_key_keeps_it_here_and_keeps_the_page_it_made()
    {
        LocalCollection collection = this.Collection();
        this.thisAnswer = _ => Json(HttpStatusCode.Created, Page());

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();

            Assert.Equal([(HttpMethod.Post, "/api/collections")], this.thisRequests);

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.True(CollectionRules.IsOwnerKey(kept.OwnerKey));
            Assert.Equal([kept.OwnerKey], this.thisKeysSent);
            Assert.Equal("k7mq3x9vbn2p", kept.ServerId);
            Assert.Equal("https://x/collections/k7mq3x9vbn2p", kept.PageUrl);
            Assert.False(CollectionSharing.HasChangesNotShared(kept));
            Assert.Equal(CollectionWindow.FirstSharedLine("https://x/collections/k7mq3x9vbn2p"), window.StatusText.Text);
            Assert.Contains("Only this computer can change or delete the page.", window.StatusText.Text, StringComparison.Ordinal);
            Assert.True(window.CopyLinkButton.IsEnabled);
            Assert.True(window.OpenPageButton.IsEnabled);

            // Owner request, 2026-11-28: "rename the "Share" to "Update share", if it is already shared" - offered with nothing
            // changed too, as every share counts its time again, the date bold.
            Assert.Equal("Update share", window.ShareButton.Content);
            Assert.True(window.ShareButton.IsEnabled);
            string deleteAt = CollectionRules.DateText(kept.DeleteAtUtc!.Value);
            Assert.Equal(
                $"Its page is deleted on {deleteAt}. Clicking \"Update share\" will update the deletion day.",
                BoldRuns.TextOf(window.DeleteDateText));
            Assert.Contains(window.DeleteDateText.Inlines!, inline => inline is Avalonia.Controls.Documents.Run { Text: var text, FontWeight: Avalonia.Media.FontWeight.Bold } && text == deleteAt);
            Assert.Equal(CollectionWindow.ThisComputerOnly, window.SharedStateText.Text);

            // Changed since: "Update share" sends the change.
            window.DescriptionBox.Text = "It went black.";
            window.SaveEdits();

            Assert.True(window.ShareButton.IsEnabled);
            Assert.Equal($"Changed since it was shared - \"Update share\" sends the change. {CollectionWindow.ThisComputerOnly}", window.SharedStateText.Text);
        });
    }

    // Shared again, under the same key, to its own page.
    [Fact]
    public async Task Update_share_sends_the_collection_to_its_page_under_its_key()
    {
        LocalCollection collection = this.SharedCollection();
        this.thisAnswer = _ => Json(HttpStatusCode.OK, Page());

        await UiTest.RunAsync(async () =>
        {
            await this.Window(collection).ShareAsync();

            Assert.Equal([(HttpMethod.Put, "/api/collections/k7mq3x9vbn2p")], this.thisRequests);
            Assert.Equal([Key], this.thisKeysSent);
            Assert.Equal(Key, this.thisStore.Find(collection.LocalId)!.OwnerKey);
        });
    }

    // "Share" never shared, "Update share" while the page is up, "Share" again once it is offline - which brings it back.
    [Fact]
    public void The_share_button_reads_Update_share_while_the_page_is_up()
    {
        LocalCollection shared = this.SharedCollection();

        Assert.Equal("Share", CollectionWindow.ShareLabel(shared with { ServerId = null }));
        Assert.Equal("Update share", CollectionWindow.ShareLabel(shared));
        Assert.Equal("Share", CollectionWindow.ShareLabel(shared with { SharingStoppedUtc = DateTimeOffset.UtcNow }));
        Assert.Equal("Share", CollectionWindow.ShareLabel(shared with { OwnerKey = null }));
    }

    // Owner request, 2026-11-28: "Rename "Open web page" to "Open shared page", but only if it is shared - if not shared, this button
    // should be disabled" - its page offline ("Stop sharing") is not shared either.
    [Fact]
    public void Open_shared_page_is_offered_only_while_it_is_shared()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            Assert.False(this.Window(collection).OpenPageButton.IsEnabled);

            this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with
            {
                OwnerKey = Key,
                ServerId = "k7mq3x9vbn2p",
                PageUrl = "https://x/collections/k7mq3x9vbn2p",
                SharingStoppedUtc = DateTimeOffset.UtcNow
            });

            Assert.False(this.Window(collection).OpenPageButton.IsEnabled);

            this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { SharingStoppedUtc = null });

            Assert.True(this.Window(collection).OpenPageButton.IsEnabled);
        });
    }

    // Owner request, 2026-11-28: "Rename "Take out" button to "Delete image"". A collection never shared asks the server nothing.
    [Fact]
    public async Task Delete_image_on_a_collection_never_shared_deletes_it_here_and_asks_the_server_nothing()
    {
        LocalCollection collection = this.Collection();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.TakeOutAsync(collection.Captures[0].Sha256);

            Assert.Empty(this.thisStore.Find(collection.LocalId)!.Captures);
            Assert.Equal(CollectionWindow.ImageDeleted, window.StatusText.Text);
            Assert.Empty(this.thisRequests);
        });
    }

    // A shared collection of two captures, as shared - its second capture's checksum with it.
    private (LocalCollection Collection, string Second) SharedWithTwo()
    {
        LocalCollection shared = this.SharedCollection();
        LocalCollection two = this.thisStore.AddCapture(
            shared.LocalId,
            CRT.Data.Tests.TestPng.Rgb(16, 10, 2),
            new LocalCapture { ComponentLabel = "U7", Pin = "13", CapturedUtc = DateTimeOffset.UtcNow }).Collection!;

        this.thisStore.Save(two with { SharedContentHash = CollectionSharing.ContentHash(CollectionSharing.RequestOf(two, two.DeleteAtUtc!.Value)) });

        return (this.thisStore.Find(shared.LocalId)!, two.Captures.Single(capture => capture.Pin == "13").Sha256);
    }

    // Owner request, 2026-11-28: "The "Delete image" should also remove it right away from the shared page, and do not touch the
    // expiry date. This date should only change when the USER does something" - that capture alone off the server under the key,
    // then here; the page followed, so nothing is said to be "changed since it was shared".
    [Fact]
    public async Task Delete_image_on_a_shared_collection_takes_it_off_its_page_at_once_and_keeps_its_date()
    {
        (LocalCollection collection, string second) = this.SharedWithTwo();
        DateTimeOffset deleteAt = collection.DeleteAtUtc!.Value;

        this.thisAnswer = _ => Json(HttpStatusCode.OK, Page(deleteAt: deleteAt));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.TakeOutAsync(second);

            Assert.Equal([(HttpMethod.Delete, $"/api/collections/k7mq3x9vbn2p/images/{second}")], this.thisRequests);
            Assert.Equal([Key], this.thisKeysSent);

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Single(kept.Captures);
            Assert.Equal(deleteAt, kept.DeleteAtUtc);
            Assert.False(CollectionSharing.HasChangesNotShared(kept));
            Assert.Equal(CollectionWindow.ImageDeletedShared, window.StatusText.Text);
            Assert.DoesNotContain("Changed since", window.SharedStateText.Text ?? string.Empty, StringComparison.Ordinal);
        });
    }

    // A change typed and not shared yet stays one - "Update share" still sends it.
    [Fact]
    public async Task Delete_image_leaves_a_change_not_shared_yet_to_Share()
    {
        (LocalCollection collection, string second) = this.SharedWithTwo();
        this.thisStore.Save(collection with { Title = "A title not shared yet" });

        this.thisAnswer = _ => Json(HttpStatusCode.OK, Page(deleteAt: collection.DeleteAtUtc!.Value));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.TakeOutAsync(second);

            Assert.True(CollectionSharing.HasChangesNotShared(this.thisStore.Find(collection.LocalId)!));
        });
    }

    // The server could not take it off: kept here too, what it said said, so it can be tried again.
    [Fact]
    public async Task Delete_image_the_server_refuses_keeps_the_capture_here()
    {
        (LocalCollection collection, string second) = this.SharedWithTwo();

        this.thisAnswer = _ => Json(HttpStatusCode.ServiceUnavailable, new CollectionRefusalAnswer(["Sharing collections is not set up on this server."]));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.TakeOutAsync(second);

            Assert.Equal(2, this.thisStore.Find(collection.LocalId)!.Captures.Count);
            Assert.Equal("Sharing collections is not set up on this server.", window.StatusText.Text);
        });
    }

    // Shared with an account before 2026-11-28, no key here owns the old page - a capture is deleted here alone, nothing sent.
    [Fact]
    public async Task Delete_image_on_a_collection_shared_with_an_account_before_deletes_it_here_only()
    {
        LocalCollection collection = this.SharedWithAnAccountBefore();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.TakeOutAsync(collection.Captures[0].Sha256);

            Assert.Empty(this.thisRequests);
            Assert.Empty(this.thisStore.Find(collection.LocalId)!.Captures);
        });
    }

    // Owner request, 2026-11-28: "ESC should also be able to close the window" - a delete being asked, Esc keeps the collection
    // first, as "Keep it" does.
    [Fact]
    public void Escape_keeps_a_collection_being_deleted_then_closes_the_window()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(collection);
            window.Show();

            try
            {
                window.AskToDelete();
                Assert.True(window.DeleteQuestion.IsVisible);

                window.RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent = Avalonia.Input.InputElement.KeyDownEvent, Key = Avalonia.Input.Key.Escape });

                Assert.False(window.DeleteQuestion.IsVisible);
                Assert.True(window.IsVisible);
                Assert.NotNull(this.thisStore.Find(collection.LocalId));

                window.RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent = Avalonia.Input.InputElement.KeyDownEvent, Key = Avalonia.Input.Key.Escape });

                Assert.False(window.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Owner request, 2026-11-28: "do save its size and position, so it will open same size and place the next time the user opens
    // it" - the place kept used as it opens, and what it was left at handed back as it closes.
    [Fact]
    public void The_window_opens_where_it_was_left_and_remembers_where_it_is_left()
    {
        LocalCollection collection = this.Collection();
        WindowPlace? remembered = null;

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(
                collection,
                place: () => new WindowPlace("Normal", 1000, 700, 60, 70),
                rememberPlace: kept => remembered = kept);

            Assert.Equal(1000, window.Width);
            Assert.Equal(700, window.Height);
            Assert.Equal(Avalonia.Controls.WindowStartupLocation.Manual, window.WindowStartupLocation);
            Assert.Equal(new Avalonia.PixelPoint(60, 70), window.Position);

            window.Show();
            window.Close();

            Assert.NotNull(remembered);
            Assert.Equal("Normal", remembered.State);
            Assert.Equal(new Avalonia.PixelPoint(60, 70), new Avalonia.PixelPoint(remembered.X, remembered.Y));
        });
    }

    [Fact]
    public async Task A_page_the_server_no_longer_has_is_forgotten_so_Share_makes_a_new_one()
    {
        LocalCollection collection = this.SharedCollection();
        this.thisAnswer = _ => Json(HttpStatusCode.Gone, new CollectionRefusalAnswer(["That collection was deleted on 2026-November-1."]));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            window.DescriptionBox.Text = "Changed";
            window.SaveEdits();
            await window.ShareAsync();

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Null(kept.ServerId);
            Assert.Null(kept.PageUrl);
            Assert.Contains("\"Share\" makes a new page", window.StatusText.Text, StringComparison.Ordinal);
            Assert.True(window.ShareButton.IsEnabled);
            Assert.Equal("Share", window.ShareButton.Content);
        });
    }

    // ###########################################################################################
    // *** A COLLECTION SHARED WITH AN ACCOUNT BEFORE 2026-11-28 IS SHARED AS A NEW PAGE, UNDER A KEY OF ITS OWN. *** No key of
    // this computer owns the old page, which goes at its date: the window says the collection is here only, "Share" makes its
    // key and a new page - its captures kept.
    // ###########################################################################################
    [Fact]
    public async Task A_collection_shared_with_an_account_before_is_shared_as_a_new_page_under_its_own_key()
    {
        LocalCollection collection = this.SharedWithAnAccountBefore();
        this.thisAnswer = _ => Json(HttpStatusCode.Created, Page("zzzzzzzzzzzz"));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            Assert.Equal(CollectionWindow.NotSharedYet, window.SharedStateText.Text);
            Assert.Equal("Share", window.ShareButton.Content);
            Assert.False(window.OpenPageButton.IsEnabled);

            await window.ShareAsync();

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Equal([(HttpMethod.Post, "/api/collections")], this.thisRequests);
            Assert.True(CollectionRules.IsOwnerKey(kept.OwnerKey));
            Assert.Equal([kept.OwnerKey], this.thisKeysSent);
            Assert.Equal("zzzzzzzzzzzz", kept.ServerId);
            Assert.Single(kept.Captures);
        });
    }

    // ... and deleted from this computer alone, asked first - the old page is left to its date, nothing sent.
    [Fact]
    public async Task A_collection_shared_with_an_account_before_is_deleted_from_this_computer_only()
    {
        LocalCollection collection = this.SharedWithAnAccountBefore();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            window.AskToDelete();
            Assert.Equal("Delete [No picture] from this computer?", window.DeleteQuestionText.Text);

            await window.DeleteAsync();

            Assert.Empty(this.thisRequests);
            Assert.Null(this.thisStore.Find(collection.LocalId));
        });
    }

    // Owner decision, 2026-11-28 ("OK to your proposal for the forum button"): "Copy link" copies in the format picked last - BBCode
    // the first time - and names it.
    [Fact]
    public async Task Copy_link_copies_in_the_format_picked_last_and_names_it()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { OwnerKey = Key, ServerId = "k7mq3x9vbn2p", PageUrl = "https://x/collections/k7mq3x9vbn2p" });

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow first = this.Window(collection);

            Assert.Equal("Copy link (BBCode)", first.CopyLinkButton.Content);

            await first.CopyAsync(first.LinkFormatShown);

            Assert.StartsWith("[url=https://x/collections/k7mq3x9vbn2p]No picture[/url]", Assert.Single(this.thisCopied), StringComparison.Ordinal);
            Assert.Equal(CollectionWindow.CopiedLine(CollectionLinkFormat.BbCode), first.StatusText.Text);

            CollectionWindow remembered = this.Window(collection, linkFormat: () => CollectionLinkFormat.Markdown);

            Assert.Equal("Copy link (Markdown)", remembered.CopyLinkButton.Content);
            Assert.Equal(CollectionLinkFormat.Markdown, remembered.LinkFormatShown);
        });
    }

    // The arrow lists the formats, each saying where it works, the one a click copies in checked; one picked there is copied and
    // remembered, so the next click repeats it.
    [Fact]
    public async Task A_format_picked_from_the_list_is_copied_and_remembered()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { OwnerKey = Key, ServerId = "k7mq3x9vbn2p", PageUrl = "https://x/collections/k7mq3x9vbn2p" });
        List<CollectionLinkFormat> remembered = [];

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection, rememberLinkFormat: remembered.Add);

            Assert.Equal(
                [
                    "BBCode - Lemon64, Forum64, AtariAge, VOGONS and most other forums",
                    "Markdown - Reddit, Discord, GitHub and Discourse",
                    "HTML - Web pages, blogs and HTML mail",
                    "Plain link - Facebook groups, email and anything else"
                ],
                window.LinkFormatItemsForTests.Select(item => BoldRuns.TextOf((TextBlock)item.Header!)));
            Assert.Equal([true, false, false, false], window.LinkFormatItemsForTests.Select(item => item.IsChecked));

            await window.CopyAsync(CollectionLinkFormat.Plain, remember: true);

            Assert.Equal([CollectionLinkFormat.Plain], remembered);
            Assert.EndsWith(": https://x/collections/k7mq3x9vbn2p", this.thisCopied.Last(), StringComparison.Ordinal);
            Assert.Equal("Copy link (plain)", window.CopyLinkButton.Content);
            Assert.Equal([false, false, false, true], window.LinkFormatItemsForTests.Select(item => item.IsChecked));

            await window.CopyAsync(window.LinkFormatShown);

            Assert.EndsWith(": https://x/collections/k7mq3x9vbn2p", this.thisCopied.Last(), StringComparison.Ordinal);
            Assert.Single(remembered);

            // Owner request, 2026-11-28: "add an option of "HTML" ... embed the link in "<a href>"".
            await window.CopyAsync(CollectionLinkFormat.Html, remember: true);

            Assert.StartsWith("<a href=\"https://x/collections/k7mq3x9vbn2p\">No picture</a>", this.thisCopied.Last(), StringComparison.Ordinal);
            Assert.Equal("Copy link (HTML)", window.CopyLinkButton.Content);
            Assert.Equal([CollectionLinkFormat.Plain, CollectionLinkFormat.Html], remembered);
        });
    }

    [Fact]
    public async Task Delete_asks_first_then_takes_the_page_off_the_server_under_its_key_and_the_collection_off_this_computer()
    {
        LocalCollection collection = this.SharedCollection();
        this.thisAnswer = _ => new HttpResponseMessage(HttpStatusCode.NoContent);

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            window.DeleteButton.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Avalonia.Controls.Button.ClickEvent));

            Assert.True(window.DeleteQuestion.IsVisible);
            Assert.Contains("and its page from the server?", window.DeleteQuestionText.Text, StringComparison.Ordinal);
            Assert.Empty(this.thisRequests);

            await window.DeleteAsync();

            Assert.Equal([(HttpMethod.Delete, "/api/collections/k7mq3x9vbn2p")], this.thisRequests);
            Assert.Equal([Key], this.thisKeysSent);
            Assert.Null(this.thisStore.Find(collection.LocalId));
        });
    }

    // "Delete collection" on the "Oscilloscope" tab's "Collections" list opens the window ASKING (Main.OpenCollectionWindowToDelete)
    // - the very question its own button asks, and nothing deleted or sent until it is answered.
    [Fact]
    public void Asked_to_delete_from_outside_the_window_asks_its_own_question()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(collection);

            window.AskToDelete();

            Assert.True(window.DeleteQuestion.IsVisible);
            Assert.False(window.ButtonsPanel.IsVisible);
            Assert.Equal("Delete [No picture] from this computer?", window.DeleteQuestionText.Text);
            Assert.NotNull(this.thisStore.Find(collection.LocalId));
            Assert.Empty(this.thisRequests);
        });
    }

    // The page could not be taken off: the collection stays, so it can be tried again.
    [Fact]
    public async Task A_page_the_server_would_not_take_off_keeps_the_collection()
    {
        LocalCollection collection = this.SharedCollection();
        this.thisAnswer = _ => new HttpResponseMessage(HttpStatusCode.InternalServerError);

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.DeleteAsync();

            Assert.NotNull(this.thisStore.Find(collection.LocalId));
            Assert.False(string.IsNullOrWhiteSpace(window.StatusText.Text));
        });
    }

    // What is copied names the captures by their places on the page - with a change not shared, it would name ones the page
    // does not have: the copy buttons wait for "Update share" (code review, 2026-11-27).
    [Fact]
    public async Task The_copy_buttons_wait_for_a_change_to_be_shared()
    {
        LocalCollection collection = this.SharedCollection();

        await UiTest.RunAsync(() =>
        {
            CollectionWindow window = this.Window(collection);

            Assert.True(window.CopyLinkButton.IsEnabled);

            window.DescriptionBox.Text = "Changed";
            window.SaveEdits();

            Assert.False(window.CopyLinkButton.IsEnabled);
            return Task.CompletedTask;
        });
    }

    // A note kept while the share was on its way stays - only the server's part is written over the collection as it is now.
    [Fact]
    public async Task An_edit_kept_while_sharing_stays()
    {
        LocalCollection collection = this.Collection();
        this.thisAnswer = request =>
        {
            if (request.Method == HttpMethod.Post)
                this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { Title = "Edited meanwhile" });

            return Json(HttpStatusCode.Created, Page());
        };

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Equal(("Edited meanwhile", "k7mq3x9vbn2p"), (kept.Title, kept.ServerId));
            Assert.True(CollectionSharing.HasChangesNotShared(kept));
        });
    }

    // Past the overlay's limit with no answer, the window stays busy until the request really ends - let go, a second press sent
    // the request again (code review, 2026-11-27).
    [Fact]
    public async Task Stop_sharing_past_the_wait_s_limit_keeps_the_window_busy_until_it_ends()
    {
        LocalCollection collection = this.SharedCollection();
        this.thisAnswer = _ => Json(HttpStatusCode.OK, Page());

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            this.thisHold = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            // The overlay's clock runs out at once.
            BusyOverlay.For(window)!.LimitOverrideForTests = _ => Task.CompletedTask;

            Task first = window.StopSharingAsync();

            for (int i = 0; i < 200 && window.StatusText.Text != CollectionWindow.NoAnswerYet; i++)
                await Task.Delay(10);

            Assert.Equal(CollectionWindow.NoAnswerYet, window.StatusText.Text);

            // Still busy: a second press sends nothing.
            await window.StopSharingAsync();

            this.thisHold.SetResult();
            await first;

            Assert.Single(this.thisRequests, request => request.Method == HttpMethod.Put);
        });
    }

    // A capture added in a component info window reaches the collection's window open on it (code review, 2026-11-27).
    [Fact]
    public void A_capture_added_elsewhere_is_listed_when_the_window_is_told()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(collection);
            window.TitleBox.Text = "Typed before";

            CollectionCaptureRow firstRow = window.RowsForTests[0];
            string first = firstRow.Sha256;
            firstRow.Caption = "A note being typed";

            this.thisStore.AddCapture(collection.LocalId, CRT.Data.Tests.TestPng.Rgb(16, 10, seed: 5), new LocalCapture { ComponentLabel = "U8", Pin = "3", CapturedUtc = DateTimeOffset.UtcNow });
            window.ShowAgain();

            Assert.Equal(2, window.RowsForTests.Count);
            Assert.Equal("Typed before", window.TitleBox.Text);
            Assert.Equal("Typed before", this.thisStore.Find(collection.LocalId)!.Title);

            // The rows are changed in place: the capture there before keeps its very row - and so its box and the note being
            // typed in it (code review, 2026-11-27: every row was made again at each capture, the caret lost).
            Assert.Same(firstRow, window.RowsForTests.Single(row => row.Sha256 == first));
            Assert.Equal("A note being typed", firstRow.Caption);
        });
    }

    // A first share whose answer was lost made its page all the same: "Share" again sends it under the very same key - kept here
    // BEFORE the first was sent - and the server changes the page it made, never a second one only the administrator could take
    // down (code review, 2026-11-27; the server's half is CollectionFlowsTests).
    [Fact]
    public async Task A_first_share_sent_again_carries_the_same_key()
    {
        LocalCollection collection = this.Collection();

        int posts = 0;

        this.thisAnswer = request =>
        {
            if (request.Method == HttpMethod.Post && posts++ == 0)
                throw new HttpRequestException("The answer was lost on the way.");

            return Json(HttpStatusCode.OK, Page());
        };

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();

            LocalCollection afterLost = this.thisStore.Find(collection.LocalId)!;

            Assert.Null(afterLost.ServerId);
            Assert.True(CollectionRules.IsOwnerKey(afterLost.OwnerKey));

            await window.ShareAsync();

            Assert.Equal([(HttpMethod.Post, "/api/collections"), (HttpMethod.Post, "/api/collections")], this.thisRequests);
            Assert.Equal([afterLost.OwnerKey, afterLost.OwnerKey], this.thisKeysSent);
            Assert.Equal("k7mq3x9vbn2p", this.thisStore.Find(collection.LocalId)!.ServerId);
        });
    }

    // A collection gone as its window is made is no window to show - Avalonia refuses to show a window closed (code review,
    // 2026-11-27).
    [Fact]
    public void A_window_for_a_collection_gone_is_not_made()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Delete(collection.LocalId);

        UiTest.Run(() =>
        {
            var host = new CollectionWindowHost(
                this.thisStore,
                new CollectionClient("https://classic-repair-toolbox.dk", new HttpClient(new AnsweringHttpHandler(this.thisAnswer))),
                OpenLink: _ => { },
                CopyText: _ => Task.CompletedTask,
                Changed: () => { });

            Assert.False(new CollectionWindow().Initialize(host, collection.LocalId));
        });
    }

    // Owner request, 2026-11-27: a thumbnail clicked opens "in another window and be zoomable" - the worklogs' viewer, gone at
    // a click outside it, its title where the capture was taken and its comment the settings.
    [Fact]
    public void A_capture_opened_from_its_thumbnail_is_shown_large_in_the_zooming_viewer()
    {
        LocalCollection collection = this.Collection();
        string sha = collection.Captures[0].Sha256;

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(collection);
            WorklogPhotoViewerWindow? shown = null;
            window.ShowPictureOverrideForTests = viewer => shown = viewer;

            window.OpenCapture(sha);

            Assert.NotNull(shown);
            Assert.Equal("U7, pin 12", shown.ShownTitleForTests);
            Assert.Equal(CollectionSharing.WorklogCommentOf(collection.Captures[0]), shown.CommentForTests);
            Assert.True(shown.ClosesWhenLeft);
        });
    }

    // Owner request, 2026-11-27: the note "a textarea that auto-expands" - several lines from the start, growing with its text.
    [Fact]
    public void A_capture_s_note_grows_with_what_is_typed()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(collection);
            window.Show();
            window.UpdateLayout();

            Avalonia.Controls.TextBox note = Avalonia.VisualTree.VisualExtensions.GetVisualDescendants(window)
                .OfType<Avalonia.Controls.TextBox>()
                .Single(box => box.PlaceholderText == "A note on this capture");

            double empty = note.Bounds.Height;

            note.Text = string.Join("\n", Enumerable.Range(1, 12).Select(line => $"Line {line}"));
            window.UpdateLayout();

            Assert.True(empty >= 64, $"The note starts {empty} high.");
            Assert.True(note.Bounds.Height > empty * 2, $"The note is {note.Bounds.Height} high with twelve lines, {empty} empty.");
            window.Close();
        });
    }

    // ---- "Stop sharing" (owner decision, 2026-11-27: "Offline, same link back") -------------------------------------------

    [Fact]
    public async Task Stop_sharing_takes_the_page_offline_and_Share_brings_it_back_under_the_same_link()
    {
        LocalCollection collection = this.SharedCollection();
        this.thisAnswer = _ => Json(HttpStatusCode.OK, Page());

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            Assert.True(window.StopSharingButton.IsVisible);

            await window.StopSharingAsync();

            Assert.Equal([(HttpMethod.Put, "/api/collections/k7mq3x9vbn2p/sharing")], this.thisRequests);
            Assert.Equal([Key], this.thisKeysSent);
            Assert.Contains("\"shared\":false", this.thisBodies.Last(), StringComparison.Ordinal);

            LocalCollection stopped = this.thisStore.Find(collection.LocalId)!;

            Assert.NotNull(stopped.SharingStoppedUtc);
            Assert.Equal("k7mq3x9vbn2p", stopped.ServerId);
            Assert.False(stopped.IsOnline);
            Assert.Equal(CollectionWindow.SharingStopped, window.StatusText.Text);
            Assert.StartsWith("Not shared at the moment", window.SharedStateText.Text, StringComparison.Ordinal);
            Assert.EndsWith(CollectionWindow.ThisComputerOnly, window.SharedStateText.Text, StringComparison.Ordinal);
            Assert.False(window.StopSharingButton.IsVisible);
            Assert.False(window.CopyLinkButton.IsEnabled);

            // Nothing changed, and still "Share" brings it back - the collection sent again under its id.
            Assert.Equal("Share", window.ShareButton.Content);
            Assert.True(window.ShareButton.IsEnabled);

            await window.ShareAsync();

            Assert.Equal((HttpMethod.Put, "/api/collections/k7mq3x9vbn2p"), this.thisRequests.Last());

            LocalCollection back = this.thisStore.Find(collection.LocalId)!;

            Assert.Null(back.SharingStoppedUtc);
            Assert.Equal("https://x/collections/k7mq3x9vbn2p", back.PageUrl);
            Assert.True(window.StopSharingButton.IsVisible);
            Assert.Equal("Update share", window.ShareButton.Content);
        });
    }

    // Its time up or taken down meanwhile: the page is forgotten, the collection kept here.
    [Fact]
    public async Task Stop_sharing_a_page_the_server_no_longer_has_forgets_the_page_and_keeps_the_collection()
    {
        LocalCollection collection = this.SharedCollection();
        this.thisAnswer = _ => Json(HttpStatusCode.Gone, new CollectionRefusalAnswer(["That collection was deleted on 2026-November-1."]));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.StopSharingAsync();

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Null(kept.ServerId);
            Assert.Null(kept.SharingStoppedUtc);
            Assert.Single(kept.Captures);
            Assert.Equal(CollectionWindow.PageGone, window.StatusText.Text);
        });
    }

    // ---- "Attach collection to worklog" (owner decision, 2026-11-27: "A copy of its captures") ----------------------------

    [Fact]
    public async Task Attach_to_worklog_hands_CRT_the_collection_as_kept_and_says_what_came_of_it()
    {
        LocalCollection collection = this.Collection();
        LocalCollection? handed = null;

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(
                collection,
                worklogIsOn: () => true,
                attach: (attached, _) =>
                {
                    handed = attached;
                    return Task.FromResult<string?>("1 capture attached to worklog #3 as photos.");
                });

            Assert.True(window.AttachToWorklogButton.IsVisible);

            // What the boxes hold is kept first.
            window.RowsForTests[0].Caption = "Stays low";

            await window.AttachToWorklogAsync();

            Assert.Equal("Stays low", handed?.Captures[0].Caption);
            Assert.Equal("1 capture attached to worklog #3 as photos.", window.StatusText.Text);
        });
    }

    // The worklog turned off, there is nothing to attach to - the button is not there, as the capture's own is not.
    [Fact]
    public void Attach_to_worklog_is_not_offered_with_the_worklog_turned_off()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            Assert.False(this.Window(collection, worklogIsOn: () => false, attach: (_, _) => Task.FromResult<string?>(null)).AttachToWorklogButton.IsVisible);
            Assert.False(this.Window(collection).AttachToWorklogButton.IsVisible);
        });
    }

    // The share reads each capture the server asks for as it goes - one taken out meanwhile stopped it halfway (code review,
    // 2026-11-27).
    [Fact]
    public async Task Nothing_is_taken_out_while_a_share_is_on_its_way()
    {
        LocalCollection collection = this.Collection();
        string sha = collection.Captures[0].Sha256;

        this.thisAnswer = _ => Json(HttpStatusCode.Created, Page());

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            this.thisHold = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            Task sharing = window.ShareAsync();

            await window.TakeOutAsync(sha);

            Assert.Equal(CollectionWindow.WaitForTheShare, window.StatusText.Text);
            Assert.False(window.LifetimeComboBox.IsEnabled);

            this.thisHold.SetResult();
            await sharing;

            Assert.Single(this.thisStore.Find(collection.LocalId)!.Captures);
            Assert.True(window.LifetimeComboBox.IsEnabled);
        });
    }
}
