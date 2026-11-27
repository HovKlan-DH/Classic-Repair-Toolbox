using System.Net;
using System.Text;
using System.Text.Json;
using ClassicRepairToolbox.Tests.Maintainer;
using CRT;
using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A collection's window (CollectionWindow): edits kept on this computer, "Share" signing in first and keeping the page it
// made, a page the server no longer has forgotten, "Copy for forum", and "Delete collection" asking first - on a fake server.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class CollectionWindowTests : IDisposable
{
    private static readonly string Sha = new('a', 64);

    private readonly TempWorkspace thisWorkspace = new();

    private readonly CollectionStore thisStore;

    private readonly List<(HttpMethod Method, string Path)> thisRequests = [];

    private readonly List<string> thisCopied = [];

    private ReviewSession? thisSession;

    private int thisSignIns;

    private bool thisSignInSucceeds = true;

    private Func<HttpRequestMessage, HttpResponseMessage> thisAnswer = _ => new HttpResponseMessage(HttpStatusCode.InternalServerError);

    public CollectionWindowTests()
    {
        this.thisStore = new CollectionStore(this.thisWorkspace.Path_("collections"));
    }

    public void Dispose() => this.thisWorkspace.Dispose();

    private static ReviewSession Session() => new("token-1", DateTimeOffset.UtcNow.AddDays(30), 1, "anna@example.com", "Anna");

    private static HttpResponseMessage Json(HttpStatusCode status, object body) =>
        new(status) { Content = new StringContent(JsonSerializer.Serialize(body, ReviewApiContract.WireSettings), Encoding.UTF8, "application/json") };

    private LocalCollection Collection()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", "PAL", "No picture", DateTimeOffset.UtcNow)!;

        return this.thisStore.AddCapture(made.LocalId, CRT.Data.Tests.TestPng.Rgb(16, 10), new LocalCapture { ComponentLabel = "U7", Pin = "12", LooksWrong = true, CapturedUtc = DateTimeOffset.UtcNow }).Collection!;
    }

    // Held, a request that is no GET waits for it to be let go - an answer on its way while the test clicks again.
    private TaskCompletionSource? thisHold;

    // The keys first shares were sent under (CollectionRoutes.ClientKeyHeader).
    private readonly List<string> thisKeysSent = [];

    private CollectionWindow Window(LocalCollection collection)
    {
        var client = new CollectionClient("https://classic-repair-toolbox.dk", new HttpClient(new AnsweringHttpHandler(async (request, _) =>
        {
            this.thisRequests.Add((request.Method, request.RequestUri!.AbsolutePath));

            if (request.Headers.TryGetValues(CollectionRoutes.ClientKeyHeader, out IEnumerable<string>? keys))
                this.thisKeysSent.AddRange(keys);

            if (this.thisHold is { } hold && request.Method != HttpMethod.Get)
                await hold.Task;

            return this.thisAnswer(request);
        })));

        var host = new CollectionWindowHost(
            this.thisStore,
            client,
            SignedIn: () => this.thisSession,
            SignIn: () =>
            {
                this.thisSignIns++;

                if (this.thisSignInSucceeds)
                    this.thisSession = Session();

                return Task.CompletedTask;
            },
            OpenLink: _ => { },
            CopyText: text =>
            {
                this.thisCopied.Add(text);
                return Task.CompletedTask;
            },
            Changed: () => { });

        var window = new CollectionWindow();
        window.Initialize(host, collection.LocalId);

        return window;
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
            Assert.True(Assert.Single(window.RowsForTests).LooksWrong);
            Assert.StartsWith("Not shared yet", window.SharedStateText.Text, StringComparison.Ordinal);

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

    [Fact]
    public async Task Share_signs_in_first_and_keeps_the_page_it_made()
    {
        LocalCollection collection = this.Collection();
        this.thisAnswer = _ => Json(HttpStatusCode.Created, new CollectionAnswer("k7mq3x9vbn2p", "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddYears(1), 1));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();

            Assert.Equal(1, this.thisSignIns);
            Assert.Equal([(HttpMethod.Post, "/api/collections")], this.thisRequests);

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Equal("k7mq3x9vbn2p", kept.ServerId);
            Assert.Equal("https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p", kept.PageUrl);
            Assert.False(CollectionSharing.HasChangesNotShared(kept));
            Assert.Contains("https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p", window.StatusText.Text, StringComparison.Ordinal);
            Assert.False(window.ShareButton.IsEnabled);
            Assert.True(window.CopyForForumButton.IsEnabled);

            // Changed since: "Share" sends the change.
            window.DescriptionBox.Text = "It went black.";
            window.SaveEdits();

            Assert.True(window.ShareButton.IsEnabled);
            Assert.Contains("changed since", window.SharedStateText.Text, StringComparison.Ordinal);
        });
    }

    // Signing in asked for - and the account window closed without it.
    [Fact]
    public async Task Nobody_signing_in_shares_nothing()
    {
        LocalCollection collection = this.Collection();
        this.thisSignInSucceeds = false;

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();

            Assert.Equal(1, this.thisSignIns);
            Assert.Empty(this.thisRequests);
            Assert.False(this.thisStore.Find(collection.LocalId)!.IsShared);
            Assert.Equal(CollectionClient.SignInFirst, window.StatusText.Text);
        });
    }

    [Fact]
    public async Task A_page_the_server_no_longer_has_is_forgotten_so_Share_makes_a_new_one()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with
        {
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p",
            DeleteAtUtc = DateTimeOffset.UtcNow.AddYears(1),
            SharedContentHash = "old"
        });

        this.thisSession = Session();

        // The comments read as the window opens still there (a page gone then is forgotten at once - the test below); the share
        // finds it gone.
        this.thisAnswer = request => request.Method == HttpMethod.Get
            ? Json(HttpStatusCode.OK, new CollectionCommentsAnswer([]))
            : Json(HttpStatusCode.Gone, new CollectionRefusalAnswer(["That collection was deleted on 2026-November-1."]));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Null(kept.ServerId);
            Assert.Null(kept.PageUrl);
            Assert.Contains("\"Share\" makes a new page", window.StatusText.Text, StringComparison.Ordinal);
            Assert.True(window.ShareButton.IsEnabled);
        });
    }

    [Fact]
    public async Task Copy_for_forum_puts_the_link_and_what_looks_wrong_on_the_clipboard()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { ServerId = "k7mq3x9vbn2p", PageUrl = "https://x/collections/k7mq3x9vbn2p" });

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.CopyAsync(CollectionSharing.ForumFormat.BbCode);

            string copied = Assert.Single(this.thisCopied);

            Assert.StartsWith("[url=https://x/collections/k7mq3x9vbn2p]No picture[/url]", copied, StringComparison.Ordinal);
            Assert.Contains("#U7-12]", copied, StringComparison.Ordinal);
        });
    }

    [Fact]
    public async Task Delete_asks_first_then_takes_the_page_off_the_server_and_the_collection_off_this_computer()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { ServerId = "k7mq3x9vbn2p", PageUrl = "https://x/collections/k7mq3x9vbn2p" });
        this.thisSession = Session();
        this.thisAnswer = _ => new HttpResponseMessage(HttpStatusCode.NoContent);

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            // What the window read as it opened (its comments) is no part of the delete.
            this.thisRequests.Clear();

            window.DeleteButton.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Avalonia.Controls.Button.ClickEvent));

            Assert.True(window.DeleteQuestion.IsVisible);
            Assert.Contains("and its page from the server?", window.DeleteQuestionText.Text, StringComparison.Ordinal);
            Assert.Empty(this.thisRequests);

            await window.DeleteAsync();

            Assert.Equal([(HttpMethod.Delete, "/api/collections/k7mq3x9vbn2p")], this.thisRequests);
            Assert.Null(this.thisStore.Find(collection.LocalId));
        });
    }

    // ---- Comments (step 3) ---------------------------------------------------------------------

    private LocalCollection Shared(long seenCommentId = 0)
    {
        LocalCollection collection = this.Collection();
        LocalCollection shared = this.thisStore.Find(collection.LocalId)! with
        {
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://x/collections/k7mq3x9vbn2p",
            SeenCommentId = seenCommentId
        };

        this.thisStore.Save(shared);
        return shared;
    }

    private static CollectionCommentAnswer Comment(long id, string sha, bool mine = false, bool hidden = false) =>
        new(id, sha, mine ? "Anna" : "Bo", ByOwner: mine, ByAdministrator: false, Mine: mine, $"Comment {id}", DateTimeOffset.UtcNow, null, hidden);

    // The comments answered for every GET of them, and 204 for anything done to one.
    private void AnswerComments(params CollectionCommentAnswer[] comments) =>
        this.thisAnswer = request => request.Method == HttpMethod.Get
            ? Json(HttpStatusCode.OK, new CollectionCommentsAnswer(comments))
            : new HttpResponseMessage(HttpStatusCode.NoContent);

    [Fact]
    public async Task A_shared_collection_s_comments_are_read_and_those_not_read_before_marked_new()
    {
        LocalCollection collection = this.Shared(seenCommentId: 5);
        string sha = collection.Captures[0].Sha256;

        this.thisSession = Session();
        this.AnswerComments(Comment(5, sha), Comment(6, sha), Comment(7, sha, mine: true));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.LoadCommentsAsync();

            CollectionCaptureRow row = Assert.Single(window.RowsForTests);

            Assert.True(row.ShowsComments);
            Assert.Equal("3 comments - 1 new", row.CommentsHeading);
            Assert.Equal([false, true, false], row.Comments.Select(comment => comment.IsNew));
            Assert.Equal(("You", true), (row.Comments[2].Author, row.Comments[2].IsMine));
            Assert.Equal([true, true, false], row.Comments.Select(comment => comment.CanHide));
            Assert.Equal([false, false, true], row.Comments.Select(comment => comment.CanDelete));

            // Read now: the popups' "new comments" goes.
            Assert.Equal(7, this.thisStore.Find(collection.LocalId)!.SeenCommentId);
        });
    }

    [Fact]
    public async Task Hiding_a_comment_asks_the_server_and_reads_the_comments_again()
    {
        LocalCollection collection = this.Shared();
        string sha = collection.Captures[0].Sha256;

        this.thisSession = Session();
        this.AnswerComments(Comment(6, sha, hidden: true));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            this.thisRequests.Clear();

            await window.CommentActionAsync((client, serverId, token) => client.SetCommentHiddenAsync(serverId, 6, true, token));

            Assert.Equal(
                [(HttpMethod.Put, "/api/collections/k7mq3x9vbn2p/comments/6/hidden"), (HttpMethod.Get, "/api/collections/k7mq3x9vbn2p/comments")],
                this.thisRequests);

            CollectionCommentRow hidden = Assert.Single(Assert.Single(window.RowsForTests).Comments);

            Assert.True(hidden.CanShow);
            Assert.Contains("hidden from the page", hidden.Details, StringComparison.Ordinal);
        });
    }

    // Signed out, nothing is asked - the comments are read once CRT is signed in.
    [Fact]
    public async Task Signed_out_no_comments_are_read()
    {
        LocalCollection collection = this.Shared();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.LoadCommentsAsync();

            Assert.Empty(this.thisRequests);
            Assert.False(Assert.Single(window.RowsForTests).ShowsComments);
        });
    }

    // The page could not be taken off: the collection stays, so it can be tried again.
    [Fact]
    public async Task A_page_the_server_would_not_take_off_keeps_the_collection()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { ServerId = "k7mq3x9vbn2p" });
        this.thisSession = Session();
        this.thisAnswer = _ => new HttpResponseMessage(HttpStatusCode.InternalServerError);

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.DeleteAsync();

            Assert.NotNull(this.thisStore.Find(collection.LocalId));
            Assert.False(string.IsNullOrWhiteSpace(window.StatusText.Text));
        });
    }

    // ---- Code review, 2026-11-27 -------------------------------------------------------------------

    // The server answers "not found" for another account's collection: with another account signed in, nothing is sent -
    // the collection kept here, and its page left to the account that shared it.
    [Fact]
    public async Task With_another_account_signed_in_nothing_is_deleted_or_shared()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with
        {
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://x/collections/k7mq3x9vbn2p",
            SharedByAccountId = 2,
            SharedByName = "Bo"
        });

        this.thisSession = Session();
        this.thisAnswer = _ => new HttpResponseMessage(HttpStatusCode.NotFound);

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            window.DescriptionBox.Text = "Changed";
            window.SaveEdits();
            await window.ShareAsync();

            Assert.Empty(this.thisRequests);
            Assert.Equal("k7mq3x9vbn2p", this.thisStore.Find(collection.LocalId)?.ServerId);
            Assert.Contains("[Bo]", window.StatusText.Text, StringComparison.Ordinal);

            // "Delete collection" asks first whether to delete it from this computer only - nothing is sent either way.
            await window.DeleteAsync();

            Assert.Empty(this.thisRequests);
            Assert.True(window.DeleteQuestion.IsVisible);
            Assert.Contains("from this computer only", window.DeleteQuestionText.Text, StringComparison.Ordinal);
            Assert.NotNull(this.thisStore.Find(collection.LocalId));
        });
    }

    // Shared with an account that is no more - a data reset deleted it - the collection still goes from this computer, asked
    // first; its page is left to its date (code review, 2026-11-27: it could never be deleted at all).
    [Fact]
    public async Task Another_account_s_collection_is_deleted_from_this_computer_only_once_asked()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { ServerId = "k7mq3x9vbn2p", SharedByAccountId = 2, SharedByName = "Bo" });
        this.thisSession = Session();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.DeleteAsync();
            await window.DeleteAsync();

            Assert.Empty(this.thisRequests);
            Assert.Null(this.thisStore.Find(collection.LocalId));
        });
    }

    // Another account's collection - one gone in a data reset - is shared again as a NEW page once "Share" is pressed again:
    // its captures kept, the old page's id let go of (code review, 2026-11-27: it could never be shared again).
    [Fact]
    public async Task Another_account_s_collection_is_shared_as_a_new_page_once_asked()
    {
        LocalCollection collection = this.Collection();
        this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with
        {
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://x/collections/k7mq3x9vbn2p",
            SharedByAccountId = 2,
            SharedByName = "Bo"
        });

        this.thisSession = Session();
        this.thisAnswer = _ => Json(HttpStatusCode.Created, new CollectionAnswer("zzzzzzzzzzzz", "https://x/collections/zzzzzzzzzzzz", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddYears(1), 1));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();
            Assert.Empty(this.thisRequests);

            await window.ShareAsync();

            (HttpMethod Method, string Path) sent = Assert.Single(this.thisRequests);
            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Equal(HttpMethod.Post, sent.Method);
            Assert.Equal(("zzzzzzzzzzzz", 1L), (kept.ServerId, kept.SharedByAccountId));
            Assert.Single(kept.Captures);
        });
    }

    // What is copied names the captures by their places on the page - with a change not shared, it would name ones the page
    // does not have: the copy buttons wait for "Share" (code review, 2026-11-27).
    [Fact]
    public async Task The_copy_buttons_wait_for_a_change_to_be_shared()
    {
        LocalCollection shared = this.Shared();
        DateTimeOffset deleteAt = DateTimeOffset.UtcNow.AddYears(1);
        LocalCollection collection = shared with
        {
            DeleteAtUtc = deleteAt,
            SharedContentHash = CollectionSharing.ContentHash(CollectionSharing.RequestOf(shared with { DeleteAtUtc = deleteAt }, deleteAt))
        };

        this.thisStore.Save(collection);

        await UiTest.RunAsync(() =>
        {
            CollectionWindow window = this.Window(collection);

            Assert.True(window.CopyForForumButton.IsEnabled);

            window.DescriptionBox.Text = "Changed";
            window.SaveEdits();

            Assert.False(window.CopyForForumButton.IsEnabled);
            Assert.False(window.CopyForRedditButton.IsEnabled);
            return Task.CompletedTask;
        });
    }

    [Fact]
    public async Task Share_keeps_the_account_it_shared_with()
    {
        LocalCollection collection = this.Collection();
        this.thisSession = Session();
        this.thisAnswer = _ => Json(HttpStatusCode.Created, new CollectionAnswer("k7mq3x9vbn2p", "https://x/collections/k7mq3x9vbn2p", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddYears(1), 1));

        await UiTest.RunAsync(async () =>
        {
            await this.Window(collection).ShareAsync();

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Equal((1L, "Anna"), (kept.SharedByAccountId, kept.SharedByName));
        });
    }

    // A note kept while the share was on its way stays - only the server's part is written over the collection as it is now.
    [Fact]
    public async Task An_edit_kept_while_sharing_stays()
    {
        LocalCollection collection = this.Collection();
        this.thisSession = Session();
        this.thisAnswer = request =>
        {
            if (request.Method == HttpMethod.Post)
                this.thisStore.Save(this.thisStore.Find(collection.LocalId)! with { Title = "Edited meanwhile" });

            return Json(HttpStatusCode.Created, new CollectionAnswer("k7mq3x9vbn2p", "https://x/collections/k7mq3x9vbn2p", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddYears(1), 1));
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

    // "Answer" pressed again while the first is on its way sends nothing more.
    [Fact]
    public async Task An_answer_clicked_twice_is_sent_once()
    {
        LocalCollection collection = this.Shared();
        string sha = collection.Captures[0].Sha256;

        this.thisSession = Session();
        this.AnswerComments();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            await window.LoadCommentsAsync();
            this.thisRequests.Clear();
            this.thisHold = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            Task first = window.CommentActionAsync((client, serverId, token) => client.AddCommentAsync(serverId, sha, "Check R7.", token));
            Task second = window.CommentActionAsync((client, serverId, token) => client.AddCommentAsync(serverId, sha, "Check R7.", token));

            this.thisHold.SetResult();
            await Task.WhenAll(first, second);

            Assert.Single(this.thisRequests, request => request.Method == HttpMethod.Post);
        });
    }

    // Past the overlay's limit with no answer, the window stays busy until the request really ends - let go, "Answer" pressed
    // again posted the answer a second time (code review, 2026-11-27).
    [Fact]
    public async Task An_answer_past_the_wait_s_limit_keeps_the_window_busy_until_it_ends()
    {
        LocalCollection collection = this.Shared();
        string sha = collection.Captures[0].Sha256;

        this.thisSession = Session();
        this.AnswerComments();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            await window.LoadCommentsAsync();
            this.thisRequests.Clear();
            this.thisHold = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            // The overlay's clock runs out at once.
            BusyOverlay.For(window)!.LimitOverrideForTests = _ => Task.CompletedTask;

            Task first = window.CommentActionAsync((client, serverId, token) => client.AddCommentAsync(serverId, sha, "Check R7.", token));

            for (int i = 0; i < 200 && window.StatusText.Text != CollectionWindow.NoAnswerYet; i++)
                await Task.Delay(10);

            Assert.Equal(CollectionWindow.NoAnswerYet, window.StatusText.Text);

            // Still busy: a second press sends nothing.
            Assert.False(await window.CommentActionAsync((client, serverId, token) => client.AddCommentAsync(serverId, sha, "Check R7.", token)));

            this.thisHold.SetResult();
            await first;

            Assert.Single(this.thisRequests, request => request.Method == HttpMethod.Post);
        });
    }

    // An answer sent leaves its box empty, so a second press sends nothing again (code review, 2026-11-27).
    [Fact]
    public async Task An_answer_sent_empties_its_box()
    {
        LocalCollection collection = this.Shared();
        string sha = collection.Captures[0].Sha256;

        this.thisSession = Session();
        this.AnswerComments();

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            await window.LoadCommentsAsync();
            this.thisRequests.Clear();

            CollectionCaptureRow row = Assert.Single(window.RowsForTests);
            row.Answer = "Check R7.";

            await window.AnswerAsync(new Avalonia.Controls.Button { Tag = sha });

            Assert.Null(row.Answer);
            Assert.Single(this.thisRequests, request => request.Method == HttpMethod.Post);

            // Refused, what was typed stays to be sent again.
            this.thisAnswer = request => request.Method == HttpMethod.Get
                ? Json(HttpStatusCode.OK, new CollectionCommentsAnswer([]))
                : Json(HttpStatusCode.BadRequest, new CollectionRefusalAnswer(["Write the comment first."]));
            row.Answer = "   ";

            await window.AnswerAsync(new Avalonia.Controls.Button { Tag = sha });

            Assert.Equal("   ", row.Answer);
        });
    }

    // Taken down on the website by the administrator, the page is forgotten as the comments are read - "Share" makes a new one
    // (code review, 2026-11-27: "There is no collection of yours with that id", and the window went on showing it shared).
    [Fact]
    public async Task A_page_taken_down_is_forgotten_as_its_comments_are_read()
    {
        LocalCollection collection = this.Shared();

        this.thisSession = Session();
        this.thisAnswer = _ => Json(HttpStatusCode.Gone, new CollectionRefusalAnswer(["That collection was deleted on 2026-November-27."]));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.LoadCommentsAsync();

            LocalCollection kept = this.thisStore.Find(collection.LocalId)!;

            Assert.Null(kept.ServerId);
            Assert.Equal(CollectionWindow.PageGone, window.StatusText.Text);
            Assert.True(window.ShareButton.IsEnabled);
            Assert.False(Assert.Single(window.RowsForTests).ShowsComments);
        });
    }

    // A comment gone meanwhile is that comment's 404 - the page stays known (code review, 2026-11-27: it was forgotten, and the
    // next "Share" made a second page).
    [Fact]
    public async Task A_comment_gone_meanwhile_leaves_the_page_known()
    {
        LocalCollection collection = this.Shared();

        this.thisSession = Session();
        this.thisAnswer = request => request.Method == HttpMethod.Get
            ? Json(HttpStatusCode.OK, new CollectionCommentsAnswer([]))
            : Json(HttpStatusCode.NotFound, new CollectionRefusalAnswer(["There is no comment of yours with that id."]));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            Assert.False(await window.CommentActionAsync((client, serverId, token) => client.SetCommentHiddenAsync(serverId, 6, true, token)));

            Assert.Equal("k7mq3x9vbn2p", this.thisStore.Find(collection.LocalId)?.ServerId);
            Assert.Equal("There is no comment of yours with that id.", window.StatusText.Text);
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
            firstRow.Answer = "An answer being typed";

            this.thisStore.AddCapture(collection.LocalId, CRT.Data.Tests.TestPng.Rgb(16, 10, seed: 5), new LocalCapture { ComponentLabel = "U8", Pin = "3", CapturedUtc = DateTimeOffset.UtcNow });
            window.ShowAgain();

            Assert.Equal(2, window.RowsForTests.Count);
            Assert.Equal("Typed before", window.TitleBox.Text);
            Assert.Equal("Typed before", this.thisStore.Find(collection.LocalId)!.Title);

            // The rows are changed in place: the capture there before keeps its very row - and so its box and the answer being
            // typed in it (code review, 2026-11-27: every row was made again at each capture, the caret lost).
            Assert.Same(firstRow, window.RowsForTests.Single(row => row.Sha256 == first));
            Assert.Equal("An answer being typed", firstRow.Answer);
        });
    }

    // A first share whose answer was lost made its page all the same: "Share" again sends it under the very same key, and the
    // server changes the page it made - never a second one only the administrator could take down (code review, 2026-11-27;
    // the server's half is CollectionFlowsTests).
    [Fact]
    public async Task A_first_share_sent_again_carries_the_same_key()
    {
        LocalCollection collection = this.Collection();
        this.thisSession = Session();

        int posts = 0;
        DateTimeOffset made = DateTimeOffset.UtcNow;

        this.thisAnswer = request =>
        {
            if (request.Method == HttpMethod.Post && posts++ == 0)
                throw new HttpRequestException("The answer was lost on the way.");

            return Json(HttpStatusCode.OK, new CollectionAnswer("k7mq3x9vbn2p", "https://x/collections/k7mq3x9vbn2p", made, made.AddYears(1), 1));
        };

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);

            await window.ShareAsync();

            Assert.Null(this.thisStore.Find(collection.LocalId)!.ServerId);

            await window.ShareAsync();

            Assert.Equal([(HttpMethod.Post, "/api/collections"), (HttpMethod.Post, "/api/collections")], this.thisRequests);
            Assert.Equal([collection.LocalId, collection.LocalId], this.thisKeysSent);
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
                SignedIn: () => null,
                SignIn: () => Task.CompletedTask,
                OpenLink: _ => { },
                CopyText: _ => Task.CompletedTask,
                Changed: () => { });

            Assert.False(new CollectionWindow().Initialize(host, collection.LocalId));
        });
    }

    // A click on "Looks wrong" is kept at once - Avalonia raises IsCheckedChanged before its binding writes the row, and the save
    // read the mark as it was (code review, 2026-11-27: "Share" stayed greyed out on a shared collection).
    [Fact]
    public void Looks_wrong_clicked_is_kept_at_once()
    {
        LocalCollection collection = this.Collection();

        UiTest.Run(() =>
        {
            CollectionWindow window = this.Window(collection);
            window.Show();
            window.UpdateLayout();

            Avalonia.Controls.CheckBox mark = Avalonia.VisualTree.VisualExtensions.GetVisualDescendants(window)
                .OfType<Avalonia.Controls.CheckBox>()
                .Single(box => Equals(box.Content, "Looks wrong"));

            Assert.True(mark.IsChecked);

            mark.IsChecked = false;

            Assert.False(this.thisStore.Find(collection.LocalId)!.Captures[0].LooksWrong);
            window.Close();
        });
    }

    // The share reads each capture the server asks for as it goes - one taken out meanwhile stopped it halfway (code review,
    // 2026-11-27).
    [Fact]
    public async Task Nothing_is_taken_out_while_a_share_is_on_its_way()
    {
        LocalCollection collection = this.Collection();
        string sha = collection.Captures[0].Sha256;

        this.thisSession = Session();
        this.thisAnswer = _ => Json(HttpStatusCode.Created, new CollectionAnswer("k7mq3x9vbn2p", "https://x/collections/k7mq3x9vbn2p", DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddYears(1), 1));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            this.thisHold = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            Task sharing = window.ShareAsync();

            window.TakeOut(sha);

            Assert.Equal(CollectionWindow.WaitForTheShare, window.StatusText.Text);
            Assert.False(window.LifetimeComboBox.IsEnabled);

            this.thisHold.SetResult();
            await sharing;

            Assert.Single(this.thisStore.Find(collection.LocalId)!.Captures);
            Assert.True(window.LifetimeComboBox.IsEnabled);
        });
    }

    // Comments read again change the rows in place - a note being typed keeps its text and the keyboard.
    [Fact]
    public async Task Comments_read_again_keep_the_rows_shown()
    {
        LocalCollection collection = this.Shared();
        string sha = collection.Captures[0].Sha256;

        this.thisSession = Session();
        this.AnswerComments(Comment(6, sha));

        await UiTest.RunAsync(async () =>
        {
            CollectionWindow window = this.Window(collection);
            CollectionCaptureRow before = Assert.Single(window.RowsForTests);
            before.Caption = "Being typed";
            window.TitleBox.Text = "Typed, not left";

            await window.LoadCommentsAsync();

            Assert.Same(before, Assert.Single(window.RowsForTests));
            Assert.Equal("Being typed", before.Caption);
            Assert.Equal("Typed, not left", window.TitleBox.Text);
            Assert.Single(before.Comments);
        });
    }
}
