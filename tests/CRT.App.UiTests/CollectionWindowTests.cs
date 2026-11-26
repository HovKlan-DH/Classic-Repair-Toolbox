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

        return this.thisStore.AddCapture(made.LocalId, [1, 2, 3], new LocalCapture { ComponentLabel = "U7", Pin = "12", LooksWrong = true, CapturedUtc = DateTimeOffset.UtcNow })!;
    }

    private CollectionWindow Window(LocalCollection collection)
    {
        var client = new CollectionClient("https://classic-repair-toolbox.dk", new HttpClient(new AnsweringHttpHandler(request =>
        {
            this.thisRequests.Add((request.Method, request.RequestUri!.AbsolutePath));
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
        this.thisAnswer = _ => Json(HttpStatusCode.Gone, new CollectionRefusalAnswer(["That collection was deleted on 2026-November-1."]));

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
            Assert.StartsWith("You - ", row.Comments[2].Heading, StringComparison.Ordinal);
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
            Assert.Contains("hidden from the page", hidden.Heading, StringComparison.Ordinal);
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
}
