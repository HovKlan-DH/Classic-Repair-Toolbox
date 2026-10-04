using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.Interactivity;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// A submission's three views as the tab shows them (owner request, 2026-09-30: "The "Files..."
// button ... can easily be missed ... Maybe one row with "Contributor info", "Board data", and then
// "Files"?") - TabMaintainer.SubmissionViews.cs, .Files.cs and .Contributor.cs.
//
// No server: the tab is handed a submission and its detail, as the queue tests do. The Files
// view's tree needs one, so what is pinned here is its button's count and that choosing it shows
// it; the tree itself is FileTreeViewTests'.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerSubmissionViewsTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(long id = 42) =>
        new(id, "Commodore/C64/250407", "pending", $"Change {id}", "c@example.com", DateTimeOffset.UtcNow.AddHours(-1), false, false, true);

    private static ReviewSubmissionDetail Detail(
        ReviewQueueRow row,
        IReadOnlyList<SubmittedFileFact>? submittedFiles = null,
        FileRemovalPreview? removals = null,
        ReviewContributorFacts? contributor = null) =>
        new(
            row,
            true,
            new ReviewChangeSummaryView(false, []),
            [],
            new ReviewSubmissionAssets([]),
            [],
            new Dictionary<string, string>(),
            new Dictionary<string, string>(),
            submittedFiles ?? [],
            null,
            removals,
            null,
            contributor);

    private static SubmittedFileFact File(string name, string hash, string? published) =>
        new($"Commodore/C64/250407/{name}", hash, 10, SubmissionFileScope.Own, IsReferenced: true, PublishedSha256: published);

    private static void Select(TabMaintainer main, ReviewQueueRow? row) =>
        typeof(TabMaintainer).GetMethod("ShowSubmission", Any)!.Invoke(main, [row]);

    private static bool Shown(TabMaintainer main, string name) => main.FindControl<Control>(name)!.IsVisible;

    private static void Click(TabMaintainer main, string button) =>
        main.FindControl<Button>(button)!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    private static readonly ReviewContributorFacts Record = new(
        "dh@example.com",
        "Dennis",
        Published: 1,
        Waiting: 0,
        ChangesRequested: 0,
        Rejected: 1,
        SignedIn: false,
        AccountCreatedUtc: null,
        Submissions:
        [
            new ContributorSubmissionEntry(9, "Commodore/C128/310378", "Wrong board", "rejected", new DateTimeOffset(2026, 9, 20, 9, 0, 0, TimeSpan.Zero), null, "This is the 310378."),
            new ContributorSubmissionEntry(7, "Commodore/C64/250407", "Fixed U8", "merged", new DateTimeOffset(2026, 9, 10, 9, 0, 0, TimeSpan.Zero), null, null)
        ],
        PublishedToStable: 0);

    // ###########################################################################################
    // A submission opens on Board data - the table stays "the default first view" (owner request,
    // 2026-09-26) - with the other two views hidden and the old "Files..." button gone.
    // ###########################################################################################
    [Fact]
    public void A_submission_opens_on_board_data_with_three_view_buttons()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row();

            Select(main, row);
            main.ShowDetail(Detail(row));

            Assert.Equal(SubmissionView.BoardData, main.ShownSubmissionView);
            Assert.True(Shown(main, "BoardDataView"));
            Assert.False(Shown(main, "FilesView"));
            Assert.False(Shown(main, "ContributorView"));

            Assert.Contains("Selected", main.FindControl<Button>("BoardDataViewButton")!.Classes);
            Assert.Null(main.FindControl<Button>("SubmissionFilesButton"));
        });
    }

    // ###########################################################################################
    // *** THE FILES BUTTON SAYS HOW MANY FILES CHANGE, BEFORE ANYTHING IS OPENED. *** It replaced
    // the lines that said a file had been replaced under its own name - which colours no cell - so
    // the count is what keeps such a file from being approved unseen. No change, no badge.
    // ###########################################################################################
    [Fact]
    public void The_files_button_counts_the_files_the_submission_changes()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row();

            Select(main, row);
            main.ShowDetail(Detail(
                row,
                submittedFiles:
                [
                    File("manual.pdf", "new", published: "old"),
                    File("Sheet1.png", "aa", published: "aa"),
                    File("extra.png", "bb", published: null)
                ],
                removals: new FileRemovalPreview(["Commodore/C64/250407/gone.png"], null)));

            Assert.Equal("3", main.FilesCountForTests);

            // No tooltip on the view buttons, as on the screen buttons (owner request, 2026-09-30).
            Assert.Null(ToolTip.GetTip(main.FindControl<Button>("FilesViewButton")!));

            main.ShowDetail(Detail(row, submittedFiles: [File("Sheet1.png", "aa", published: "aa")]));
            Assert.Null(main.FilesCountForTests);
        });
    }

    // ###########################################################################################
    // Choosing a view shows it and hides the others - and switching HIDES, it never closes: back on
    // Board data, the table view is the same control it was.
    // ###########################################################################################
    [Fact]
    public void Choosing_a_view_shows_it_and_coming_back_finds_the_table_as_it_was()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row();

            Select(main, row);
            main.ShowDetail(Detail(row, contributor: Record));

            Control table = main.FindControl<Control>("SubmissionTable")!;

            Click(main, "ContributorViewButton");

            Assert.Equal(SubmissionView.Contributor, main.ShownSubmissionView);
            Assert.True(Shown(main, "ContributorView"));
            Assert.False(Shown(main, "BoardDataView"));
            Assert.Contains("Selected", main.FindControl<Button>("ContributorViewButton")!.Classes);
            Assert.DoesNotContain("Selected", main.FindControl<Button>("BoardDataViewButton")!.Classes);

            Click(main, "FilesViewButton");
            Assert.True(Shown(main, "FilesView"));
            Assert.False(Shown(main, "ContributorView"));

            Click(main, "BoardDataViewButton");
            Assert.True(Shown(main, "BoardDataView"));
            Assert.Same(table, main.FindControl<Control>("SubmissionTable"));
        });
    }

    // ###########################################################################################
    // THE FILES VIEW AGAINST A STUBBED SERVER (code review, 2026-10-01).
    // ###########################################################################################
    private static string FilesAnswer() =>
        JsonSerializer.Serialize(
            new SubmissionFilesAnswer(
                "Commodore/C64/250407",
                [new SystemFileEntry("Commodore/C64/250407/a.png", SystemFileChange.Added, SystemFileSource.Submission, "aa")]),
            ReviewApiContract.WireSettings);

    private static void UseServer(TabMaintainer main, AnsweringHttpHandler server)
    {
        typeof(TabMaintainer).GetField("thisClient", Any)!
            .SetValue(main, new ReviewApiClient("https://review.invalid", new HttpClient(server)));
        typeof(TabMaintainer).GetField("thisSession", Any)!
            .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));
    }

    private static async Task SettleAsync(Func<bool> done)
    {
        for (int attempt = 0; attempt < 500 && !done(); attempt++)
        {
            await Task.Delay(2);
            Dispatcher.UIThread.RunJobs();
        }

        Assert.True(done());
    }

    // *** THE WAIT NAMES THE SYSTEM FROM THE SUBMISSION'S OWN RECORD. *** It read the queue's selected
    // row, and a submission decided elsewhere stays open with no row in the list - so the overlay
    // said "Working out 's files...". Here, as in that case, the tab has no queue row at all.
    [Fact]
    public async Task The_files_wait_names_the_system_even_when_the_submission_has_no_queue_row()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            BusyOverlay overlay = MaintainerTabHost.AddOverlay(main);
            ReviewQueueRow row = Row();

            Select(main, row);
            main.ShowDetail(Detail(row));
            Assert.Null(main.SelectedQueueRowForTests);

            string? sentence = null;
            UseServer(main, new AnsweringHttpHandler(_ =>
            {
                sentence = overlay.Message;
                return AnsweringHttpHandler.Json(FilesAnswer());
            }));

            await main.ShowSubmissionViewAsync(SubmissionView.Files);

            Assert.Equal(MaintainerWaitWording.ReadingSubmissionFiles("Commodore/C64/250407"), sentence);
        });
    }

    // *** A TREE ON SCREEN IS READ AGAIN WHEN THE SUBMISSION CHANGES UNDER IT. *** The marker was
    // cleared but nothing drew the tree again while the maintainer stayed on the Files view - so a
    // list built before another maintainer's amendment stayed up to be approved from. A detail
    // shown again for the SAME object changes nothing (ReapplyApprovalGate); a new one re-reads.
    // ON SCREEN means in a shown window - see the off-screen case below.
    [Fact]
    public async Task A_files_tree_on_screen_is_read_again_when_the_submissions_detail_is_read_again()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            MaintainerTabHost.AddOverlay(main);
            var window = new Window { Content = main.Parent };
            window.Show();
            ReviewQueueRow row = Row();

            Select(main, row);
            ReviewSubmissionDetail first = Detail(row);
            main.ShowDetail(first);

            int requests = 0;
            UseServer(main, new AnsweringHttpHandler(_ =>
            {
                requests++;
                return AnsweringHttpHandler.Json(FilesAnswer());
            }));

            await main.ShowSubmissionViewAsync(SubmissionView.Files);
            Assert.Equal(1, requests);

            // The same detail again: nothing to read.
            main.ShowDetail(first);
            await Task.Delay(30);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(1, requests);

            // Another maintainer's amendment: a new detail, and the tree on screen is read again.
            main.ShowDetail(Detail(row));
            await SettleAsync(() => requests == 2);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** NOT WHILE THE TAB IS OFF SCREEN (code review, 2026-10-01). *** The badge's minute check
    // re-reads the detail while CRT shows another tab, and the read waits under MAIN's "please
    // wait" - which dimmed and blocked a schematic nobody had asked anything of. Off screen the
    // tree is only marked stale; coming back to the tab reads it.
    // ###########################################################################################
    [Fact]
    public async Task A_files_tree_is_not_read_again_while_the_tab_is_off_screen_but_is_on_coming_back()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            MaintainerTabHost.AddOverlay(main);
            Control host = (Control)main.Parent!;
            var window = new Window { Content = host };
            window.Show();
            ReviewQueueRow row = Row();

            Select(main, row);
            main.ShowDetail(Detail(row));

            int requests = 0;
            UseServer(main, new AnsweringHttpHandler(_ =>
            {
                requests++;
                return AnsweringHttpHandler.Json(FilesAnswer());
            }));

            await main.ShowSubmissionViewAsync(SubmissionView.Files);
            Assert.Equal(1, requests);

            // Another tab chosen in CRT: the TabControl takes the tab out of the window.
            window.Content = null;

            main.ShowDetail(Detail(row));
            await Task.Delay(30);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(1, requests);

            // Back to the tab: the stale tree is read now.
            window.Content = host;
            await SettleAsync(() => requests == 2);

            window.Close();
        });
    }

    // ###########################################################################################
    // The Contributor view: who, whether the address was checked, the counts, and every other
    // submission with where it went and what the contributor was told - the owner's
    // "trustworthiness history".
    // ###########################################################################################
    [Fact]
    public void The_contributor_view_shows_the_whole_record()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row();

            Select(main, row);
            main.ShowDetail(Detail(row, contributor: Record));

            Assert.Equal(
                [
                    "Dennis (dh@example.com)",
                    "Sent without an account - the address was typed in when sending, and nothing checked it.",
                    "[2] submissions in total whereof [1] published to BETA and [1] rejected",
                    "Other submissions, newest first",
                    "Wrong board",
                    "#9 - Commodore / C128 / 310378 - Not accepted - sent 2026-September-20",
                    "Told the contributor: This is the 310378.",
                    "Fixed U8",
                    "#7 - Commodore / C64 / 250407 - Published to the BETA source - sent 2026-September-10"
                ],
                main.ContributorViewTextsForTests());
        });
    }

    // ###########################################################################################
    // Another submission chosen: back to Board data, with nothing of the previous one's contributor
    // or files left to be read as this one's. The same submission's detail read again (the minute
    // check, a save) keeps the view the maintainer is on.
    // ###########################################################################################
    [Fact]
    public void Another_submission_opens_on_board_data_and_the_same_one_keeps_its_view()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow first = Row(41);

            Select(main, first);
            main.ShowDetail(Detail(first, submittedFiles: [File("extra.png", "bb", published: null)], contributor: Record));

            Click(main, "ContributorViewButton");

            // The same submission's detail again: the view stays.
            main.ShowDetail(Detail(first, submittedFiles: [File("extra.png", "bb", published: null)], contributor: Record));
            Assert.Equal(SubmissionView.Contributor, main.ShownSubmissionView);

            Select(main, Row(42));

            Assert.Equal(SubmissionView.BoardData, main.ShownSubmissionView);
            Assert.True(Shown(main, "BoardDataView"));
            Assert.Empty(main.ContributorViewTextsForTests());
            Assert.Null(main.FilesCountForTests);
        });
    }
}
