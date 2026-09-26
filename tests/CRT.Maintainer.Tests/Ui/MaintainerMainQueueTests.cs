using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer.Tests.Ui;

// ###########################################################################################
// The maintainer window around the table (owner requests, 2026-09-26): the queue grouped by
// board - a heading per board, each submission its comment and one grey line
// (MaintainerMain.QueueItems.cs) - the lines about what the table cannot show, and the queue's
// buttons, which ran out of their column and over the panel beside it.
//
// The window is BUILT, never shown: its OnOpened restores the real signed-in session and fetches
// the live queue. Where layout matters, its CONTENT is moved into a plain window and that is shown.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MaintainerMainQueueTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(
        long id = 42,
        bool? isNewSystem = false,
        bool? awaitsYou = true,
        string summary = "  Corrected U8.  ",
        bool touchesSharedFiles = false,
        string systemId = "Commodore/C64/250407",
        string state = "pending") =>
        new(id, systemId, state, summary, "c@example.com",
            DateTimeOffset.UtcNow.AddDays(-2).AddMinutes(-5), touchesSharedFiles, isNewSystem, awaitsYou);

    private static ReviewSubmissionDetail Detail(
        ReviewQueueRow row,
        bool isNewSystem = false,
        bool canPublish = true,
        ApprovalStatus? approval = null,
        IReadOnlyList<ReviewSectionView>? sections = null,
        IReadOnlyList<ReviewFindingView>? findings = null,
        IReadOnlyList<SubmittedFileFact>? submittedFiles = null) =>
        new(
            row,
            canPublish,
            new ReviewChangeSummaryView(isNewSystem, sections ?? []),
            findings ?? [],
            new ReviewSubmissionAssets([]),
            [],
            new Dictionary<string, string>(),
            new Dictionary<string, string>(),
            submittedFiles ?? [],
            approval);

    // A shared-file change this account has already approved: waiting for the other approver.
    private static ApprovalStatus ApprovedByYou() =>
        new(
            Required: [ApproverRole.Maintainer, ApproverRole.Administrator],
            Given: [new GivenApproval(ApproverRole.Administrator, "Dennis", DateTimeOffset.UtcNow)],
            WaitingFor: [ApproverRole.Maintainer],
            YourRole: ApproverRole.Administrator,
            CanApprove: false,
            ApprovalPublishes: false,
            YouApproved: true);

    private static void Queue(MaintainerMain main, params ReviewQueueRow[] rows) =>
        main.ApplyQueueResponse(new ReviewQueueResponse(CanPublish: true, Submissions: rows, IsAdministrator: true));

    private static void Select(MaintainerMain main, ReviewQueueRow? row) =>
        typeof(MaintainerMain).GetMethod("ShowSubmission", Any)!.Invoke(main, [row]);

    private static bool Shown(MaintainerMain main, string name) => main.FindControl<Control>(name)!.IsVisible;

    // ###########################################################################################
    // *** THE QUEUE GROUPED BY BOARD (owner request, 2026-09-26). *** Each board once, as a
    // heading; its submissions under it as their comment and one grey line. "New system" is on the
    // heading of a board with nothing published; a published board carries no badge. The queue's
    // oldest-first order holds: the board with the longest-waiting submission leads.
    // ###########################################################################################
    [Fact]
    public void The_queue_is_grouped_by_board_with_each_submission_in_two_short_lines()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            Queue(
                main,
                Row(id: 41, touchesSharedFiles: true),
                Row(id: 50, systemId: "Commodore/C65/Prototype", isNewSystem: true, summary: "First data for the C65."),
                Row(id: 42, summary: "Added the CIA pictures."));

            Assert.Equal(
                [
                    "Commodore / C64 / 250407",
                    "Corrected U8.", "Waiting 2 days - changes shared files",
                    "Added the CIA pictures.", "Waiting 2 days",
                    "Commodore / C65 / Prototype", "New system",
                    "First data for the C65.", "Waiting 2 days"
                ],
                main.QueueTextsForTests());
        });
    }

    // ###########################################################################################
    // Only the exception is marked: a submission NOT waiting for this account is dimmed, and says
    // it is with the other approver - where "Awaiting your review" sat on nearly every row.
    // ###########################################################################################
    [Fact]
    public void A_submission_not_waiting_for_you_is_dimmed_and_says_why()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            Queue(main, Row(id: 41), Row(id: 42, awaitsYou: false, state: "approved", touchesSharedFiles: true));

            Assert.False(main.QueueEntryIsDimmedForTests(41));
            Assert.True(main.QueueEntryIsDimmedForTests(42));
            Assert.Contains("Waiting 2 days - changes shared files - with the other approver", main.QueueTextsForTests());
        });
    }

    // ###########################################################################################
    // *** THE OPENED SUBMISSION FOLLOWS WHAT IT SAYS ITSELF. *** The queue said it waited for you
    // on a published board; the detail - judged against the tree as it is now, and what the
    // Approve button follows - says the board is new and your part is done. The list follows it.
    // ###########################################################################################
    [Fact]
    public void The_opened_submission_follows_what_the_submission_itself_says()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            ReviewQueueRow row = Row(isNewSystem: false, awaitsYou: true);

            Queue(main, row);
            Select(main, row);

            Assert.DoesNotContain(ReviewQueueDisplay.NewSystemBadge, main.QueueTextsForTests());

            main.ShowDetail(Detail(row, isNewSystem: true, approval: ApprovedByYou()));

            Assert.Contains(ReviewQueueDisplay.NewSystemBadge, main.QueueTextsForTests());
            Assert.True(main.QueueEntryIsDimmedForTests(42));
            Assert.Contains("Waiting 2 days - with the other approver", main.QueueTextsForTests());
        });
    }

    // ###########################################################################################
    // A heading is not a submission: a real click on it selects nothing, and a click on a
    // submission selects THAT submission - by item, since the headings shift every position.
    // ###########################################################################################
    [Fact]
    public void A_heading_cannot_be_selected_and_a_submission_is_selected_by_itself()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            typeof(MaintainerMain).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);

            // Two boards, so the second board's submission sits two headings down.
            Queue(main, Row(id: 41), Row(id: 50, systemId: "Commodore/C65/Prototype", summary: "C65."));

            object? content = main.Content;
            main.Content = null;

            var window = new Window { Content = content, Width = main.Width, Height = main.Height };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            ListBox queue = main.FindControl<ListBox>("QueueList")!;
            List<ListBoxItem> items = queue.GetVisualDescendants().OfType<ListBoxItem>().ToList();

            ListBoxItem heading = items.First(item => item.Classes.Contains("QueueHeading"));
            Assert.False(heading.IsEnabled);

            Click(window, heading);
            Assert.Null(main.SelectedQueueRowForTests);

            ListBoxItem c65 = items.Single(item => (item.Tag as ReviewQueueRow)?.Id == 50);
            Click(window, c65);
            Assert.Equal(50, main.SelectedQueueRowForTests?.Id);

            window.Close();
        });
    }

    private static void Click(Window window, Control control)
    {
        Point centre = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;
        window.MouseDown(centre, Avalonia.Input.MouseButton.Left);
        window.MouseUp(centre, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    // The panel has no header of its own any more - the row carries it - so it says only that
    // nothing is chosen, and nothing once something is.
    [Fact]
    public void The_panel_repeats_nothing_of_the_row()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            Assert.True(Shown(main, "NoSubmissionText"));

            Select(main, Row());

            Assert.False(Shown(main, "NoSubmissionText"));
            Assert.Null(main.FindControl<Control>("SubmissionHeaderPanel"));
        });
    }

    // A line built of runs has no Text of its own - its words are in its Inlines.
    private static string ShownText(TextBlock block) =>
        block.Text ?? string.Concat(block.Inlines!.OfType<Avalonia.Controls.Documents.Run>().Select(run => run.Text));

    // ###########################################################################################
    // Who sent it, and how their other submissions went, above the table (2026-09-26) - gone
    // again when nothing is selected.
    // ###########################################################################################
    [Fact]
    public void The_contributor_and_their_record_are_shown_above_the_table()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            ReviewQueueRow row = Row();
            Select(main, row);

            TextBlock line = main.FindControl<TextBlock>("ContributorText")!;

            main.ShowDetail(Detail(row) with { Contributor = new ReviewContributorFacts("dh@hinet.dk", null, 2, 0, 0, 1) });

            Assert.True(line.IsVisible);
            Assert.Equal("From dh@hinet.dk - [3] other submissions: [2] published, [1] rejected", ShownText(line));

            main.ShowDetail(Detail(row) with { Contributor = new ReviewContributorFacts("dh@hinet.dk", null, 0, 0, 0, 0) });
            Assert.Equal("From dh@hinet.dk - no other submissions", ShownText(line));

            Select(main, row);
            Assert.False(line.IsVisible);
        });
    }

    // ###########################################################################################
    // What the table cannot show is listed above it - only when there is something to list.
    // ###########################################################################################
    [Fact]
    public void Highlights_and_warnings_the_table_cannot_show_are_listed_above_it()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            ReviewQueueRow row = Row();
            Select(main, row);

            main.ShowDetail(Detail(row));
            Assert.False(Shown(main, "NotInTablePanel"));

            var highlights = new ReviewSectionView(
                ReviewSummary.SectionComponentHighlights,
                Added: [],
                Removed: [],
                Changed: [$"Schematic 1{BoardDraftNaturalKeys.Separator}U8"],
                Renamed: [],
                FieldChanges: new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>());

            main.ShowDetail(Detail(
                row,
                sections: [highlights],
                findings: [new ReviewFindingView("x", "U8", "The picture is very large.", IsError: false)],
                submittedFiles:
                [
                    new SubmittedFileFact(
                        "Commodore/C64/250407/KiCad data/board.kicad_pcb", "aa", 10,
                        SubmissionFileScope.Own, IsReferenced: false, PublishedSha256: null)
                ]));

            StackPanel panel = main.FindControl<StackPanel>("NotInTablePanel")!;
            Assert.True(panel.IsVisible);
            Assert.Equal(
                [
                    "Component highlights: [1] changed (Schematic 1 / U8)",
                    "KiCad data included: [1] file ([1] new)",
                    "Warning from the automatic checks: The picture is very large. [U8]"
                ],
                panel.Children.OfType<TextBlock>().Select(ShownText));

            // The count's number alone is bold - a Run of its own inside the line.
            TextBlock counted = panel.Children.OfType<TextBlock>().First();
            Assert.Equal(
                ["1"],
                counted.Inlines!.OfType<Avalonia.Controls.Documents.Run>()
                    .Where(run => run.FontWeight == Avalonia.Media.FontWeight.Bold)
                    .Select(run => run.Text));

            // Another submission starts clean.
            Select(main, row);
            Assert.False(panel.IsVisible);
        });
    }

    // ------------------------------------------------------------------ The queue keeps itself up to date (2026-09-26)

    // A submission's table, for the tests that open one.
    private static ReviewTableData Table() =>
        new(
            Version: 1,
            Published: new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CIA" }] },
            Submitted: new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CIA 6526" }] });

    private static void EditTheTable(MaintainerMain main)
    {
        BoardTableDocument document = main.TableEditorForTests.CommitAndGetDocument()!;
        BoardTableRow row = document.FindSheet(BoardWorkbookSchema.SheetComponents)!.Rows.First(candidate => !candidate.IsDeleted);
        row.Cells[BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName)].Text = "Edited";
    }

    private static ReviewQueueResponse Answer(params ReviewQueueRow[] rows) =>
        new(CanPublish: true, Submissions: rows, IsAdministrator: false);

    [Fact]
    public void There_is_no_refresh_button()
    {
        UiTest.Run(() => Assert.Null(new MaintainerMain().FindControl<Button>("RefreshButton")));
    }

    // ###########################################################################################
    // *** A CHECK NOBODY ASKED FOR NEVER DISTURBS THE OPEN SUBMISSION. *** A new submission joins
    // the list; the one being worked on stays selected, and its table - with an edit in it - is
    // left exactly as it was.
    // ###########################################################################################
    [Fact]
    public void A_background_check_updates_the_list_and_leaves_the_open_table_alone()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            ReviewQueueRow open = Row(id: 42);

            Queue(main, open);
            Select(main, open);
            main.OpenTableForTests(open, Table());
            EditTheTable(main);

            main.ApplyQueueResponse(Answer(open, Row(id: 43, summary: "Just arrived.")), background: true);

            Assert.Contains("Just arrived.", main.QueueTextsForTests());
            Assert.Equal(42, main.SelectedQueueRowForTests?.Id);
            Assert.True(main.IsTableOpen);
            Assert.True(main.TableEditorForTests.HasUnsavedChanges);
        });
    }

    // ###########################################################################################
    // *** DECIDED ELSEWHERE WHILE OPEN: IT STAYS, UNDECIDABLE. *** Another maintainer approved it;
    // the queue's own check finds it gone. The table stays on screen with the decisions off and one
    // line saying why - until the maintainer picks another. A refresh after a decision HERE closes
    // it, as before.
    // ###########################################################################################
    [Fact]
    public void A_submission_decided_elsewhere_stays_on_screen_with_its_decisions_off()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            ReviewQueueRow open = Row(id: 42);

            Queue(main, open);
            Select(main, open);
            main.OpenTableForTests(open, Table());
            main.ShowDetail(Detail(open));

            Assert.True(main.FindControl<Button>("ApproveButton")!.IsEnabled);

            main.ApplyQueueResponse(Answer(), background: true);

            Assert.True(main.IsTableOpen);
            Assert.All(
                new[] { "ApproveButton", "RejectButton", "RequestChangesButton" },
                name => Assert.False(main.FindControl<Button>(name)!.IsEnabled));
            Assert.Equal(ReviewDecisionWording.DecidedElsewhere, main.FindControl<TextBlock>("DecisionMessageText")!.Text);

            main.ApplyQueueResponse(Answer());

            Assert.False(main.IsTableOpen);
        });
    }

    // Opening another submission after that makes its decisions possible again.
    [Fact]
    public void The_next_submission_opened_can_be_decided_again()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            ReviewQueueRow gone = Row(id: 42);
            ReviewQueueRow next = Row(id: 43);

            Queue(main, gone, next);
            Select(main, gone);
            main.ShowDetail(Detail(gone));
            main.ApplyQueueResponse(Answer(next), background: true);

            Select(main, next);
            main.ShowDetail(Detail(next));

            Assert.All(
                new[] { "ApproveButton", "RejectButton", "RequestChangesButton" },
                name => Assert.True(main.FindControl<Button>(name)!.IsEnabled));
        });
    }

    // ------------------------------------------------------------------ The look (2026-09-26)

    // The account, named without "Signed in as" - somebody with two accounts still sees which one
    // is acting.
    [Fact]
    public void The_footer_names_the_account_and_nothing_more()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            typeof(MaintainerMain).GetField("thisSession", Any)!
                .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));
            typeof(MaintainerMain).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);

            Assert.Equal("Dennis (dh@example.com)", main.FindControl<TextBlock>("SignedInAsText")!.Text);
        });
    }

    // The three decisions in CRT's own red, white on IndianRed - the table's "Save changes" colours.
    // Approve at the far right (owner request, 2026-09-26), apart from the other two.
    [Fact]
    public void The_decision_buttons_are_CRTs_red_with_approve_at_the_right()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            typeof(MaintainerMain).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);

            object? content = main.Content;
            main.Content = null;

            // Wide: headless text is far wider than real text, and the gap is what is checked.
            var window = new Window { Content = content, Width = 1800, Height = 720 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            foreach (string name in new[] { "RequestChangesButton", "RejectButton", "ApproveButton" })
            {
                Button button = main.FindControl<Button>(name)!;

                Assert.Equal(Avalonia.Media.Colors.IndianRed, ((Avalonia.Media.ISolidColorBrush)button.Background!).Color);
                Assert.Equal(Avalonia.Media.Colors.White, ((Avalonia.Media.ISolidColorBrush)button.Foreground!).Color);
            }

            StackPanel bar = main.FindControl<StackPanel>("DecisionPanel")!;
            bar.IsVisible = true;
            Dispatcher.UIThread.RunJobs();

            double Right(Control control) => control.TranslatePoint(new Point(control.Bounds.Width, 0), window)!.Value.X;
            double Left(Control control) => control.TranslatePoint(default, window)!.Value.X;

            Button approve = main.FindControl<Button>("ApproveButton")!;
            Assert.Equal(Right(bar), Right(approve), precision: 0);
            double gap = Left(approve) - Right(main.FindControl<Button>("RejectButton")!);
            Assert.True(gap > 100, $"bar right {Right(bar)}, approve {Left(approve)}-{Right(approve)}, gap {gap}");

            window.Close();
        });
    }

    // ###########################################################################################
    // Maintainers and Unused files are the administrator's screens: offered only when the server
    // says this account is one (the server refuses everybody else anyway). Production is for all.
    // ###########################################################################################
    [Fact]
    public void The_administrators_screens_are_offered_only_to_an_administrator()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            main.ApplyQueueResponse(new ReviewQueueResponse(CanPublish: true, Submissions: [], IsAdministrator: false));

            Assert.False(Shown(main, "MaintainersButton"));
            Assert.False(Shown(main, "UnusedFilesButton"));
            Assert.True(Shown(main, "ProductionButton"));

            main.ApplyQueueResponse(new ReviewQueueResponse(CanPublish: true, Submissions: [], IsAdministrator: true));

            Assert.True(Shown(main, "MaintainersButton"));
            Assert.True(Shown(main, "UnusedFilesButton"));
        });
    }

    // ###########################################################################################
    // *** THE REPORTED OVERLAP, AND THE ROWS. *** Title and four buttons in one row of Auto columns
    // were wider than the queue's column, so for an administrator they ran out of it and over the
    // panel beside it. Every button now ends inside the column, and so does every line of a row -
    // a long description wraps rather than running off, at the window's own size.
    // ###########################################################################################
    [Fact]
    public void The_queue_buttons_and_rows_stay_inside_the_queue_column()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            typeof(MaintainerMain).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);
            Queue(main, Row(summary: string.Join(' ', Enumerable.Repeat("A very long description of what changed.", 8))));

            // Its content in a plain window - MaintainerMain's own would sign in to the server.
            object? content = main.Content;
            main.Content = null;

            var window = new Window { Content = content, Width = main.Width, Height = main.Height };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            ListBox queue = main.FindControl<ListBox>("QueueList")!;
            double columnRight = queue.TranslatePoint(new Point(queue.Bounds.Width, 0), window)!.Value.X;

            IEnumerable<Control> checkedControls = new[] { "ProductionButton", "MaintainersButton", "UnusedFilesButton" }
                .Select(name => (Control)main.FindControl<Button>(name)!)
                .Concat(queue.GetVisualDescendants().OfType<TextBlock>().Where(text => text.IsEffectivelyVisible));

            foreach (Control control in checkedControls)
            {
                double right = control.TranslatePoint(new Point(control.Bounds.Width, 0), window)!.Value.X;

                Assert.True(right <= columnRight + 0.5, $"{control.Name ?? (control as TextBlock)?.Text} ends at {right}, past the queue column's edge at {columnRight}.");
            }

            window.Close();
        });
    }
}
