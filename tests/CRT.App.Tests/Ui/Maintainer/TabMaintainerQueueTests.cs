using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The maintainer window around the table (owner requests, 2026-09-26): the queue grouped by
// board - a heading per board, each submission its comment and one grey line
// (TabMaintainer.QueueItems.cs) - the lines about what the table cannot show, and the queue's
// buttons, which ran out of their column and over the panel beside it.
//
// The window is BUILT, never shown: its OnOpened restores the real signed-in session and fetches
// the live queue. Where layout matters, its CONTENT is moved into a plain window and that is shown.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerQueueTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(
        long id = 42,
        bool? isNewBoard = false,
        bool? awaitsYou = true,
        string summary = "  Corrected U8.  ",
        bool touchesSharedFiles = false,
        string boardId = "Commodore/C64/250407",
        string state = "pending") =>
        new(id, boardId, state, summary, "c@example.com",
            DateTimeOffset.UtcNow.AddDays(-2).AddMinutes(-5), touchesSharedFiles, isNewBoard, awaitsYou);

    private static ReviewSubmissionDetail Detail(
        ReviewQueueRow row,
        bool isNewBoard = false,
        bool canPublish = true,
        ApprovalStatus? approval = null,
        IReadOnlyList<ReviewSectionView>? sections = null,
        IReadOnlyList<ReviewFindingView>? findings = null,
        IReadOnlyList<SubmittedFileFact>? submittedFiles = null) =>
        new(
            row,
            canPublish,
            new ReviewChangeSummaryView(isNewBoard, sections ?? []),
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

    private static void Queue(TabMaintainer main, params ReviewQueueRow[] rows) =>
        main.ApplyQueueResponse(new ReviewQueueResponse(CanPublish: true, Submissions: rows, IsAdministrator: true));

    private static void Select(TabMaintainer main, ReviewQueueRow? row) =>
        typeof(TabMaintainer).GetMethod("ShowSubmission", Any)!.Invoke(main, [row]);

    private static bool Shown(TabMaintainer main, string name) => main.FindControl<Control>(name)!.IsVisible;

    // ###########################################################################################
    // *** THE QUEUE GROUPED BY BOARD (owner request, 2026-09-26). *** Each board once, as a
    // heading; its submissions under it as their comment and one grey line. "New board" is on the
    // heading of a board with nothing published; a published board carries no badge. The queue's
    // oldest-first order holds: the board with the longest-waiting submission leads.
    // ###########################################################################################
    [Fact]
    public void The_queue_is_grouped_by_board_with_each_submission_in_two_short_lines()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Queue(
                main,
                Row(id: 41, touchesSharedFiles: true),
                Row(id: 50, boardId: "Commodore/C65/Prototype", isNewBoard: true, summary: "First data for the C65."),
                Row(id: 42, summary: "Added the CIA pictures."));

            Assert.Equal(
                [
                    "Commodore / C64 / 250407",
                    "Corrected U8.", "Waiting 2 days - replaces a shared file",
                    "Added the CIA pictures.", "Waiting 2 days",
                    "Commodore / C65 / Prototype", "New board",
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
            var main = new TabMaintainer();

            Queue(main, Row(id: 41), Row(id: 42, awaitsYou: false, state: "approved", touchesSharedFiles: true));

            Assert.False(main.QueueEntryIsDimmedForTests(41));
            Assert.True(main.QueueEntryIsDimmedForTests(42));
            Assert.Contains("Waiting 2 days - replaces a shared file - with the other approver", main.QueueTextsForTests());
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
            var main = new TabMaintainer();
            ReviewQueueRow row = Row(isNewBoard: false, awaitsYou: true);

            Queue(main, row);
            Select(main, row);

            Assert.DoesNotContain(ReviewQueueDisplay.NewBoardBadge, main.QueueTextsForTests());

            main.ShowDetail(Detail(row, isNewBoard: true, approval: ApprovedByYou()));

            Assert.Contains(ReviewQueueDisplay.NewBoardBadge, main.QueueTextsForTests());
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
            var main = new TabMaintainer();
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);

            // Two boards, so the second board's submission sits two headings down.
            Queue(main, Row(id: 41), Row(id: 50, boardId: "Commodore/C65/Prototype", summary: "C65."));


            var window = new Window { Content = main, Width = 1100, Height = 720 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            ListBox queue = main.FindControl<ListBox>("QueueList")!;
            List<ListBoxItem> items = queue.GetVisualDescendants().OfType<ListBoxItem>().ToList();

            ListBoxItem heading = items.First(item => item.Classes.Contains("QueueHeading"));
            Assert.False(heading.IsEnabled);

            // Showing the tab opened the first submission (owner request, 2026-09-30,
            // TabMaintainer.OpenOnEntry.cs); a click on a heading leaves it as it is.
            Assert.Equal(41, main.SelectedQueueRowForTests?.Id);

            Click(window, heading);
            Assert.Equal(41, main.SelectedQueueRowForTests?.Id);
            Assert.NotSame(heading, queue.SelectedItem);

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
            var main = new TabMaintainer();

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
    // *** WHO SENT IT IS THE CONTRIBUTOR VIEW'S ALONE (owner request, 2026-09-30: "I do not need to
    // see highlighted data, as this should be available under "Contributor""). *** The one-line
    // "From dh@hinet.dk - [3] other submissions: ..." above the table is gone; the same name and
    // counts open the Contributor view.
    // ###########################################################################################
    [Fact]
    public void The_contributor_is_not_repeated_above_the_table_but_opens_the_Contributor_view()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row();
            Select(main, row);

            main.ShowDetail(Detail(row) with { Contributor = new ReviewContributorFacts("dh@hinet.dk", null, 2, 0, 0, 1) });

            Assert.Null(main.FindControl<TextBlock>("ContributorText"));

            IReadOnlyList<string> texts = main.ContributorViewTextsForTests();
            Assert.Equal("dh@hinet.dk", texts[0]);
            Assert.Contains("[3] submissions in total whereof [2] published and [1] rejected", texts);
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
            var main = new TabMaintainer();
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

            // The new KiCad file is NOT a line here any more (owner request, 2026-09-30) - it is
            // the Files button's count, the view that shows it.
            // A section's changes one a line under a counted heading, set in under it (owner
            // request, 2026-10-04).
            Assert.Equal(
                [
                    "Component highlights have [1] change:",
                    "Changed component [U8] on schematic \"Schematic 1\"",
                    "Warning from the automatic checks: The picture is very large. [U8]"
                ],
                panel.Children.OfType<TextBlock>().Select(ShownText));
            Assert.Equal(
                [0d, 16d, 0d],
                panel.Children.OfType<TextBlock>().Select(block => block.Margin.Left));
            Assert.Equal("1", main.FilesCountForTests);

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

    private static void EditTheTable(TabMaintainer main)
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
        UiTest.Run(() => Assert.Null(new TabMaintainer().FindControl<Button>("RefreshButton")));
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
            var main = new TabMaintainer();
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
            var main = new TabMaintainer();
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
            var main = new TabMaintainer();
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

    // The account, named under "Logged in as:" (owner request, 2026-10-04) - somebody with two
    // accounts still sees which one is acting.
    [Fact]
    public void The_footer_names_the_account_and_nothing_more()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            typeof(TabMaintainer).GetField("thisSession", Any)!
                .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);

            Assert.Equal("Dennis (dh@example.com)", TabMaintainer.TextOf(main.FindControl<TextBlock>("SignedInAsText")!));
        });
    }

    // The three decisions in CRT's own red, white on IndianRed - the table's "Save changes" colours.
    // Approve at the far right (owner request, 2026-09-26), apart from the other two.
    [Fact]
    public void Approve_is_green_the_other_two_red_with_approve_at_the_right()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);


            // Wide: headless text is far wider than real text, and the gap is what is checked.
            var window = new Window { Content = main, Width = 1800, Height = 720 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            // ###########################################################################################
            // *** THE COLOUR SAYS WHICH WAY THE DECISION GOES (owner request, 2026-09-26: "then
            // there are two red ones ... and then one accepting it, being green. Should be
            // logical"). *** All three were red, so the one accepting action looked like the two
            // that send the submission back. Asserted on the RESOLVED colour, not on the resource
            // key: a DynamicResource naming a key this application does not define draws an
            // unstyled button in silence, which is how a green Approve could quietly go grey.
            // ###########################################################################################
            foreach (string name in new[] { "RequestChangesButton", "RejectButton" })
            {
                Button button = main.FindControl<Button>(name)!;

                Assert.Equal(Avalonia.Media.Colors.IndianRed, ((Avalonia.Media.ISolidColorBrush)button.Background!).Color);
                Assert.Equal(Avalonia.Media.Colors.White, ((Avalonia.Media.ISolidColorBrush)button.Foreground!).Color);
            }

            Button approveButton = main.FindControl<Button>("ApproveButton")!;
            Assert.Equal(
                Avalonia.Media.Color.Parse("#4F8A5B"),
                ((Avalonia.Media.ISolidColorBrush)approveButton.Background!).Color);
            Assert.Equal(Avalonia.Media.Colors.White, ((Avalonia.Media.ISolidColorBrush)approveButton.Foreground!).Color);

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
    // *** THE ROWS STAY IN THE QUEUE'S COLUMN. *** Every line of a row ends inside the column - a
    // long description wraps rather than running off over the panel beside it, at the window's own
    // size. (The four screen buttons were checked here too while they sat in this column; since
    // 2026-10-01 they are a tab strip across the whole tab - TabMaintainerScreenTabsTests.)
    // ###########################################################################################
    [Fact]
    public void The_queue_rows_stay_inside_the_queue_column()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);
            Queue(main, Row(summary: string.Join(' ', Enumerable.Repeat("A very long description of what changed.", 8))));

            // The tab in a plain window, for a layout pass. (It moved its CONTENT into one while it
            // was a Window, whose own showing signed in to the server; a tab restores a session only
            // from ReviewSessionStore, which no test points at a real file.)

            var window = new Window { Content = main, Width = 1100, Height = 720 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            ListBox queue = main.FindControl<ListBox>("QueueList")!;
            double columnRight = queue.TranslatePoint(new Point(queue.Bounds.Width, 0), window)!.Value.X;

            List<TextBlock> checkedControls = queue.GetVisualDescendants().OfType<TextBlock>().Where(text => text.IsEffectivelyVisible).ToList();

            Assert.NotEmpty(checkedControls);

            foreach (Control control in checkedControls)
            {
                double right = control.TranslatePoint(new Point(control.Bounds.Width, 0), window)!.Value.X;

                Assert.True(right <= columnRight + 0.5, $"{control.Name ?? (control as TextBlock)?.Text} ends at {right}, past the queue column's edge at {columnRight}.");
            }

            window.Close();
        });
    }

    // ###########################################################################################
    // *** WHY APPROVE IS OFF IS THE AMBER NOTICE PANEL (owner request, 2026-10-09: "this should look
    // the same as the previous highlighted panel ... The UI should have a uniform and consistent
    // look"). *** It was an orange line. Now it is the Boards screen's read-only panel, drawn alike:
    // read off a SHOWN window, both panels have the same fill, edge and ink, each with its icon -
    // and the panel is gone again once nothing holds Approve back.
    // ###########################################################################################
    [Fact]
    public void Why_Approve_is_off_is_the_same_amber_panel_as_the_Boards_screens_read_only_notice()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            var window = new Window { Content = main, Width = 1400, Height = 900 };
            window.Show();

            ReviewQueueRow open = Row(id: 42);
            main.ApplyBetaListAsync(new ProductionListResponse(true,
            [
                new ProductionBoardRow("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-27", "h", null, null)
            ]), background: true).GetAwaiter().GetResult();

            Queue(main, open);
            Select(main, open);
            main.ShowDetail(Detail(open));

            // The Boards screen's panel, raised the way a board waiting in BETA raises it.
            BoardDetailView board = main.BoardDetailForTests;
            board.OpenTableForTests(new BoardTableAnswer("Commodore/C64/250407", new string('f', 64), new SubmissionRows(), MayEdit: false, "It waits in BETA."));
            Dispatcher.UIThread.RunJobs();

            Border queueNotice = main.FindControl<Border>("BeforeApprovingNotice")!;
            Border boardNotice = board.FindControl<Border>("ReadOnlyNotice")!;
            TextBlock queueLine = main.FindControl<StackPanel>("BeforeApprovingPanel")!.Children.OfType<TextBlock>().Single();
            TextBlock boardLine = board.FindControl<TextBlock>("ReadOnlyNoticeText")!;

            Assert.True(queueNotice.IsVisible);
            Assert.Equal(OneSubmissionInBeta.BusyMessage("Commodore/C64/250407"), ShownText(queueLine));

            Assert.Contains("Notice", queueNotice.Classes);
            Assert.Contains("Notice", boardNotice.Classes);
            Assert.Equal(Colour(boardNotice.Background), Colour(queueNotice.Background));
            Assert.Equal(Colour(boardNotice.BorderBrush), Colour(queueNotice.BorderBrush));
            Assert.Equal(boardNotice.CornerRadius, queueNotice.CornerRadius);
            Assert.Equal(Colour(boardLine.Foreground), Colour(queueLine.Foreground));
            Assert.Equal(boardLine.FontWeight, queueLine.FontWeight);
            Assert.Single(queueNotice.GetVisualDescendants().OfType<TextBlock>(), block => block.Classes.Contains("NoticeIcon"));
            Assert.Single(boardNotice.GetVisualDescendants().OfType<TextBlock>(), block => block.Classes.Contains("NoticeIcon"));

            // *** AND IN THE SAME PLACE (owner request, same day: "one is shown as the first info and
            // the other is shown above the table? Again, this kind of UI should be unified"). ***
            // Each directly above its own view switch, and said whichever view is open.
            Assert.Equal("SubmissionViewBar", NextSibling(queueNotice).Name);
            Assert.Equal("BoardSectionBar", NextSibling(boardNotice).Name);

            // Outside every one of the board's views - it sat inside Board data, gone on the others.
            for (StyledElement? parent = boardNotice.Parent; parent is not null; parent = parent.Parent)
                Assert.NotEqual("SectionsPanel", (parent as Control)?.Name);

            board.ShowSectionAsync(BoardSection.History).GetAwaiter().GetResult();
            Assert.True(boardNotice.IsVisible);

            // Promoted: nothing holds Approve back, and the panel goes with its last reason.
            main.ApplyBetaListAsync(new ProductionListResponse(true, []), background: true).GetAwaiter().GetResult();

            Assert.False(queueNotice.IsVisible);

            window.Close();
        });

        static Avalonia.Media.Color Colour(Avalonia.Media.IBrush? brush) => ((Avalonia.Media.ISolidColorBrush)brush!).Color;

        static Control NextSibling(Control control)
        {
            var parent = (Panel)control.Parent!;
            return parent.Children[parent.Children.IndexOf(control) + 1];
        }
    }

    // ###########################################################################################
    // *** ONE SUBMISSION IN BETA PER BOARD, ON SCREEN (owner decision, 2026-09-27). *** With an
    // earlier submission of the board waiting in "Beta > Prod", Approve is off - a line above the
    // submission's three views says why, in the server's own words - while Reject and Request changes
    // stay on. Another board in BETA is no reason.
    // ###########################################################################################
    [Fact]
    public void Approve_is_off_while_an_earlier_submission_of_the_board_waits_in_BETA()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow open = Row(id: 42);

            main.ApplyBetaListAsync(new ProductionListResponse(true,
            [
                new ProductionBoardRow("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-27", "h", null, null, true)
            ]), background: true).GetAwaiter().GetResult();

            Queue(main, open);
            Select(main, open);
            main.ShowDetail(Detail(open));

            Button approve = main.FindControl<Button>("ApproveButton")!;
            Assert.False(approve.IsEnabled);
            // The reason is said above the views, not in a tooltip on the button: a tooltip at the
            // bottom of the window lands over the pointer and takes the click (2026-09-30).
            Assert.Null(ToolTip.GetTip(approve));
            Assert.True(main.FindControl<Button>("RejectButton")!.IsEnabled);
            Assert.Contains(
                OneSubmissionInBeta.BusyMessage("Commodore/C64/250407"),
                main.FindControl<StackPanel>("BeforeApprovingPanel")!.Children.OfType<TextBlock>().Select(ShownText));

            // *** AND IT STAYS SAID ON THE FILES AND CONTRIBUTOR VIEWS (code review, 2026-10-01). ***
            // It was a line inside the Board data view, hidden with it - Approve greyed out with no
            // reason anywhere. The panel must sit outside every one of the three views.
            StackPanel reason = main.FindControl<StackPanel>("BeforeApprovingPanel")!;
            foreach (string view in new[] { "BoardDataView", "FilesView", "ContributorView" })
            {
                Control panel = main.FindControl<Control>(view)!;
                for (StyledElement? parent = reason.Parent; parent is not null; parent = parent.Parent)
                    Assert.NotSame(panel, parent);
            }

            // Another board's submission is not held back.
            ReviewQueueRow other = Row(id: 43, boardId: "Commodore/C128/310378");
            Queue(main, open, other);
            Select(main, other);
            main.ShowDetail(Detail(other));

            Assert.True(main.FindControl<Button>("ApproveButton")!.IsEnabled);
        });
    }

    // ###########################################################################################
    // *** THE BLOCK FOLLOWS THE BETA LIST WITHOUT REOPENING THE SUBMISSION (code review,
    // 2026-09-27). *** It was worked out only when the detail was shown, so a board promoted by
    // another maintainer kept Approve off with a stale reason until the submission was chosen
    // again - and at sign-in a detail shown before the lists arrived was never blocked at all.
    // ###########################################################################################
    [Fact]
    public void Approve_follows_the_BETA_list_as_it_is_refreshed_with_the_submission_open()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow open = Row(id: 42);
            var inBeta = new ProductionBoardRow("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-27", "h", null, null, true);

            // The detail first, the lists after - the order a sign-in can deliver them in.
            Queue(main, open);
            Select(main, open);
            main.ShowDetail(Detail(open));

            Button approve = main.FindControl<Button>("ApproveButton")!;
            Assert.True(approve.IsEnabled);

            main.ApplyBetaListAsync(new ProductionListResponse(true, [inBeta]), background: true).GetAwaiter().GetResult();

            Assert.False(approve.IsEnabled);
            Assert.Contains(
                OneSubmissionInBeta.BusyMessage("Commodore/C64/250407"),
                main.FindControl<StackPanel>("BeforeApprovingPanel")!.Children.OfType<TextBlock>().Select(ShownText));

            // Promoted by somebody else: the next check's list no longer has it.
            main.ApplyBetaListAsync(new ProductionListResponse(true, []), background: true).GetAwaiter().GetResult();

            Assert.True(approve.IsEnabled);
            Assert.DoesNotContain(
                OneSubmissionInBeta.BusyMessage("Commodore/C64/250407"),
                main.FindControl<StackPanel>("BeforeApprovingPanel")!.Children.OfType<TextBlock>().Select(ShownText));
        });
    }

    // Decided elsewhere, the decisions stay off whatever the BETA list says next.
    [Fact]
    public void A_refreshed_BETA_list_never_reopens_a_submission_decided_elsewhere()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow open = Row(id: 42);
            var inBeta = new ProductionBoardRow("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-27", "h", null, null, true);

            main.ApplyBetaListAsync(new ProductionListResponse(true, [inBeta]), background: true).GetAwaiter().GetResult();
            Queue(main, open);
            Select(main, open);
            main.ShowDetail(Detail(open));

            main.ApplyQueueResponse(Answer(), background: true);
            main.ApplyBetaListAsync(new ProductionListResponse(true, []), background: true).GetAwaiter().GetResult();

            Assert.All(
                new[] { "ApproveButton", "RejectButton", "RequestChangesButton" },
                name => Assert.False(main.FindControl<Button>(name)!.IsEnabled));
        });
    }

    // ###########################################################################################
    // *** THE WHOLE WINDOW WAITS WHILE A DECISION IS SENT (owner report, 2026-09-28). *** A small
    // green "Working..." under the table was all an approval showed while it published to BETA,
    // and the screen read as hung. Now the overlay a push-back uses is up - faded screen, every
    // click taken, the sentence naming the board - for as long as the request is in flight, read
    // from INSIDE the fake server's answer. And it is lifted again on a refusal, with the reason
    // where it always was.
    // ###########################################################################################
    [Fact]
    public async Task The_window_waits_while_an_approval_is_sent_and_comes_back_after_a_refusal()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row(boardId: "Commodore/C128/310378");
            Select(main, row);
            main.ShowDetail(Detail(row));

            bool overlayUp = false;
            string? sentence = null;
            BusyOverlay overlay = MaintainerTabHost.AddOverlay(main);

            var server = new AnsweringHttpHandler(_ =>
            {
                overlayUp = overlay.IsVisible && overlay.IsBusy;
                sentence = overlay.Message;
                return AnsweringHttpHandler.Refused();
            });

            typeof(TabMaintainer).GetField("thisClient", Any)!
                .SetValue(main, new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            typeof(TabMaintainer).GetField("thisSession", Any)!
                .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

            await (Task)typeof(TabMaintainer).GetMethod("DecideAsync", Any)!.Invoke(main, [ReviewDecisionKind.Approve])!;

            Assert.True(overlayUp);
            Assert.Equal(ReviewDecisionWording.Waiting(ReviewDecisionKind.Approve, null, "Commodore/C128/310378"), sentence);

            Assert.False(overlay.IsVisible);
            Assert.Equal(1, main.FindControl<Grid>("QueuePanel")!.Opacity);

            TextBlock message = main.FindControl<TextBlock>("DecisionMessageText")!;
            Assert.True(message.IsVisible);
            Assert.NotEqual("Working...", message.Text);

            // The decision is off only while it is in flight.
            Assert.True(main.FindControl<Button>("ApproveButton")!.IsEnabled);
        });
    }

    // The overlay shows the busy pointer wherever the mouse is over the window - the one a wait in
    // the tab runs under, which is CRT's window's since 2026-09-29 (MaintainerTabHost).
    [Fact]
    public void The_please_wait_layer_shows_the_busy_pointer()
    {
        UiTest.Run(() =>
        {
            BusyOverlay overlay = MaintainerTabHost.AddOverlay(new TabMaintainer());

            Assert.Equal(new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Wait).ToString(), overlay.Cursor?.ToString());
        });
    }

    // ###########################################################################################
    // *** NO ANSWER IN TWO MINUTES IS CHECKED, NOT CALLED A FAILURE (owner decision, 2026-09-28:
    // "it must be solid in validating if it did finish"). *** The approval's request never answers;
    // the limit passes (the test's clock, not two real minutes); the submission is read again, the
    // server says it is merged - and the maintainer is told it DID publish, rather than being
    // invited to press Approve a second time over a publish that worked.
    // ###########################################################################################
    [Fact]
    public async Task An_approval_with_no_answer_is_checked_and_reported_as_published_when_it_was()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row(boardId: "Commodore/C128/310378");
            Select(main, row);
            main.ShowDetail(Detail(row));

            BusyOverlay overlay = MaintainerTabHost.AddOverlay(main);

            // The first wait - the approval - runs out at once; every later one never does.
            int limits = 0;
            overlay.LimitOverrideForTests = token => ++limits == 1 ? Task.CompletedTask : Task.Delay(Timeout.Infinite, token);

            var server = new AnsweringHttpHandler((request, token) =>
            {
                string path = request.RequestUri!.AbsolutePath;

                if (path.EndsWith("/approve", StringComparison.Ordinal))
                    return AnsweringHttpHandler.NeverAsync(token);

                if (request.Method == HttpMethod.Get && path.EndsWith("/submissions/42", StringComparison.Ordinal))
                {
                    return Task.FromResult(AnsweringHttpHandler.Json(
                        """{"canPublish":true,"submission":{"id":42,"boardId":"Commodore/C128/310378","state":"merged"},"findings":[]}"""));
                }

                return Task.FromResult(AnsweringHttpHandler.Refused());
            });

            typeof(TabMaintainer).GetField("thisClient", Any)!
                .SetValue(main, new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            typeof(TabMaintainer).GetField("thisSession", Any)!
                .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

            await (Task)typeof(TabMaintainer).GetMethod("DecideAsync", Any)!.Invoke(main, [ReviewDecisionKind.Approve])!;

            Assert.Equal(
                "The server did not answer within 2 minutes, but it did finish: the submission is published to BETA.",
                main.FindControl<TextBlock>("DecisionMessageText")!.Text);
            Assert.False(overlay.IsVisible);
        });
    }

    // ###########################################################################################
    // *** THE DECISION COMMENT BELONGS TO ONE SUBMISSION (owner report, 2026-10-02: "I can see my
    // last rejection comment in the textarea field. This field should be blanked when
    // submitted/rejected"). *** Never emptied, the comment written for one contributor sat ready to
    // go to the next with one click. Moving between submissions keeps each one's own comment, so a
    // half-written one is still there on coming back.
    // ###########################################################################################
    [Fact]
    public void A_comment_typed_for_one_submission_is_not_shown_on_another_and_is_back_on_returning()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow first = Row(id: 42);
            ReviewQueueRow second = Row(id: 43);
            Queue(main, first, second);
            TextBox comment = main.FindControl<TextBox>("DecisionCommentTextBox")!;

            Select(main, first);
            comment.Text = "The U8 picture is of the wrong board revision.";

            Select(main, second);
            Assert.True(string.IsNullOrEmpty(comment.Text));
            comment.Text = "Please add the pin numbers.";

            Select(main, first);
            Assert.Equal("The U8 picture is of the wrong board revision.", comment.Text);

            Select(main, second);
            Assert.Equal("Please add the pin numbers.", comment.Text);
        });
    }

    // The reported case: a rejection goes through, the queue empties, and the next submission to
    // arrive opens with an EMPTY box - not the rejection just sent.
    [Fact]
    public async Task A_rejection_that_went_through_empties_the_comment_box_for_the_next_submission()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row(id: 15);
            Queue(main, row);
            Select(main, row);
            main.ShowDetail(Detail(row));
            MaintainerTabHost.AddOverlay(main);

            TextBox comment = main.FindControl<TextBox>("DecisionCommentTextBox")!;
            comment.Text = "The schematic image is the wrong board revision.";

            var server = new AnsweringHttpHandler(request =>
            {
                string path = request.RequestUri!.AbsolutePath;

                if (path.EndsWith("/reject", StringComparison.Ordinal))
                    return AnsweringHttpHandler.Json("""{"state":"rejected"}""");

                if (path.EndsWith("/queue", StringComparison.Ordinal))
                    return AnsweringHttpHandler.Json("""{"canPublish":true,"isAdministrator":true,"submissions":[]}""");

                return AnsweringHttpHandler.Refused();
            });

            typeof(TabMaintainer).GetField("thisClient", Any)!
                .SetValue(main, new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            typeof(TabMaintainer).GetField("thisSession", Any)!
                .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

            await (Task)typeof(TabMaintainer).GetMethod("DecideAsync", Any)!.Invoke(main, [ReviewDecisionKind.Reject])!;

            Assert.True(string.IsNullOrEmpty(comment.Text));

            // The next submission arrives and is opened.
            ReviewQueueRow next = Row(id: 16);
            Queue(main, next);
            Select(main, next);

            Assert.True(string.IsNullOrEmpty(comment.Text));

            // Nor does the sent comment come back if the rejected one is ever shown again.
            Select(main, row);
            Assert.True(string.IsNullOrEmpty(comment.Text));
        });
    }

    // A decision the server refused keeps the comment, so it can be sent again.
    [Fact]
    public async Task A_refused_rejection_keeps_the_comment()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row(id: 15);
            Queue(main, row);
            Select(main, row);
            main.ShowDetail(Detail(row));
            MaintainerTabHost.AddOverlay(main);

            TextBox comment = main.FindControl<TextBox>("DecisionCommentTextBox")!;
            comment.Text = "The schematic image is the wrong board revision.";

            var server = new AnsweringHttpHandler(_ => AnsweringHttpHandler.Refused());

            typeof(TabMaintainer).GetField("thisClient", Any)!
                .SetValue(main, new ReviewApiClient("https://review.invalid", new HttpClient(server)));
            typeof(TabMaintainer).GetField("thisSession", Any)!
                .SetValue(main, new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(1), 1, "dh@example.com", "Dennis"));

            await (Task)typeof(TabMaintainer).GetMethod("DecideAsync", Any)!.Invoke(main, [ReviewDecisionKind.Reject])!;

            Assert.Equal("The schematic image is the wrong board revision.", comment.Text);
        });
    }
}
