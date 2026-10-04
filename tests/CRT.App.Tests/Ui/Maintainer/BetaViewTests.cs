using Avalonia.LogicalTree;
using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The "BETA" screen's right-hand panel, BetaView - the "Publish to production" window's panel
// until the four screens replaced the windows (owner request, 2026-09-27).
//
// Its contents were reworked the same day ("I am not sure if the shown information in the
// right-side panel is any helpful ... can you propose something?"): it used to be the file copy
// list and nothing else - the mechanics of a copy, which cannot tell a maintainer whether the data
// is right, while the thing the button actually does (push named people's accepted work to every
// CRT user) was not on screen at all. The panel is now, in order: WHO this carries, what needs a
// second look, then the files - as a folder tree of production after the publish since 2026-09-28
// (owner request), which replaced the collapsed copy list.
//
// Built without a server and drawn through ShowPlanForTests; with no client it asks nobody.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BetaViewTests
{
    private static readonly DateTimeOffset Decided = new(2026, 9, 24, 8, 30, 0, TimeSpan.Zero);

    private static ProductionSystemRow Row() =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-September-25", "hash", "2026-May-14", null);

    private static ProductionPlanView Plan(
        IReadOnlyList<CarriedSubmission>? carrying = null,
        params PromotionFile[] files) =>
        new(
            "Commodore/C64/250407",
            "2026-September-25",
            new string('c', 64),
            TouchesSharedFiles: false,
            CanPublish: true,
            Refusal: null,
            UnchangedCount: 1180,
            Files: files,
            Problems: [],
            Approval: null,
            Removals: null,
            Carrying: carrying);

    // Sizes in the tree (owner request, 2026-10-04): every file's size from the plan's FileSizes - a
    // removed one as production holds it - and none for a file the plan does not size.
    [Fact]
    public void Each_file_in_the_tree_shows_the_size_the_plan_gives_it()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(carrying: [], files: [BetaViewTests.File("Commodore/C64/250407/Images/a.png")]) with
                {
                    UnchangedFiles = ["Commodore/C64/250407/Images/b.png"],
                    Removals = new FileRemovalPreview(["Commodore/C64/250407/Images/old.png"], null),
                    FileSizes = new Dictionary<string, long>
                    {
                        ["Commodore/C64/250407/Images/a.png"] = 2048,
                        ["Commodore/C64/250407/Images/old.png"] = 512
                    }
                },
                BetaViewTests.Row());

            FileTreeView tree = window.FileTreeForTests;
            tree.FindControl<CheckBox>("OnlyChangedCheckBox")!.IsChecked = false;

            Assert.Equal("2.0 KB", tree.RowsForTests.Single(row => row.Name == "a.png").Size);
            Assert.Equal("512 bytes", tree.RowsForTests.Single(row => row.Name == "old.png").Size);
            Assert.Equal(string.Empty, tree.RowsForTests.Single(row => row.Name == "b.png").Size);
        });
    }

    private static PromotionFile File(string path) =>
        new(path, new string('a', 64), PromotionChange.Added, PromotionStage.Content, IsShared: false);

    private static IEnumerable<string> TextOf(StackPanel panel) =>
        panel.Children.OfType<TextBlock>().Select(block => block.Text ?? string.Empty);

    [Fact]
    public void The_panel_builds()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            Assert.NotNull(window.FindControl<StackPanel>("CarryingPanel"));
            Assert.NotNull(window.FindControl<StackPanel>("AttentionPanel"));
            Assert.NotNull(window.FindControl<FileTreeView>("FileTree"));
        });
    }

    // ###########################################################################################
    // *** THE CONTRIBUTORS ARE NAMED, AND THEY COME FIRST. *** This is the question the button
    // asks and the panel could not previously answer.
    // ###########################################################################################
    [Fact]
    public void The_panel_names_whose_work_the_promotion_carries()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(
                    carrying: [new CarriedSubmission(7, "hest@mailscan.dk", "Corrected U8.", BetaViewTests.Decided)],
                    files: BetaViewTests.File("Commodore/C64/250407/sheet.png")),
                BetaViewTests.Row());

            List<string> lines = BetaViewTests.TextOf(
                window.FindControl<StackPanel>("CarryingPanel")!).ToList();

            Assert.Equal("1 contribution goes out to everyone with this:", lines[0]);
            Assert.Contains(lines, line => line.Contains("hest@mailscan.dk", StringComparison.Ordinal));
            Assert.Contains(lines, line => line.Contains("Corrected U8.", StringComparison.Ordinal));
        });
    }

    // Said even when there is none, so the space cannot read as "the question was not asked".
    [Fact]
    public void A_promotion_carrying_nothing_says_so()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(BetaViewTests.Plan(carrying: []), BetaViewTests.Row());

            Assert.Equal(
                "No contributions are waiting to go out with this.",
                BetaViewTests.TextOf(window.FindControl<StackPanel>("CarryingPanel")!).First());
        });
    }

    // ###########################################################################################
    // *** THE FILES ARE A TREE OF PRODUCTION AFTER THE PUBLISH (owner request, 2026-09-28). ***
    // Only what changes shows at first, each file under all its folders; unticking the box shows
    // everything the system has. It replaced the collapsed copy list, which named the copies and
    // nothing around them.
    // ###########################################################################################
    [Fact]
    public void The_files_show_as_a_tree_of_what_the_publish_changes()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(carrying: [], files: [BetaViewTests.File("Commodore/C64/250407/Images/a.png")]) with
                {
                    UnchangedFiles = ["Commodore/C64/250407/Images/b.png", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"],
                    Removals = new FileRemovalPreview(["Commodore/C64/250407/Images/old.png"], null)
                },
                BetaViewTests.Row());

            FileTreeView tree = window.FileTreeForTests;
            Assert.True(tree.IsVisible);

            // Only what changes: the three folders down to Images, then the new and the removed file.
            Assert.Equal(
                ["Commodore", "C64", "250407", "Images", "a.png", "old.png"],
                tree.RowsForTests.Select(row => row.Name));
            Assert.Equal("new", tree.RowsForTests.Single(row => row.Name == "a.png").PillText);
            Assert.Equal("removed", tree.RowsForTests.Single(row => row.Name == "old.png").PillText);

            // The counts are the panel's own line above it, not said twice.
            Assert.False(tree.ShowSummary);

            // No sizes in this plan (an older server): no size shown rather than a wrong one.
            Assert.All(tree.RowsForTests, row => Assert.Equal(string.Empty, row.Size));

            tree.FindControl<CheckBox>("OnlyChangedCheckBox")!.IsChecked = false;

            Assert.Contains(tree.RowsForTests, row => row.Name == "b.png");
            Assert.Contains(tree.RowsForTests, row => row.Name == "Data C64 250407 v2.0.0.xlsx");
        });
    }

    // Nothing to copy (production brought level by hand): the tree is there, and shows nothing
    // while only the changes are asked for.
    [Fact]
    public void A_promotion_with_nothing_to_copy_shows_an_empty_changes_tree()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(carrying: []) with { UnchangedFiles = ["Commodore/C64/250407/Images/b.png"] },
                BetaViewTests.Row());

            Assert.True(window.FileTreeForTests.IsVisible);
            Assert.Empty(window.FileTreeForTests.RowsForTests);
        });
    }

    // ###########################################################################################
    // *** THE TREE IS NEVER INSIDE THE SCROLLED COLUMN. *** A list measured with unlimited height
    // builds every one of a board's ~2,000 rows - the table's rule, for the same reason. The
    // details above it scroll on their own.
    // ###########################################################################################
    [Fact]
    public void The_tree_is_not_inside_a_ScrollViewer()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            Assert.Empty(window.FindControl<FileTreeView>("FileTree")!.GetLogicalAncestors().OfType<ScrollViewer>());
        });
    }

    // The counts belong WITH the list they count, under the contributors - not at the top.
    [Fact]
    public void The_file_counts_are_read_after_whose_work_this_carries()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(carrying: [], files: BetaViewTests.File("a.png")),
                BetaViewTests.Row());

            TextBlock summary = window.FindControl<TextBlock>("PlanSummaryText")!;
            StackPanel carrying = window.FindControl<StackPanel>("CarryingPanel")!;

            // Read off the TREE rather than off screen coordinates: the panel is not shown, so
            // nothing has been laid out and every Bounds is still zero. The order of the children IS
            // the reading order here.
            Control column = (Control)carrying.Parent!;
            List<Control> order = [.. ((Panel)column).Children];

            int carryingAt = order.IndexOf(carrying);
            int summaryAt = order.FindIndex(child => child == summary || (child is Panel panel && panel.Children.Contains(summary)));

            Assert.True(carryingAt >= 0 && summaryAt >= 0, "both sections must be in the scrolled column");
            Assert.True(summaryAt > carryingAt, "the file counts must be read after whose work this carries");
        });
    }

    // ###########################################################################################
    // *** A PLAN THAT GOES AWAY TAKES ITS FILES WITH IT (code review, 2026-09-27, about the list
    // the tree replaced). *** Refresh, a deselect or a failed plan request must never leave the
    // previous board's files on screen.
    // ###########################################################################################
    [Fact]
    public void Clearing_the_plan_clears_the_previous_boards_files()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(carrying: [], files: [BetaViewTests.File("a.png"), BetaViewTests.File("b.png")]),
                BetaViewTests.Row());

            Assert.Equal(2, window.FileTreeForTests.RowsForTests.Count);

            window.ShowPlanForTests(null, null);

            Assert.False(window.FileTreeForTests.IsVisible);
            Assert.Empty(window.FileTreeForTests.RowsForTests);
        });
    }

    // ###########################################################################################
    // *** PUSHING BACK NEEDS NO TICK (owner decision, 2026-09-27). *** The box says "I have checked
    // this board in BETA and it is right", which is the OPPOSITE of what pushing back means. A
    // board whose publish is blocked - a shared file waiting for the administrator, say - is
    // exactly one a maintainer may want to push back, so it must not depend on canPublish either.
    // ###########################################################################################
    [Fact]
    public void Push_back_is_offered_without_the_checked_in_beta_tick()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(carrying: [], files: BetaViewTests.File("a.png")),
                BetaViewTests.Row());

            // The tick is deliberately untouched, so the publish stays off.
            Assert.False(window.FindControl<CheckBox>("CheckedInBetaCheckBox")!.IsChecked);
            Assert.False(window.FindControl<Button>("PublishButton")!.IsEnabled);

            Assert.True(window.FindControl<Button>("RollBackButton")!.IsEnabled);

            // "Reject" (owner request, 2026-09-28) is the same roll back, so the same rule.
            Assert.True(window.FindControl<Button>("RejectButton")!.IsEnabled);
        });
    }

    // A board the server refuses to publish can still be pushed back - that is often WHY.
    [Fact]
    public void Push_back_is_offered_even_when_the_server_refuses_the_publish()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(
                BetaViewTests.Plan(carrying: [], files: BetaViewTests.File("a.png"))
                    with { CanPublish = false },
                BetaViewTests.Row());

            Assert.True(window.FindControl<Button>("RollBackButton")!.IsEnabled);
        });
    }

    // With no system selected there is nothing to push back.
    [Fact]
    public void Push_back_is_off_with_no_system_selected()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            window.ShowPlanForTests(null, null);

            Assert.False(window.FindControl<Button>("RollBackButton")!.IsEnabled);
            Assert.False(window.FindControl<Button>("RejectButton")!.IsEnabled);
        });
    }

    // ###########################################################################################
    // *** REMOVALS MOVED OUT OF THE FILE LIST. *** What a promotion DELETES from production is the
    // least visible change and the one that cannot be undone, so it belongs above the collapsed
    // list rather than inside it - where collapsing the list would have hidden it entirely.
    // ###########################################################################################
    [Fact]
    public void What_the_promotion_removes_is_shown_outside_the_collapsed_file_list()
    {
        UiTest.Run(() =>
        {
            var window = new BetaView();

            ProductionPlanView plan = BetaViewTests.Plan(carrying: [], files: BetaViewTests.File("a.png"))
                with { Removals = new FileRemovalPreview(["Commodore/C64/250407/old.png"], null) };

            window.ShowPlanForTests(plan, BetaViewTests.Row());

            Assert.Contains(
                BetaViewTests.TextOf(window.FindControl<StackPanel>("AttentionPanel")!),
                line => line.Contains("old.png", StringComparison.Ordinal));
        });
    }

    // ###########################################################################################
    // *** THE BUTTONS SHOW ONLY WITH A SYSTEM CHOSEN (2026-09-27). *** In a window of its own the
    // bar sat there disabled; beside a list it is the Review screen's rule instead - nothing chosen,
    // no decisions, and "Select a system" said where the name goes.
    // ###########################################################################################
    [Fact]
    public void With_nothing_chosen_the_buttons_are_hidden_and_the_panel_says_what_to_do()
    {
        UiTest.Run(() =>
        {
            var view = new BetaView();

            view.ShowPlanForTests(null, null);

            Assert.False(view.FindControl<StackPanel>("ActionPanel")!.IsVisible);
            Assert.Equal("Select a system", view.FindControl<TextBlock>("SelectedSystemText")!.Text);

            view.ShowPlanForTests(BetaViewTests.Plan(carrying: []), BetaViewTests.Row());

            Assert.True(view.FindControl<StackPanel>("ActionPanel")!.IsVisible);
            Assert.StartsWith("Commodore / C64 / 250407", view.FindControl<TextBlock>("SelectedSystemText")!.Text, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // *** WHAT HAPPENED IS SAID OUTSIDE THE BUTTON BAR. *** A publish usually takes the system out
    // of the list, which empties the panel and hides the bar - so a message inside the bar would be
    // hidden the moment it had something to say.
    // ###########################################################################################
    [Fact]
    public void A_message_is_read_even_when_the_panel_has_emptied()
    {
        UiTest.Run(() =>
        {
            var view = new BetaView();

            view.ShowPlanForTests(null, null);
            view.ShowMessage("Commodore/C64/250407 is published to production.", isError: false);

            TextBlock message = view.FindControl<TextBlock>("MessageText")!;

            Assert.True(message.IsVisible);
            Assert.DoesNotContain(message, view.FindControl<StackPanel>("ActionPanel")!.GetLogicalDescendants());
        });
    }

    // Somebody else published or pushed it back: the panel empties and says so, rather than keep
    // offering buttons for a BETA state that no longer waits.
    [Fact]
    public void A_system_gone_from_the_list_empties_the_panel_and_says_why()
    {
        UiTest.Run(() =>
        {
            var view = new BetaView();

            view.ShowPlanForTests(BetaViewTests.Plan(carrying: []), BetaViewTests.Row());
            view.ShowGoneElsewhere();

            Assert.Null(view.ShownRow);
            Assert.False(view.FindControl<StackPanel>("ActionPanel")!.IsVisible);
            Assert.Contains("no longer waiting to go to stable", view.FindControl<TextBlock>("MessageText")!.Text, StringComparison.Ordinal);
        });
    }
}
