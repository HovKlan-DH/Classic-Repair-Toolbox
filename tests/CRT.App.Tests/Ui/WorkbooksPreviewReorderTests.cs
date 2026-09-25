using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// Drag-to-reorder for the Workbooks board pane's schematic previews - the state machine behind the
// gesture, driven through TabWorkbooks' own seams.
//
// WHAT CANNOT BE TESTED HERE, and why these seams exist: the gesture starts with real pointer
// capture, which Avalonia grants only to a control under a live input device, so the pointer half
// (press on the caption, capture the panel, threshold, release) is verified by running the app.
// What IS pinned is everything the handlers then call - the placeholder swap, the moves, the commit
// and the teardown - which is where the reported crash was.
//
// THE CRASH THIS FILE EXISTS FOR: "A null control cannot be added to a Controls collection", thrown
// from BeginPreviewDrag the moment a preview was dragged upward. Removing the dragged preview from
// the panel detached it, which dropped the pointer capture it was holding and fired
// PointerCaptureLost SYNCHRONOUSLY, in the middle of the swap; that handler nulled the placeholder
// field, and the Insert immediately after was then handed null. Two things fixed it and both are
// pinned below: the capture and the move/release handlers moved to BoardPreviewPanel (which stays
// in the tree for the whole gesture), and every Children mutation now holds its control in a local
// rather than re-reading a field that a re-entrant handler can clear.
//
// COLLECTION NOTE: "HeadlessUi", like the other board-pane tests - constructing a control needs the
// shared dispatcher thread, and every assertion runs inside UiTest.Run.
[Collection("HeadlessUi")]
public sealed class WorkbooksPreviewReorderTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private readonly string thisBoardKey = "Commodore 64|250469 " + Guid.NewGuid().ToString("N");

    public void Dispose()
    {
        this.LoadWorklog();
        this.thisWorkspace.Dispose();
    }

    private string LoadWorklog()
    {
        string root = this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(root);
        return root;
    }

    private string WriteSchematicImage(string fileName)
    {
        string path = this.thisWorkspace.Path_(fileName);

        using var renderTarget = new RenderTargetBitmap(new PixelSize(40, 40), new Vector(96, 96));
        using (var context = renderTarget.CreateDrawingContext())
        {
            context.DrawRectangle(Brushes.White, null, new Rect(0, 0, 40, 40));
        }

        using (var stream = File.Create(path))
        {
            renderTarget.Save(stream, PngBitmapEncoderOptions.Default);
        }

        return path;
    }

    /// <summary>
    /// A workbook with one worklog on each of three schematics, so the pane builds three previews -
    /// the minimum for a move that is neither to the top nor to the bottom.
    /// </summary>
    private (TabWorkbooks Tab, int WorkbookId) BuildTabWithThreePreviews()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Recap", "");
        Assert.NotNull(workbook);

        var schematics = new List<BoardSchematicEntry>();

        foreach (string name in new[] { "Alpha", "Bravo", "Charlie" })
        {
            schematics.Add(new BoardSchematicEntry
            {
                SchematicName = name,
                SchematicImageFile = this.WriteSchematicImage($"{name}.png")
            });

            WorklogManager.AddEntry(
                workbook!.Id, name, new Rect(0, 0, 10, 10), $"{name} fault", "", "Issue", "Open", Array.Empty<string>());
        }

        var tab = new TabWorkbooks
        {
            BoardKeyOverrideForTests = this.thisBoardKey,
            CurrentBoardDataOverrideForTests = new BoardData { Schematics = schematics }
        };

        tab.ActivateWorkbookOverrideForTests = (boardKey, workbookId) =>
        {
            UserSettings.SetActiveWorkbookId(boardKey, workbookId);
            tab.RefreshWorkbooks();
        };

        tab.RefreshWorkbooks();
        tab.SelectWorkbookForTests(workbook!.Id);

        // Shown in a real window so the previews are actually arranged - BeginPreviewDrag sizes the
        // placeholder from the dragged preview's Bounds.Height, which is 0 without a layout pass.
        var window = new Window { Width = 800, Height = 600, Content = tab };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (tab, workbook.Id);
    }

    private static StackPanel PreviewPanel(TabWorkbooks tab) =>
        tab.GetControl<StackPanel>("BoardPreviewPanel");

    private static List<string> ShownSchematicNames(TabWorkbooks tab) =>
        PreviewPanel(tab).Children
            .OfType<Border>()
            .Select(border => border.Tag as string)
            .Where(name => name != null)
            .Select(name => name!)
            .ToList();

    // THE REPORTED CRASH, reduced to the call that threw. Beginning a drag swaps the dragged preview
    // out for the placeholder; against the pre-fix version this threw ArgumentNullException from
    // Children.Insert, because the removal re-entered the teardown and nulled the field the Insert
    // then read.
    [Fact]
    public void Beginning_a_drag_swaps_in_a_placeholder_without_throwing()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var panel = PreviewPanel(tab);
            Assert.Equal(3, panel.Children.Count);

            var dragged = (Border)panel.Children[0];
            tab.SetDraggingPreviewForTests(dragged, "Alpha");

            Assert.True(tab.BeginPreviewDragForTests(dragged));

            // The placeholder stands exactly where the preview was, so nothing below it moves.
            Assert.Equal(3, panel.Children.Count);
            Assert.NotNull(tab.PreviewDropPlaceholderForTests);
            Assert.Same(tab.PreviewDropPlaceholderForTests, panel.Children[0]);
            Assert.DoesNotContain(dragged, panel.Children);
        });
    }

    // The placeholder is sized to the preview it replaced, or everything below it jumps by the
    // difference the moment the drag begins.
    [Fact]
    public void The_placeholder_matches_the_height_of_the_preview_it_replaced()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var dragged = (Border)PreviewPanel(tab).Children[0];
            double draggedHeight = dragged.Bounds.Height;

            Assert.True(draggedHeight > 0, "the preview was never arranged - the window layout pass did not run");

            tab.SetDraggingPreviewForTests(dragged, "Alpha");
            tab.BeginPreviewDragForTests(dragged);

            Assert.Equal(draggedHeight, tab.PreviewDropPlaceholderForTests!.Height);
        });
    }

    [Fact]
    public void Moving_the_placeholder_puts_it_at_the_requested_slot()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var panel = PreviewPanel(tab);
            var dragged = (Border)panel.Children[0];

            tab.SetDraggingPreviewForTests(dragged, "Alpha");
            tab.BeginPreviewDragForTests(dragged);

            tab.MovePreviewPlaceholderToForTests(2);

            Assert.Same(tab.PreviewDropPlaceholderForTests, panel.Children[2]);
            Assert.Equal(3, panel.Children.Count);
        });
    }

    // Dropping persists the new order and the pane comes back showing it. This is the whole feature
    // end to end, short of the pointer input itself.
    [Fact]
    public void Dropping_a_preview_saves_the_new_order_and_redraws_the_pane_in_it()
    {
        UiTest.Run(() =>
        {
            var (tab, workbookId) = this.BuildTabWithThreePreviews();

            // Alphabetical to begin with - nothing has been dragged yet.
            Assert.Equal(new[] { "Alpha", "Bravo", "Charlie" }, ShownSchematicNames(tab));

            var dragged = (Border)PreviewPanel(tab).Children[2];
            tab.SetDraggingPreviewForTests(dragged, "Charlie");

            tab.BeginPreviewDragForTests(dragged);
            tab.MovePreviewPlaceholderToForTests(0);
            tab.CommitPreviewDropForTests();

            Dispatcher.UIThread.RunJobs();

            Assert.Equal(new[] { "Charlie", "Alpha", "Bravo" }, ShownSchematicNames(tab));

            // And on disk, in the workbook's own index.json - so it survives to the next session.
            var stored = WorklogManager.GetWorkbooksForBoard(this.thisBoardKey).Single(w => w.Id == workbookId);
            Assert.Equal(new[] { "Charlie", "Alpha", "Bravo" }, stored.SchematicOrder);
        });
    }

    // An abandoned drag (capture lost to another control, the window deactivating, a refresh
    // rebuilding the pane) must put the preview back where it was and leave the placeholder nowhere
    // in the panel - a stranded placeholder is a permanent empty slot.
    [Fact]
    public void Abandoning_a_drag_restores_the_preview_and_removes_the_placeholder()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var panel = PreviewPanel(tab);
            var dragged = (Border)panel.Children[0];

            tab.SetDraggingPreviewForTests(dragged, "Alpha");
            tab.BeginPreviewDragForTests(dragged);
            tab.MovePreviewPlaceholderToForTests(2);

            tab.RemovePreviewPlaceholderForTests();

            Assert.Null(tab.PreviewDropPlaceholderForTests);
            Assert.Equal(3, panel.Children.Count);
            Assert.Contains(dragged, panel.Children);
            Assert.Equal(new[] { "Alpha", "Bravo", "Charlie" }, ShownSchematicNames(tab).OrderBy(n => n).ToArray());
        });
    }

    // The teardown is reached from several paths (release, lost capture, and a commit that has
    // already run it), and a second pass must not remove a preview or double-insert one. It clears
    // its own field FIRST precisely so a re-entrant call is a no-op.
    [Fact]
    public void Tearing_down_a_drag_twice_leaves_the_pane_intact()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var panel = PreviewPanel(tab);
            var dragged = (Border)panel.Children[1];

            tab.SetDraggingPreviewForTests(dragged, "Bravo");
            tab.BeginPreviewDragForTests(dragged);

            tab.RemovePreviewPlaceholderForTests();
            tab.RemovePreviewPlaceholderForTests();

            Assert.Equal(3, panel.Children.Count);
            Assert.Equal(1, panel.Children.Count(child => ReferenceEquals(child, dragged)));
        });
    }

    // A drag that never began (the press never passed the movement threshold) has no placeholder, so
    // tearing it down must do nothing at all rather than reaching into the panel.
    [Fact]
    public void Tearing_down_a_gesture_that_never_became_a_drag_changes_nothing()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var panel = PreviewPanel(tab);
            var dragged = (Border)panel.Children[0];

            tab.SetDraggingPreviewForTests(dragged, "Alpha");
            tab.RemovePreviewPlaceholderForTests();

            Assert.Equal(3, panel.Children.Count);
            Assert.Equal(new[] { "Alpha", "Bravo", "Charlie" }, ShownSchematicNames(tab));
        });
    }

    // A refresh can rebuild the pane between the press and the first movement, leaving the captured
    // preview no longer in the panel. Beginning a drag against it must report failure so the caller
    // abandons the gesture, rather than working against a control that is not on screen.
    [Fact]
    public void Beginning_a_drag_on_a_preview_the_pane_no_longer_holds_reports_failure()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var stale = new Border();
            tab.SetDraggingPreviewForTests(stale, "Alpha");

            Assert.False(tab.BeginPreviewDragForTests(stale));
            Assert.Null(tab.PreviewDropPlaceholderForTests);
            Assert.Equal(3, PreviewPanel(tab).Children.Count);
        });
    }
    // ---------------------------------------------------------------------------------------------
    // THE GRAB AREA: what starts a drag and what selects the schematic.
    //
    // The whole red-bordered panel drags EXCEPT the schematic image, which selects instead. Both
    // behaviours hang off the same preview Border, so neither can rely on e.Handled to exclude the
    // other (Avalonia runs every handler registered on one element regardless) - each tests the
    // press position against the image layer instead, through the one shared IsWithinPreviewImage.
    //
    // The PRESS ITSELF still needs a live input device, so what is pinned here is that boundary
    // helper and the cursors that advertise it. The rest of the gesture is covered above.
    // ---------------------------------------------------------------------------------------------

    // The caption reads as an ordinary label, not a heading - it names the schematic beside two
    // other panels that already carry real headings.
    [Fact]
    public void The_schematic_caption_is_not_bold()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var caption = ((Control)PreviewPanel(tab).Children[0])
                .GetSelfAndVisualDescendants()
                .OfType<TextBlock>()
                .First(block => block.Text == "Alpha");

            Assert.NotEqual(FontWeight.Bold, caption.FontWeight);
        });
    }

    // The cursor is what tells the user the panel can be dragged, so it must cover the draggable
    // area - all of it except the image, which keeps the Hand that says "click to select".
    [Fact]
    public void The_panel_shows_a_move_cursor_while_the_image_keeps_the_click_cursor()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var preview = (Border)PreviewPanel(tab).Children[0];

            // Cursor has no value equality, so these are compared by their string form - the same
            // reason WorkbooksBoardPreviewTests compares cursors that way.
            Assert.Equal(
                new Cursor(StandardCursorType.SizeNorthSouth).ToString(),
                preview.Cursor?.ToString());

            var imageLayer = FindImageLayer(preview);

            Assert.Equal(
                new Cursor(StandardCursorType.Hand).ToString(),
                imageLayer.Cursor?.ToString());
        });
    }

    // The boundary both handlers share. A press on the caption is outside the image (so it drags),
    // a press on the image is inside it (so it selects), and - the case a naive bounds test gets
    // wrong - a press on a "#N" pill counts as INSIDE, because the pills sit on their own canvas
    // above the image and must never start a drag.
    [Fact]
    public void The_image_boundary_separates_dragging_from_selecting()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            var preview = (Border)PreviewPanel(tab).Children[0];
            var imageLayer = FindImageLayer(preview);

            var caption = preview.GetSelfAndVisualDescendants()
                .OfType<TextBlock>()
                .First(block => block.Text == "Alpha");

            var image = imageLayer.GetSelfAndVisualDescendants().OfType<Image>().First();

            Assert.False(TabWorkbooks.IsWithinPreviewImage(caption, imageLayer), "the caption drags the panel");
            Assert.True(TabWorkbooks.IsWithinPreviewImage(image, imageLayer), "the image selects the schematic");
            Assert.True(TabWorkbooks.IsWithinPreviewImage(imageLayer, imageLayer), "the layer itself counts as inside");

            // A pill, reached through the badge canvas that sits over the image. Asserted
            // unconditionally (each of these three worklogs draws one), so a fixture that stopped
            // producing pills fails here rather than quietly skipping the case this test is for.
            var pill = imageLayer.GetSelfAndVisualDescendants()
                .OfType<Canvas>()
                .SelectMany(canvas => canvas.Children.OfType<Border>())
                .FirstOrDefault();

            Assert.NotNull(pill);
            Assert.True(TabWorkbooks.IsWithinPreviewImage(pill!, imageLayer), "a worklog pill must never start a drag");
        });
    }

    // The image layer is the Grid holding the Image plus its overlay and badge canvas - the same
    // control BuildSchematicPreview hands to both handlers as the boundary.
    private static Control FindImageLayer(Border preview) =>
        preview.GetSelfAndVisualDescendants()
            .OfType<Grid>()
            .First(grid => grid.GetSelfAndVisualDescendants().OfType<Image>().Any());

    // ---------------------------------------------------------------------------------------------
    // THE FOUR HEADER ACTION BUTTONS (Edit/Delete workbook, Export to PDF/ZIP) and their shared
    // width.
    //
    // They line up on the same edges across the two rows, which means one width for all four. That
    // width used to be a hardcoded Width="100" in the markup, and it was too narrow for the widest
    // label: "Delete workbook" rendered as "Delete workboo", clipped with nothing reporting it. It
    // is now measured from whichever button is genuinely widest.
    // ---------------------------------------------------------------------------------------------

    private static Button[] HeaderActionButtons(TabWorkbooks tab) => new[]
    {
        tab.GetControl<Button>("EditWorkbookButton"),
        tab.GetControl<Button>("DeleteWorkbookButton"),
        tab.GetControl<Button>("ExportWorkbookButton"),
        tab.GetControl<Button>("ExportWorkbookZipButton")
    };

    // THE REPORTED BUG. Every label must fit inside the shared width - a button narrower than the
    // text it holds silently cuts the text off, so this compares the imposed width against what each
    // button's own content actually asks for.
    [Fact]
    public void Every_header_action_button_is_wide_enough_for_its_own_label()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();
            tab.EqualiseHeaderActionButtonWidthsForTests();

            foreach (var button in HeaderActionButtons(tab))
            {
                double shared = button.Width;

                // What this button would need if it sized itself: measured at auto, then restored.
                button.Width = double.NaN;
                button.Measure(Size.Infinity);
                double natural = button.DesiredSize.Width;
                button.Width = shared;

                Assert.True(
                    shared >= natural,
                    $"[{button.Content}] is {shared} wide but needs {natural} - its label is being clipped");
            }
        });
    }

    // The other half of the requirement: one width for all four, so their edges line up across the
    // two rows. Sizing each to its own label would fix the clipping and leave the rows ragged.
    [Fact]
    public void The_four_header_action_buttons_share_one_width()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();
            tab.EqualiseHeaderActionButtonWidthsForTests();

            var widths = HeaderActionButtons(tab).Select(button => button.Width).ToList();

            Assert.All(widths, width => Assert.False(double.IsNaN(width), "the shared width was never applied"));
            Assert.Single(widths.Distinct());
        });
    }

    // The shared width is the WIDEST button's, not an average or the first one's - which is what
    // makes "fits every label" and "all four equal" hold at the same time.
    [Fact]
    public void The_shared_width_is_the_widest_buttons_own_width()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            // Natural widths first, with no shared width imposed.
            var buttons = HeaderActionButtons(tab);
            double widestNatural = 0;

            foreach (var button in buttons)
            {
                button.Width = double.NaN;
                button.Measure(Size.Infinity);
                widestNatural = Math.Max(widestNatural, button.DesiredSize.Width);
            }

            tab.EqualiseHeaderActionButtonWidthsForTests();

            Assert.Equal(widestNatural, buttons[0].Width, precision: 3);
        });
    }

    // Equalising twice must not grow the buttons: each pass clears the width back to auto before
    // measuring, so it measures the LABEL rather than the width the previous pass imposed. Without
    // that reset the buttons would ratchet wider on every tab switch, since this runs on attach.
    [Fact]
    public void Equalising_twice_does_not_widen_the_buttons()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            tab.EqualiseHeaderActionButtonWidthsForTests();
            double afterFirst = HeaderActionButtons(tab)[0].Width;

            tab.EqualiseHeaderActionButtonWidthsForTests();
            tab.EqualiseHeaderActionButtonWidthsForTests();

            Assert.Equal(afterFirst, HeaderActionButtons(tab)[0].Width, precision: 3);
        });
    }

    // ###########################################################################################
    // REORDERING IS REFUSED WHILE A SEARCH IS FILTERING THE PANE.
    //
    // The stored order is read back off the panel, and WorkbookSchematicOrder.ApplyMove drops any
    // name that is not currently shown. With a filter on, the panel holds only the schematics whose
    // worklogs matched - so one drag would replace the whole stored order with that subset and
    // permanently discard the hand-placed position of every schematic the filter had hidden,
    // silently, with no undo, and only visible once the search was cleared again.
    //
    // Both tests below fail against the unguarded version: the first because the drag would start,
    // the second because the commit would write the two-name subset over the three-name order.
    // ###########################################################################################
    [Fact]
    public void A_drag_does_not_start_while_a_search_is_filtering_the_pane()
    {
        UiTest.Run(() =>
        {
            var (tab, _) = this.BuildTabWithThreePreviews();

            // Matches only the Bravo worklog's title, so the pane narrows to that one schematic.
            tab.GetControl<TextBox>("FindRepairTextBox").Text = "Bravo";
            tab.RefreshWorkbooks();
            Dispatcher.UIThread.RunJobs();

            var panel = PreviewPanel(tab);
            var dragged = panel.Children.OfType<Border>().First();

            tab.SetDraggingPreviewForTests(dragged, dragged.Tag as string);

            Assert.False(tab.BeginPreviewDragForTests(dragged));
            Assert.Null(tab.PreviewDropPlaceholderForTests);
        });
    }

    // The commit is guarded as well as the start, because the query is DEBOUNCED - a search can land
    // in the middle of a gesture that legitimately began on an unfiltered pane. The stored order
    // must survive that untouched.
    [Fact]
    public void A_drop_that_lands_while_a_search_is_active_does_not_rewrite_the_stored_order()
    {
        UiTest.Run(() =>
        {
            var (tab, workbookId) = this.BuildTabWithThreePreviews();

            // A real hand-placed order first, so there is something for a filtered commit to lose.
            var dragged = (Border)PreviewPanel(tab).Children[2];
            tab.SetDraggingPreviewForTests(dragged, "Charlie");
            tab.BeginPreviewDragForTests(dragged);
            tab.MovePreviewPlaceholderToForTests(0);
            tab.CommitPreviewDropForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(
                new[] { "Charlie", "Alpha", "Bravo" },
                WorklogManager.GetWorkbooksForBoard(this.thisBoardKey).Single(w => w.Id == workbookId).SchematicOrder);

            // Now begin a drag on the unfiltered pane, then let a search land before the drop.
            var second = (Border)PreviewPanel(tab).Children[0];
            tab.SetDraggingPreviewForTests(second, second.Tag as string);
            Assert.True(tab.BeginPreviewDragForTests(second));

            tab.GetControl<TextBox>("FindRepairTextBox").Text = "Bravo";
            tab.RefreshWorkbooks();
            Dispatcher.UIThread.RunJobs();

            tab.CommitPreviewDropForTests();
            Dispatcher.UIThread.RunJobs();

            // All three names still there, in the order the user actually chose - not the filtered
            // subset the panel was showing at the moment of the drop.
            Assert.Equal(
                new[] { "Charlie", "Alpha", "Bravo" },
                WorklogManager.GetWorkbooksForBoard(this.thisBoardKey).Single(w => w.Id == workbookId).SchematicOrder);
        });
    }

    // ###########################################################################################
    // THE SHARED WIDTH IS APPLIED EVEN WHEN THE TAB FIRST ATTACHES WITH NOTHING SELECTED.
    //
    // The equalisation runs on attach, immediately after a RefreshWorkbooks that HIDES the actions
    // panel whenever no workbook is selected - the ordinary state on a board with no workbooks, and
    // in AllBoards scope. Hidden buttons have no applied template, so that pass measures zero, bails
    // on its own "widest <= 0" guard and leaves all four at auto for the rest of the session, unless
    // the user happens to switch tabs and back. Selecting a workbook then has to retry it.
    //
    // Fails against the version that ran the equalisation only from OnAttachedToVisualTree.
    // ###########################################################################################
    [Fact]
    public void Selecting_a_workbook_on_a_board_that_had_none_still_equalises_the_action_buttons()
    {
        UiTest.Run(() =>
        {
            this.LoadWorklog();

            // A board with NO workbooks: the header actions panel is hidden on the first refresh.
            var tab = new TabWorkbooks
            {
                BoardKeyOverrideForTests = this.thisBoardKey,
                CurrentBoardDataOverrideForTests = new BoardData()
            };

            var window = new Window { Width = 800, Height = 600, Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Assert.False(tab.GetControl<StackPanel>("WorkbookHeaderActionsPanel").IsVisible);

            // Now a workbook exists, so the panel comes back - and the widths must follow.
            var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Recap", "");
            Assert.NotNull(workbook);

            tab.RefreshWorkbooks();
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.GetControl<StackPanel>("WorkbookHeaderActionsPanel").IsVisible);

            var widths = HeaderActionButtons(tab).Select(button => button.Width).ToList();

            Assert.All(widths, width => Assert.False(double.IsNaN(width), "a button was left at auto width"));
            Assert.All(widths, width => Assert.True(width > 0));

            // One width, not four - which is the whole point of the pass.
            Assert.Single(widths.Distinct());
        });
    }
}
