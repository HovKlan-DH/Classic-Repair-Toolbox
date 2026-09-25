using Avalonia;
using Avalonia.Input;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The label editor's actual behaviour: entering and leaving the mode, selecting rectangles,
// moving and resizing them, deleting, keyboard nudging, and undo/redo.
//
// This is the tab's largest untested area - TabSchematics.LabelEditor.cs and
// .LabelEditor.Interaction.cs were at ~1% between them, because everything in the editor is
// private and reached through pointer handlers a headless test cannot drive end to end.
//
// HOW THESE DRIVE IT, and what that does and does not prove:
//
// The editor's pointer handlers do two jobs - convert a pointer position into bitmap pixels
// (which needs a laid-out control and a decoded bitmap, neither available headlessly), then act
// on the result. These tests enter at the second job, through the `...ForTests` seams in
// TabSchematics.LabelEditor.TestSeams.cs, in the same bitmap-pixel space the real handlers hand
// over. So the EDITING behaviour below is the shipped behaviour; the pointer-to-pixel conversion
// ahead of it is still only verified by running the app.
//
// Board data comes from CurrentBoardDataOverrideForTests, and the "current schematic" from a
// real SchematicThumbnail placed in the thumbnail list - the same source GetCurrentSchematicName
// reads in the running app.
// ###########################################################################################
[Collection("HeadlessUi")]
public class LabelEditorInteractionTests
{
    private const string Schematic = "Board top";
    private const string OtherSchematic = "Board bottom";

    // Three rectangles on the editable schematic, deliberately far apart so a 1px snap threshold
    // cannot pull one onto another and confuse an assertion about a move or resize.
    private static readonly Rect U1Rect = new(100, 100, 50, 40);
    private static readonly Rect U2Rect = new(400, 400, 60, 30);
    private static readonly Rect U3Rect = new(800, 800, 20, 20);

    // -----------------------------------------------------------------------------------------
    // Entering and leaving the mode
    // -----------------------------------------------------------------------------------------

    // Entering loads a WORKING COPY of the current schematic's rows. It is a copy precisely so
    // Cancel can throw the edits away - see the cancel test below.
    [Fact]
    public void Entering_the_editor_loads_the_current_schematics_rectangles()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();

            tab.BeginLabelEditorModeForTests();

            Assert.True(tab.IsLabelEditorModeForTests);
            Assert.Equal(Schematic, tab.LabelEditorSchematicNameForTests);
            Assert.Equal(
                new[] { U1Rect, U2Rect, U3Rect },
                tab.LabelEditorWorkingRowsForTests.Select(row => row.Rect));
        });
    }

    // Only the CURRENT schematic's rows are loaded. The board carries rows for others, and
    // editing a rectangle that belongs to a schematic you are not looking at would be invisible.
    [Fact]
    public void Entering_the_editor_ignores_other_schematics_rectangles()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();

            tab.BeginLabelEditorModeForTests();

            Assert.DoesNotContain("U9", tab.LabelEditorWorkingRowsForTests.Select(row => row.BoardLabel));
        });
    }

    [Fact]
    public void Cancelling_the_editor_leaves_the_mode_and_drops_the_working_copy()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.CancelLabelEditorChangesForTests();

            Assert.False(tab.IsLabelEditorModeForTests);
            Assert.Empty(tab.LabelEditorWorkingRowsForTests);
        });
    }

    // The point of the working copy: an edit made in the editor and then cancelled must not have
    // reached the board data. Cancel is the user's escape hatch and has to be trustworthy.
    [Fact]
    public void Cancelling_after_a_move_leaves_the_board_data_untouched()
    {
        UiTest.Run(() =>
        {
            BoardData board = CreateBoard();
            TabSchematics tab = CreateTab(board);
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(25, 25));
            tab.CompleteLabelEditorDragForTests();

            tab.CancelLabelEditorChangesForTests();

            // Still the original X/Y as strings, exactly as loaded.
            var u1 = board.ComponentHighlights.First(h => h.BoardLabel == "U1");
            Assert.Equal("100", u1.X);
            Assert.Equal("100", u1.Y);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Selection
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void Selecting_a_rectangle_selects_exactly_that_one()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.SelectLabelEditorRowForTests(1);

            Assert.Equal(1, tab.SelectedLabelEditorCountForTests);
            Assert.True(tab.IsLabelEditorRowSelectedForTests(1));
            Assert.False(tab.IsLabelEditorRowSelectedForTests(0));
        });
    }

    // Selecting a different rectangle REPLACES the selection rather than adding to it - the
    // plain-click behaviour, as distinct from the toggle below.
    [Fact]
    public void Selecting_another_rectangle_replaces_the_previous_selection()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.SelectLabelEditorRowForTests(0);
            tab.SelectLabelEditorRowForTests(2);

            Assert.Equal(1, tab.SelectedLabelEditorCountForTests);
            Assert.True(tab.IsLabelEditorRowSelectedForTests(2));
            Assert.False(tab.IsLabelEditorRowSelectedForTests(0));
        });
    }

    [Fact]
    public void Toggling_builds_a_multi_selection_and_removes_from_it_again()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.ToggleLabelEditorRowForTests(0);
            tab.ToggleLabelEditorRowForTests(1);
            Assert.Equal(2, tab.SelectedLabelEditorCountForTests);

            tab.ToggleLabelEditorRowForTests(0);
            Assert.Equal(1, tab.SelectedLabelEditorCountForTests);
            Assert.True(tab.IsLabelEditorRowSelectedForTests(1));
        });
    }

    [Fact]
    public void Clearing_the_selection_deselects_everything()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();
            tab.ToggleLabelEditorRowForTests(0);
            tab.ToggleLabelEditorRowForTests(1);

            tab.ClearLabelEditorSelectionForTests();

            Assert.Equal(0, tab.SelectedLabelEditorCountForTests);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Moving
    // -----------------------------------------------------------------------------------------

    // A move translates the rectangle and must NOT resize it - the bug this guards against is a
    // move implemented as "set both corners from the delta", which stretches on every drag.
    [Fact]
    public void Dragging_a_rectangle_moves_it_without_changing_its_size()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(30, 20));
            tab.CompleteLabelEditorDragForTests();

            Rect moved = tab.LabelEditorWorkingRowsForTests[0].Rect;

            Assert.Equal(U1Rect.X + 30, moved.X);
            Assert.Equal(U1Rect.Y + 20, moved.Y);
            Assert.Equal(U1Rect.Width, moved.Width);
            Assert.Equal(U1Rect.Height, moved.Height);
        });
    }

    // The drag delta is measured from the DRAG START, not from the previous update - otherwise
    // intermediate pointer moves compound and the rectangle races away from the pointer.
    [Fact]
    public void Successive_drag_updates_do_not_compound()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(10, 10));
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(30, 20));
            tab.CompleteLabelEditorDragForTests();

            Rect moved = tab.LabelEditorWorkingRowsForTests[0].Rect;

            // 30/20 from the START, not 10+30 / 10+20.
            Assert.Equal(U1Rect.X + 30, moved.X);
            Assert.Equal(U1Rect.Y + 20, moved.Y);
        });
    }

    // A multi-selection moves as one: every member shifts by the same delta and their relative
    // layout is preserved.
    [Fact]
    public void Dragging_a_multi_selection_moves_every_member_by_the_same_delta()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.ToggleLabelEditorRowForTests(0);
            tab.ToggleLabelEditorRowForTests(1);

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(15, 25));
            tab.CompleteLabelEditorDragForTests();

            Rect first = tab.LabelEditorWorkingRowsForTests[0].Rect;
            Rect second = tab.LabelEditorWorkingRowsForTests[1].Rect;

            Assert.Equal(U1Rect.X + 15, first.X);
            Assert.Equal(U1Rect.Y + 25, first.Y);
            Assert.Equal(U2Rect.X + 15, second.X);
            Assert.Equal(U2Rect.Y + 25, second.Y);
        });
    }

    // A rectangle outside the selection must not move with it.
    [Fact]
    public void Dragging_a_selection_leaves_unselected_rectangles_alone()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.SelectLabelEditorRowForTests(0);
            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(40, 40));
            tab.CompleteLabelEditorDragForTests();

            Assert.Equal(U3Rect, tab.LabelEditorWorkingRowsForTests[2].Rect);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Resizing
    // -----------------------------------------------------------------------------------------

    // Dragging the left edge moves the LEFT edge and leaves the right one where it was. Getting
    // this wrong (moving the whole rect, or the opposite edge) is the classic resize bug.
    [Fact]
    public void Resizing_the_left_edge_moves_that_edge_and_leaves_the_right_one()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            double originalRight = U1Rect.Right;

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.ResizeLeft);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(10, 0));
            tab.CompleteLabelEditorDragForTests();

            Rect resized = tab.LabelEditorWorkingRowsForTests[0].Rect;

            Assert.Equal(U1Rect.X + 10, resized.X);
            Assert.Equal(originalRight, resized.Right);
            Assert.Equal(U1Rect.Height, resized.Height);
        });
    }

    [Fact]
    public void Resizing_the_bottom_edge_moves_that_edge_and_leaves_the_top_one()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.BottomLeft, LabelEditorDragMode.ResizeBottom);
            tab.UpdateLabelEditorDragForTests(U1Rect.BottomLeft + new Point(0, 15));
            tab.CompleteLabelEditorDragForTests();

            Rect resized = tab.LabelEditorWorkingRowsForTests[0].Rect;

            Assert.Equal(U1Rect.Y, resized.Y);
            Assert.Equal(U1Rect.Height + 15, resized.Height);
            Assert.Equal(U1Rect.X, resized.X);
        });
    }

    // A corner handle moves BOTH of its edges at once.
    [Fact]
    public void Resizing_a_corner_moves_both_of_its_edges()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.BottomRight, LabelEditorDragMode.ResizeBottomRight);
            tab.UpdateLabelEditorDragForTests(U1Rect.BottomRight + new Point(20, 10));
            tab.CompleteLabelEditorDragForTests();

            Rect resized = tab.LabelEditorWorkingRowsForTests[0].Rect;

            Assert.Equal(U1Rect.X, resized.X);
            Assert.Equal(U1Rect.Y, resized.Y);
            Assert.Equal(U1Rect.Width + 20, resized.Width);
            Assert.Equal(U1Rect.Height + 10, resized.Height);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Deleting
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void Deleting_a_rectangle_removes_only_that_one()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.DeleteLabelEditorRowForTests(1);

            Assert.Equal(
                new[] { U1Rect, U3Rect },
                tab.LabelEditorWorkingRowsForTests.Select(row => row.Rect));
        });
    }

    // -----------------------------------------------------------------------------------------
    // Keyboard nudging
    // -----------------------------------------------------------------------------------------

    // Arrow keys move by exactly 1px - the fine-adjustment path that a mouse cannot hit reliably.
    [Fact]
    public void An_arrow_key_nudges_the_selection_by_one_pixel()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();
            tab.SelectLabelEditorRowForTests(0);

            Assert.True(tab.ApplySelectedLabelEditorKeyboardStepForTests(Key.Right, KeyModifiers.None));

            Rect nudged = tab.LabelEditorWorkingRowsForTests[0].Rect;
            Assert.Equal(U1Rect.X + 1, nudged.X);
            Assert.Equal(U1Rect.Width, nudged.Width);
        });
    }

    // Shift+arrow EXPANDS in the pressed direction rather than moving.
    [Fact]
    public void Shift_and_an_arrow_key_expands_the_selection()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();
            tab.SelectLabelEditorRowForTests(0);

            Assert.True(tab.ApplySelectedLabelEditorKeyboardStepForTests(Key.Right, KeyModifiers.Shift));

            Rect expanded = tab.LabelEditorWorkingRowsForTests[0].Rect;
            Assert.Equal(U1Rect.Width + 1, expanded.Width);
        });
    }

    // With nothing selected there is nothing to nudge, and the key must be reported as unhandled
    // so it can fall through to whatever else wants it.
    [Fact]
    public void A_keyboard_step_with_nothing_selected_is_not_handled()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            Assert.False(tab.ApplySelectedLabelEditorKeyboardStepForTests(Key.Right, KeyModifiers.None));
        });
    }

    // Shift and Alt together are contradictory (expand and shrink at once), so the step is refused
    // rather than picking one arbitrarily.
    [Fact]
    public void Shift_and_alt_together_are_refused()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();
            tab.SelectLabelEditorRowForTests(0);

            Assert.False(tab.ApplySelectedLabelEditorKeyboardStepForTests(
                Key.Right, KeyModifiers.Shift | KeyModifiers.Alt));

            Assert.Equal(U1Rect, tab.LabelEditorWorkingRowsForTests[0].Rect);
        });
    }

    // Outside the editor the keyboard step must do nothing at all - arrow keys belong to panning
    // and thumbnail navigation there.
    [Fact]
    public void A_keyboard_step_outside_the_editor_is_not_handled()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();

            Assert.False(tab.ApplySelectedLabelEditorKeyboardStepForTests(Key.Right, KeyModifiers.None));
        });
    }

    // -----------------------------------------------------------------------------------------
    // Undo and redo
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void Undo_restores_the_rectangle_a_drag_moved()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(30, 30));
            tab.CompleteLabelEditorDragForTests();

            Assert.True(tab.TryUndoLabelEditorChangeForTests());
            Assert.Equal(U1Rect, tab.LabelEditorWorkingRowsForTests[0].Rect);
        });
    }

    [Fact]
    public void Redo_reapplies_an_undone_move()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(30, 30));
            tab.CompleteLabelEditorDragForTests();

            tab.TryUndoLabelEditorChangeForTests();
            Assert.True(tab.TryRedoLabelEditorChangeForTests());

            Rect redone = tab.LabelEditorWorkingRowsForTests[0].Rect;
            Assert.Equal(U1Rect.X + 30, redone.X);
            Assert.Equal(U1Rect.Y + 30, redone.Y);
        });
    }

    // Nothing to undo is reported as such rather than silently doing nothing, so the caller can
    // leave the keypress unhandled.
    [Fact]
    public void Undo_with_no_history_reports_that_it_did_nothing()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            Assert.False(tab.TryUndoLabelEditorChangeForTests());
        });
    }

    [Fact]
    public void Undo_outside_the_editor_does_nothing()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();

            Assert.False(tab.TryUndoLabelEditorChangeForTests());
        });
    }

    // Two edits undo in reverse order - the stack behaviour a user relies on when backing out of
    // several changes.
    [Fact]
    public void Two_moves_undo_in_reverse_order()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft + new Point(10, 0));
            tab.CompleteLabelEditorDragForTests();

            Rect afterFirst = tab.LabelEditorWorkingRowsForTests[0].Rect;

            tab.StartLabelEditorDragForTests(0, afterFirst.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(afterFirst.TopLeft + new Point(0, 10));
            tab.CompleteLabelEditorDragForTests();

            tab.TryUndoLabelEditorChangeForTests();
            Assert.Equal(afterFirst, tab.LabelEditorWorkingRowsForTests[0].Rect);

            tab.TryUndoLabelEditorChangeForTests();
            Assert.Equal(U1Rect, tab.LabelEditorWorkingRowsForTests[0].Rect);
        });
    }

    // A drag that changed nothing must not push an undo entry, or the user presses undo and
    // apparently nothing happens.
    [Fact]
    public void A_drag_that_moved_nothing_records_no_undo_step()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();
            tab.BeginLabelEditorModeForTests();

            tab.StartLabelEditorDragForTests(0, U1Rect.TopLeft, LabelEditorDragMode.Move);
            tab.UpdateLabelEditorDragForTests(U1Rect.TopLeft);
            tab.CompleteLabelEditorDragForTests();

            Assert.False(tab.TryUndoLabelEditorChangeForTests());
        });
    }

    // -----------------------------------------------------------------------------------------
    // Fixture
    // -----------------------------------------------------------------------------------------

    // Board data with three highlights on the editable schematic plus one on another, so the
    // "current schematic only" rule has something to exclude.
    private static BoardData CreateBoard()
    {
        var board = new BoardData();

        void AddHighlight(string schematic, string label, Rect rect)
        {
            board.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = schematic,
                BoardLabel = label,
                X = rect.X.ToString(System.Globalization.CultureInfo.InvariantCulture),
                Y = rect.Y.ToString(System.Globalization.CultureInfo.InvariantCulture),
                Width = rect.Width.ToString(System.Globalization.CultureInfo.InvariantCulture),
                Height = rect.Height.ToString(System.Globalization.CultureInfo.InvariantCulture),
            });
        }

        AddHighlight(Schematic, "U1", U1Rect);
        AddHighlight(Schematic, "U2", U2Rect);
        AddHighlight(Schematic, "U3", U3Rect);
        AddHighlight(OtherSchematic, "U9", new Rect(10, 10, 10, 10));

        return board;
    }

    // A tab whose "current schematic" is Schematic, via a real SchematicThumbnail in the
    // thumbnail list - the same place GetCurrentSchematicName reads it from in the running app.
    // No bitmap is needed: SchematicThumbnail carries Name independently of its image.
    private static TabSchematics CreateTab(BoardData? board = null)
    {
        var tab = new TabSchematics
        {
            CurrentBoardDataOverrideForTests = board ?? CreateBoard(),
        };

        var thumbnail = new SchematicThumbnail { Name = Schematic };
        tab.SchematicsThumbnailList.ItemsSource = new List<SchematicThumbnail> { thumbnail };
        tab.SchematicsThumbnailList.SelectedItem = thumbnail;

        return tab;
    }
}
