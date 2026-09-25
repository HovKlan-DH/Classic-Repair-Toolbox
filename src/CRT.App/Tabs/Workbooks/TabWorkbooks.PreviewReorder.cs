using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // DRAG-TO-REORDER for the board pane's schematic previews - the same gesture the full worklog
    // editor gives an entry's photos, applied to the previews in the middle panel, with the result
    // persisted in the workbook's own index.json (WorkbookRecord.SchematicOrder).
    //
    // WHERE THE DRAG MAY START, and why it matters: only OUTSIDE the schematic image - in practice
    // the caption above it and the preview's own padding. Inside the image, a press still selects
    // the schematic (and a press on a "#N" pill still opens its editor), because those are what the
    // image is for. A drag that could start anywhere would make it impossible to click a board
    // without risking a reorder, and the pills sit on top of the image as well.
    //
    // The grab area advertises itself: the caption carries a north/south resize cursor, so the one
    // place a drag starts is the one place the pointer changes shape. The image keeps the Hand
    // cursor that says "click me".
    //
    // WHAT THE USER SEES WHILE DRAGGING: the dragged preview is replaced in the list by an outlined
    // placeholder of the same height, which moves as the pointer crosses the midpoint of its
    // neighbours - so the gap always shows exactly where a release would drop it. This mirrors
    // WorklogEntryEditorWindow's photo reorder, deliberately: the two are the same gesture and
    // should not feel different.
    //
    // The pointer is CAPTURED on the preview for the duration, which is what lets the drag continue
    // over the images below it (each of which would otherwise take the pointer events itself) and
    // guarantees the release arrives here even if it happens outside the pane entirely.
    //
    // Part of the TabWorkbooks partial class - see TabWorkbooks.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // The gap between previews in BoardPreviewPanel. Mirrors the StackPanel's own Spacing="12"
        // in TabWorkbooks.axaml, and is needed here because ResolveDropIndex has to account for the
        // space between previews when working out which midpoint the pointer has passed. If that
        // markup value changes, this must change with it - a mismatch does not break the drag, it
        // just makes the drop land a slot early or late near the bottom of a long pane.
        private const double PreviewPanelSpacing = 12.0;

        // How far the pointer must move before a press becomes a drag rather than a click. Without
        // it, the tiny movement between button-down and button-up on an ordinary click starts a
        // drag, and the caption becomes impossible to click without reordering something.
        private const double PreviewDragThreshold = 4.0;

        // One north/south cursor for every caption this pane builds, for the reason the pane's
        // shared HandCursor exists: Cursor is IDisposable and holds an OS handle, and these are
        // rebuilt on every refresh.
        private static readonly Cursor PreviewDragCursor = new(StandardCursorType.SizeNorthSouth);

        // The preview currently being dragged, null when no drag is in progress. Holding the control
        // rather than the name because the placeholder has to be swapped in for this exact control
        // and back out again afterwards.
        private Border? thisDraggingPreview;

        // The schematic that preview shows, captured when the drag starts. Read at the drop rather
        // than re-derived: a refresh landing mid-drag can rebuild the pane underneath us, and the
        // name is what survives that.
        private string? thisDraggingSchematicName;

        // Where the press happened, to measure PreviewDragThreshold against.
        private Point thisPreviewDragStartPoint;

        // Whether the threshold has been passed and the placeholder is actually in the panel. False
        // between the press and the first real movement, which is the window in which the gesture is
        // still just a click.
        private bool thisIsDraggingPreview;

        // The stand-in shown where the dragged preview would land. Same height as the preview it
        // replaces, so nothing below it jumps as the drag begins.
        private Border? thisPreviewDropPlaceholder;

        // ###########################################################################################
        // Makes one built preview draggable. Called from BuildSchematicPreview once the preview
        // Border and its image layer both exist.
        //
        // THE GRAB AREA IS THE WHOLE PANEL EXCEPT THE IMAGE. Everything in the red-bordered panel -
        // the caption, the padding around it, the space beside and below the image - starts a drag;
        // the schematic image itself does not, because a press there selects the schematic (and a
        // press on a "#N" pill opens that worklog). The image is a far easier target than a one-line
        // caption, so excluding it and taking the rest is both a much bigger grab area and an
        // unambiguous split: the picture acts on the picture, the panel around it moves the panel.
        //
        // Only the PRESS is wired per preview. Everything after it (move, release, lost capture)
        // belongs to BoardPreviewPanel, wired once in EnsurePreviewDragHandlersOnPanel.
        //
        // WHY NOT ON THE PREVIEW ITSELF: starting a drag REMOVES the preview from the panel and puts
        // the placeholder in its place. A control taken out of the visual tree loses pointer capture
        // immediately and stops receiving pointer events, so a drag captured on the preview died the
        // instant it began - and the synchronous PointerCaptureLost that removal fires re-entered
        // the very swap that was still running, nulling the placeholder field mid-way and crashing
        // on "A null control cannot be added to a Controls collection". The panel stays put for the
        // whole gesture, so it is the thing that can hold the capture.
        // ###########################################################################################
        private void EnablePreviewReorder(Border preview, Control imageArea, string schematicName)
        {
            // Set on the PREVIEW and overridden on the image, rather than set only on a small grab
            // strip: the pointer must say "move me" across the whole draggable area, which is now
            // everything the panel covers apart from the picture.
            //
            // The image keeps the Hand the preview used to carry for its whole surface, because a
            // click there still selects the schematic. Two cursors, matching the two behaviours -
            // one cursor over an area that does two different things is what would mislead.
            //
            // While a SEARCH is filtering the pane, no drag is offered at all and the caption keeps
            // the ordinary arrow - see IsPreviewReorderAvailable for why reordering a filtered pane
            // cannot be allowed to persist anything.
            preview.Cursor = this.IsPreviewReorderAvailable() ? PreviewDragCursor : Cursor.Default;
            imageArea.Cursor = HandCursor;

            preview.PointerPressed += (_, e) =>
            {
                if (!e.GetCurrentPoint(preview).Properties.IsLeftButtonPressed)
                    return;

                // Re-tested at the press as well as at build time: the query is debounced, so the
                // pane on screen can have been built before the user finished typing.
                if (!this.IsPreviewReorderAvailable())
                    return;

                // A press that landed on the image (or on a pill over it) is a selection or a
                // worklog click, not a drag. Tested by hit-testing the SOURCE rather than by
                // comparing coordinates: the pills sit on their own canvas above the image and must
                // count as "inside" it too, and a bounds test would have to reproduce the layout to
                // know that.
                if (e.Source is Visual source && IsWithinPreviewImage(source, imageArea))
                    return;

                this.EnsurePreviewDragHandlersOnPanel();

                this.thisDraggingPreview = preview;
                this.thisDraggingSchematicName = schematicName;
                this.thisPreviewDragStartPoint = e.GetPosition(this.BoardPreviewPanel);
                this.thisIsDraggingPreview = false;

                e.Pointer.Capture(this.BoardPreviewPanel);

                // Stops the press also reaching the preview's own selection handler. Grabbing the
                // panel to move it is not a request to change which schematic the right-hand panel
                // is showing.
                e.Handled = true;
            };
        }

        // ###########################################################################################
        // Whether the previews may be reordered at all right now - false whenever the "Find a
        // previous repair" box is filtering the pane.
        //
        // WHY A SEARCH DISABLES IT OUTRIGHT: the order that gets stored is the one read back off the
        // panel, and WorkbookSchematicOrder.ApplyMove documents that names not currently shown are
        // DROPPED from the result. With a filter on, the panel holds only the schematics whose
        // worklogs matched, so one drag would replace the whole stored order with that subset and
        // permanently discard the hand-placed position of every schematic the filter had hidden -
        // silently, with no undo, and only noticed once the search was cleared again.
        //
        // Disabling beats the alternatives. Merging the filtered result back into the stored order
        // means deciding where each hidden name sits relative to a list the user has just
        // rearranged, which is exactly the guess ApplyMove's header refuses to make; and persisting
        // nothing while still animating the drag would show a reorder that quietly does not stick.
        // A drag simply does not start, and the caption keeps the plain arrow rather than the
        // north/south cursor that advertises one.
        // ###########################################################################################
        private bool IsPreviewReorderAvailable()
        {
            return this.thisSearchQuery.IsEmpty;
        }

        // ###########################################################################################
        // Whether a pressed control IS the preview's image layer or sits inside it - the one
        // boundary that separates "select this schematic" from "start dragging this panel". Both
        // handlers in BuildSchematicPreview test against this, so there is a single definition of
        // where the image ends rather than two that could disagree.
        //
        // Walks the VISUAL tree, which is what pointer events travel: the badge canvas and its pills
        // are visual children of the image layer, so a press on a pill correctly reports as inside
        // the image area.
        // ###########################################################################################
        internal static bool IsWithinPreviewImage(Visual candidate, Control ancestor)
        {
            for (Visual? current = candidate; current != null; current = current.GetVisualParent())
            {
                if (ReferenceEquals(current, ancestor))
                    return true;
            }

            return false;
        }

        // ###########################################################################################
        // Wires the move/release/lost-capture half of the gesture onto BoardPreviewPanel, once for
        // the life of the tab.
        //
        // ONCE, not per preview: the panel outlives every preview in it (RefreshBoardPreviews clears
        // and rebuilds its children on every board change, workbook switch and entry save), so
        // subscribing per preview would stack a fresh set of handlers on the same long-lived control
        // at every rebuild. The flag is the whole guard - these are never unsubscribed, because the
        // panel and this tab have the same lifetime.
        // ###########################################################################################
        private bool thisPreviewDragHandlersWired;

        private void EnsurePreviewDragHandlersOnPanel()
        {
            if (this.thisPreviewDragHandlersWired || this.BoardPreviewPanel == null)
                return;

            this.thisPreviewDragHandlersWired = true;

            this.BoardPreviewPanel.PointerMoved += (_, e) =>
            {
                var dragged = this.thisDraggingPreview;
                if (dragged == null)
                    return;

                var current = e.GetPosition(this.BoardPreviewPanel);

                if (!this.thisIsDraggingPreview)
                {
                    if (Math.Abs(current.Y - this.thisPreviewDragStartPoint.Y) < PreviewDragThreshold)
                        return;

                    if (!this.BeginPreviewDrag(dragged))
                    {
                        this.EndPreviewDrag(e.Pointer);
                        return;
                    }
                }

                this.MovePreviewPlaceholderTo(this.ResolvePreviewDropIndex(current.Y));
            };

            this.BoardPreviewPanel.PointerReleased += (_, e) =>
            {
                if (this.thisDraggingPreview == null)
                    return;

                // A press that never passed the threshold is a click on the caption, not a drag.
                // Nothing was moved and nothing is persisted.
                if (this.thisIsDraggingPreview)
                {
                    this.CommitPreviewDrop();
                }

                this.EndPreviewDrag(e.Pointer);
            };

            // A capture lost to anything else (another control taking the pointer, the window
            // deactivating, a refresh rebuilding the pane mid-drag) must not strand the placeholder
            // in the panel as a permanent empty slot - the same failure the editor's photo drag
            // guards against with its window-level handler.
            //
            // Now that the PANEL holds the capture this no longer fires as a side effect of the swap
            // itself, but BeginPreviewDrag stays re-entrancy-safe regardless: this handler runs
            // synchronously from whatever caused the loss, and it must never be able to null the
            // placeholder while that placeholder is being put into the panel.
            this.BoardPreviewPanel.PointerCaptureLost += (_, _) =>
            {
                if (this.thisDraggingPreview == null)
                    return;

                this.RemovePreviewPlaceholder();
                this.ResetPreviewDragState();
            };
        }

        // ###########################################################################################
        // TEST SEAMS for the drag, exposed because the gesture itself cannot be driven headlessly:
        // it needs real pointer capture, which Avalonia grants only to a control under a live input
        // device. These let a test run the STATE MACHINE - begin the drag, move the placeholder,
        // commit or abandon - against a real preview panel, which is where the reported crash was.
        //
        // Deliberately not a "simulate the whole gesture" helper: each step is what the real
        // handlers call, in the order they call it, so a test that passes here is exercising the
        // shipped path rather than a parallel one.
        // ###########################################################################################
        internal bool BeginPreviewDragForTests(Border preview) => this.BeginPreviewDrag(preview);

        internal void MovePreviewPlaceholderToForTests(int targetIndex) => this.MovePreviewPlaceholderTo(targetIndex);

        internal void CommitPreviewDropForTests() => this.CommitPreviewDrop();

        internal void RemovePreviewPlaceholderForTests() => this.RemovePreviewPlaceholder();

        internal void SetDraggingPreviewForTests(Border? preview, string? schematicName)
        {
            this.thisDraggingPreview = preview;
            this.thisDraggingSchematicName = schematicName;
        }

        internal Border? PreviewDropPlaceholderForTests => this.thisPreviewDropPlaceholder;

        // ###########################################################################################
        // Swaps the dragged preview out for a placeholder of the same height. Returns false when the
        // pane no longer holds it (a refresh landed between the press and the first movement), so
        // the caller can abandon the drag rather than working against a stale control.
        // ###########################################################################################
        private bool BeginPreviewDrag(Border preview)
        {
            if (this.BoardPreviewPanel == null)
                return false;

            // Refused outright while a search is filtering the pane - the panel then holds only the
            // matched schematics, and committing an order read off it would discard the stored
            // positions of everything the filter hid. See IsPreviewReorderAvailable.
            if (!this.IsPreviewReorderAvailable())
                return false;

            int index = this.BoardPreviewPanel.Children.IndexOf(preview);
            if (index < 0)
                return false;

            // A Rectangle inside the Border, not a dashed Border: only shapes carry StrokeDashArray,
            // Border has no dashed-edge option at all. Exactly what the editor's photo placeholder
            // does, and for the same reason - see WorklogEntryEditorWindow.axaml's own comment.
            var outline = new Avalonia.Controls.Shapes.Rectangle
            {
                Stroke = ResolveThemeBrushStatic("Text_Fail_Fg"),
                StrokeThickness = 2,
                StrokeDashArray = new Avalonia.Collections.AvaloniaList<double> { 4, 3 },
                RadiusX = 3,
                RadiusY = 3,
                Fill = ResolveThemeBrushStatic("Bg")
            };

            // The same instruction the editor's photo placeholder carries, worded for a schematic.
            // The gesture is not discoverable from the gap alone, and this is the moment the user
            // is looking straight at it.
            var hint = new TextBlock
            {
                Text = "Move the schematic up/down and release mouse button for new location",
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            };

            var placeholderBody = new Panel();
            placeholderBody.Children.Add(outline);
            placeholderBody.Children.Add(hint);

            var placeholder = new Border
            {
                Height = preview.Bounds.Height,
                Child = placeholderBody
            };

            // THROUGH A LOCAL, AND PUBLISHED TO THE FIELD LAST. Mutating Children can raise handlers
            // synchronously (a removal detaches the control, which drops pointer capture and fires
            // PointerCaptureLost right here), and those handlers null thisPreviewDropPlaceholder as
            // part of tearing a drag down. Assigning the field before the swap meant such a handler
            // could clear it BETWEEN the RemoveAt and the Insert, so the Insert was handed null -
            // the reported "A null control cannot be added to a Controls collection" crash. A local
            // cannot be cleared from underneath this method.
            this.BoardPreviewPanel.Children.RemoveAt(index);
            this.BoardPreviewPanel.Children.Insert(index, placeholder);

            this.thisPreviewDropPlaceholder = placeholder;
            this.thisIsDraggingPreview = true;
            return true;
        }

        // ###########################################################################################
        // Which slot the pointer is currently over, measured against the heights actually on screen.
        //
        // The heights include the placeholder, which is exactly as tall as the preview it replaced,
        // so the midpoints the maths walks are the ones the user can see. The arithmetic itself is
        // WorkbookSchematicOrder.ResolveDropIndex - pure and unit tested, because an off-by-one
        // there is the difference between a drop landing where the gap is and one slot away from it.
        // ###########################################################################################
        private int ResolvePreviewDropIndex(double pointerY)
        {
            var heights = this.BoardPreviewPanel.Children
                .Select(child => child.Bounds.Height)
                .ToList();

            return WorkbookSchematicOrder.ResolveDropIndex(heights, PreviewPanelSpacing, pointerY);
        }

        // ###########################################################################################
        // Moves the placeholder to the given slot, leaving the panel untouched when it is already
        // there - otherwise every pointer-move frame would remove and re-insert the same control,
        // re-running layout for nothing.
        // ###########################################################################################
        private void MovePreviewPlaceholderTo(int targetIndex)
        {
            if (this.thisPreviewDropPlaceholder == null || this.BoardPreviewPanel == null)
                return;

            // Held in a local across the swap, for the reason BeginPreviewDrag gives: a Children
            // mutation can raise a handler synchronously, and those handlers null this field.
            var placeholder = this.thisPreviewDropPlaceholder;

            int currentIndex = this.BoardPreviewPanel.Children.IndexOf(placeholder);
            if (currentIndex < 0)
                return;

            int clamped = Math.Clamp(targetIndex, 0, this.BoardPreviewPanel.Children.Count - 1);
            if (clamped == currentIndex)
                return;

            this.BoardPreviewPanel.Children.RemoveAt(currentIndex);
            this.BoardPreviewPanel.Children.Insert(clamped, placeholder);
        }

        // ###########################################################################################
        // Persists the new order and rebuilds the pane from it.
        //
        // The placeholder's index IS the drop target - it has been tracking the pointer all along,
        // so what the user sees is what gets stored, with no second calculation that could disagree
        // with the gap on screen.
        //
        // The names handed to ApplyMove are the ones currently DISPLAYED, read back off the panel in
        // panel order (the placeholder standing in for the dragged one), so the stored order is
        // exactly the arrangement on screen - see WorkbookSchematicOrder.ApplyMove for why it takes
        // the displayed order rather than combining the old stored order itself.
        // ###########################################################################################
        private void CommitPreviewDrop()
        {
            if (this.thisPreviewDropPlaceholder == null ||
                this.thisDraggingSchematicName == null ||
                this.thisSelectedWorkbookId <= 0)
            {
                return;
            }

            // Last line of defence for the filtered case IsPreviewReorderAvailable already refuses
            // to start a drag in: the query is debounced, so a search can land mid-gesture, and
            // persisting an order read off a filtered panel would discard the stored positions of
            // everything the filter hid. The pane is put back by the caller's rebuild.
            if (!this.IsPreviewReorderAvailable())
            {
                this.RemovePreviewPlaceholder();
                this.ResetPreviewDragState();
                this.StartFreshBoardPass();
                return;
            }

            int targetIndex = this.BoardPreviewPanel.Children.IndexOf(this.thisPreviewDropPlaceholder);
            if (targetIndex < 0)
                return;

            var displayedNames = this.BoardPreviewPanel.Children
                .Select(child => ReferenceEquals(child, this.thisPreviewDropPlaceholder)
                    ? this.thisDraggingSchematicName
                    : (child as Border)?.Tag as string)
                .Where(name => !string.IsNullOrEmpty(name))
                .Select(name => name!)
                .ToList();

            var newOrder = WorkbookSchematicOrder.ApplyMove(displayedNames, this.thisDraggingSchematicName, targetIndex);

            // Kept in step with disk BEFORE the rebuild below reads it. The pane resolves the
            // stored order off the header record rather than re-reading every workbook from disk
            // (see ResolveShownWorkbookRecord), and this path rebuilds through StartFreshBoardPass
            // rather than RefreshWorkbooks - so nothing else would refresh that record, and the
            // rebuild would put the previews straight back into the pre-drop order.
            if (this.thisHeaderWorkbook != null && this.thisHeaderWorkbook.Id == this.thisSelectedWorkbookId)
            {
                this.thisHeaderWorkbook.SchematicOrder = newOrder;
            }

            if (!WorklogManager.UpdateWorkbookSchematicOrder(this.thisSelectedWorkbookId, newOrder))
            {
                // The write failed, so the pane must not be left showing an order that is not on
                // disk - the rebuild below puts it back to the stored one.
                Logger.Warning($"Could not save the schematic order for workbook [#{this.thisSelectedWorkbookId}]");
            }

            // Removes the placeholder and puts the real previews back, now in the saved order. A
            // full pass rather than re-inserting the dragged control by hand: the previews carry
            // selection borders and badge layout that a rebuild gets right for free.
            this.RemovePreviewPlaceholder();
            this.ResetPreviewDragState();
            this.StartFreshBoardPass();
        }

        // ###########################################################################################
        // Takes the placeholder back out of the panel and puts the dragged preview where it sat, so
        // an abandoned drag leaves the pane exactly as it was. Safe to call when there is no
        // placeholder, which is every path that ends a gesture that never became a drag.
        // ###########################################################################################
        private void RemovePreviewPlaceholder()
        {
            if (this.thisPreviewDropPlaceholder == null || this.BoardPreviewPanel == null)
            {
                this.thisPreviewDropPlaceholder = null;
                return;
            }

            // Both controls read into locals, and the FIELD CLEARED FIRST, before any Children
            // mutation. A mutation can re-enter this method synchronously (see BeginPreviewDrag);
            // clearing first makes that re-entry a no-op via the null guard above, instead of a
            // second pass that removes the placeholder twice or inserts the preview twice.
            var placeholder = this.thisPreviewDropPlaceholder;
            var dragged = this.thisDraggingPreview;

            this.thisPreviewDropPlaceholder = null;

            int index = this.BoardPreviewPanel.Children.IndexOf(placeholder);
            if (index >= 0)
            {
                this.BoardPreviewPanel.Children.RemoveAt(index);

                if (dragged != null && !this.BoardPreviewPanel.Children.Contains(dragged))
                {
                    this.BoardPreviewPanel.Children.Insert(index, dragged);
                }
            }
        }

        // ###########################################################################################
        // Ends the gesture: releases the pointer capture and clears the drag state. Called on
        // release whether or not a drag actually happened.
        // ###########################################################################################
        private void EndPreviewDrag(IPointer pointer)
        {
            pointer.Capture(null);

            this.RemovePreviewPlaceholder();
            this.ResetPreviewDragState();
        }

        private void ResetPreviewDragState()
        {
            this.thisDraggingPreview = null;
            this.thisDraggingSchematicName = null;
            this.thisIsDraggingPreview = false;
        }
    }
}
