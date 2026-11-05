using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE PREVIEWS ON A SCHEMATIC IMAGE'S CARD (owner request, 2026-10-29: "I love that I live can see
    // the result of changing 'Marked components', but I would like to see the same for the 'On the
    // thumbnail' and on the 'Trace on the other side' ... maybe to zoom in/out"; owner's pick of the
    // three ways offered: zoom, a thumbnail-sized preview and the traces).
    //
    //   - THE PREVIEWS' WIDTH is dragged at the handle beside them on any card, and every card follows
    //     (PreviewWidth). The host keeps it per board (PreviewWidthChosen - TabDrafts.Splitters.cs).
    //   - ZOOMED IN, THE IMAGE IS DECODED SHARPER (LoadSharperAsync): in steps up to SharpestWidth
    //     (PicturePreviewGeometry.SharpWidth), off the UI thread, ONE card at a time - another card's
    //     sharp image goes back to its small one - and ONE DECODE PER CARD AT A TIME, the widest asked
    //     meanwhile decoded once it ends (code review, 2026-10-30: each zoom step started one). Never
    //     disposed by hand once drawn: one still drawn takes the render thread down, so the collector
    //     lets it go - only a decode never shown is disposed at once.
    //   - "TRACES ON THE PREVIEW": THIS SIDE (in the Selected components colour, as the Schematics tab
    //     draws a lit net's own-side copper) and THE OTHER SIDE (in the trace colour), each its own tick,
    //     both drawn when both are ticked (owner request, 2026-10-29: "so we can have both in preview").
    //     Both sides are read at once, by the host as "Mark all automatically" reads the board
    //     (ReadPreviewTraces - KiCadPreviewTraces placing them). The ticks are kept by the image's name
    //     AND its file while the part is open (two images may share a name - SchematicRows.RefOf), so
    //     cards made again (an image added, moved, deleted, the draft read again after an alignment)
    //     read their traces again - once for both - unless the image can no longer show them. A read
    //     answers only the latest one asked for its card (TracesAsk): another view picked or the image
    //     renamed meanwhile reads again, and an older read ending last never wins.
    //   - AN IMAGE WITH NO KICAD DATA has its KiCad options greyed out (owner request, 2026-10-29) - the
    //     row's HasKiCadData and CanShowTraces, the reason in the tooltip.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        private double thisPreviewWidth = ContributeWidths.DefaultPreviewWidth;

        // The width the cards have - 0 until they are laid out (ContributeWidths.PreviewShown).
        private double thisCardsWidth;

        // The card whose image is decoded sharp now - one at a time.
        private ComponentPaneBoardPictureRow? thisSharpRow;

        // Where a drag of a previews' handle started, and the width then.
        private (double FromX, double Width)? thisPreviewDrag;

        // The images whose traces are ticked, by name and file (TicksKey), with which sides - kept while
        // the part is open.
        private readonly Dictionary<string, (bool ThisSide, bool OtherSide)> thisTracesTicked = new(StringComparer.OrdinalIgnoreCase);

        // What an image's ticks are kept under - its name AND its file, as a card is told apart.
        private static string TicksKey(string name, string file) => string.Join('\u001F', name.Trim(), file.Trim());

        private static string TicksKey(ComponentPaneBoardPictureRow row) => ComponentPane.TicksKey(row.Name, row.File);

        // ###########################################################################################
        // The previews' width on every card - set by the host for the board shown, and by a drag.
        // Always inside ContributeWidths.PreviewRange - and SHOWN no wider than leaves the settings beside
        // them their room (owner request, 2026-10-30: they could be dragged to 0px): a drag stops there,
        // and a pane made narrower narrows the previews, which widen again with it up to the width chosen.
        // ###########################################################################################
        internal double PreviewWidth
        {
            get => this.thisPreviewWidth;
            set
            {
                double width = double.IsFinite(value)
                    ? Math.Clamp(value, ContributeWidths.PreviewRange.Min, ContributeWidths.PreviewRange.Max)
                    : ContributeWidths.DefaultPreviewWidth;

                // A drag goes no further than the previews can be shown.
                if (this.thisPreviewDrag is not null)
                {
                    width = ContributeWidths.PreviewShown(width, this.thisCardsWidth);
                }

                this.thisPreviewWidth = width;
                this.ShowPreviewWidth();

                // Previews wider than their pictures were read for are read again - never while a
                // drag runs, which would read every card at every step (EndPreviewDrag does then), and
                // only while the cards are what the pane shows.
                if (this.thisPreviewDrag is null && this.thisBoardPart == PaneBoardPart.Pictures)
                {
                    this.LoadPreviews();
                }
            }
        }

        // A drag of the previews' handle ended at this width - the host keeps it for the board.
        internal event EventHandler<double>? PreviewWidthChosen;

        // The width every card shows its previews at: the one chosen, as far as the cards' room allows.
        private void ShowPreviewWidth()
        {
            double shown = ContributeWidths.PreviewShown(this.thisPreviewWidth, this.thisCardsWidth);

            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures)
            {
                row.PreviewWidth = shown;
            }
        }

        // The cards laid out at another width - the previews follow.
        internal void UseCardsWidth(double width)
        {
            this.thisCardsWidth = width;
            this.ShowPreviewWidth();
        }

        internal double ShownPreviewWidthForTests => ContributeWidths.PreviewShown(this.thisPreviewWidth, this.thisCardsWidth);

        // ###########################################################################################
        // Both sides' traces of the image - its name and its file - for its card, or why there are
        // none, in the contributor's words. The host's (TabDrafts.MarkAutomatically.cs): it reads the
        // board's KiCad data as the Schematics tab draws it.
        // ###########################################################################################
        internal Func<string, string, Task<(PreviewTraces Traces, string? Problem)>>? ReadPreviewTraces { get; set; }

        // ###########################################################################################
        // A card made: its previews' width, its sharper image asked for, its traces as ticked - read
        // once for both sides. Ticks the image can no longer show (its alignment or KiCad view undone)
        // are let go of, as CommitCadName lets them go (code review, 2026-10-30: put back, they were
        // greyed out and still ticked, with no way to untick them).
        // ###########################################################################################
        private void UsePreviews(ComponentPaneBoardPictureRow row)
        {
            row.PreviewWidth = ContributeWidths.PreviewShown(this.thisPreviewWidth, this.thisCardsWidth);
            row.WantSharper = wanted => this.TrackPreviewLoad(this.LoadSharperAsync(row, wanted));
            row.TracesTickChanged = this.OnTracesTickChanged;

            string key = ComponentPane.TicksKey(row);

            if (!this.thisTracesTicked.TryGetValue(key, out (bool ThisSide, bool OtherSide) ticked))
            {
                return;
            }

            if (!row.CanShowTraces)
            {
                this.thisTracesTicked.Remove(key);
                return;
            }

            row.SetTicks(ticked.ThisSide, ticked.OtherSide, tell: false);
            this.LoadTraces(row);
        }

        // ---- the previews' width --------------------------------------------------------------------

        private void OnPreviewHandlePressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Control handle || !e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed)
            {
                return;
            }

            this.BeginPreviewDrag(e.GetPosition(this).X);
            e.Pointer.Capture(handle);
            e.Handled = true;
        }

        private void OnPreviewHandleMoved(object? sender, PointerEventArgs e) => this.FollowPreviewDrag(e.GetPosition(this).X);

        // ###########################################################################################
        // From the width SHOWN, under the pointer (code review, 2026-11-04). From the width chosen, wider than the
        // cards have room for, nothing moved until the pointer had covered the difference.
        // ###########################################################################################
        private void BeginPreviewDrag(double fromX) =>
            this.thisPreviewDrag = (fromX, ContributeWidths.PreviewShown(this.thisPreviewWidth, this.thisCardsWidth));

        private void FollowPreviewDrag(double toX)
        {
            if (this.thisPreviewDrag is { } drag)
            {
                this.PreviewWidth = ContributeWidths.PreviewDragged(drag.Width, drag.FromX, toX);
            }
        }

        private void OnPreviewHandleReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (this.thisPreviewDrag is null)
            {
                return;
            }

            e.Pointer.Capture(null);
            this.EndPreviewDrag();
        }

        private void OnPreviewHandleCaptureLost(object? sender, PointerCaptureLostEventArgs e) => this.EndPreviewDrag();

        private void EndPreviewDrag()
        {
            if (this.thisPreviewDrag is null)
            {
                return;
            }

            this.thisPreviewDrag = null;
            this.PreviewWidthChosen?.Invoke(this, this.thisPreviewWidth);
            this.LoadPreviews();
        }

        // A drag of the previews' handle making them `by` wider, start to end, as the pointer's handlers drag.
        internal void DragPreviewWidthForTests(double by)
        {
            this.BeginPreviewDrag(0);
            this.FollowPreviewDrag(-by);
            this.EndPreviewDrag();
        }

        // ---- a sharper image for a zoomed preview ------------------------------------------------------

        // ###########################################################################################
        // The image of `row` decoded `wanted` pixels across for its zoomed preview - ONE decode per
        // card at a time (code review, 2026-10-30: every zoom step started one, each a bitmap of up to
        // 25 MB): an ask while one runs is kept, the widest, and decoded once it ends.
        // ###########################################################################################
        private async Task LoadSharperAsync(ComponentPaneBoardPictureRow row, int wanted)
        {
            if (row.SharpDecoding)
            {
                row.SharpWantedNext = Math.Max(row.SharpWantedNext, wanted);
                return;
            }

            row.SharpDecoding = true;

            try
            {
                while (wanted > 0)
                {
                    row.SharpWantedNext = 0;
                    await this.DecodeSharperAsync(row, wanted);
                    wanted = row.SharpWantedNext;
                }
            }
            finally
            {
                row.SharpDecoding = false;
                row.SharpWantedNext = 0;
            }
        }

        // ###########################################################################################
        // One decode of LoadSharperAsync: in PicturePreviewGeometry.SharpWidth's steps, another card's
        // sharp image let go of first, so one is held. Nothing when the card already shows as many, or
        // went meanwhile. A decode not shown - the card gone, or a sharper one asked for meanwhile,
        // which the next decode makes - was never drawn, so it is disposed at once.
        // ###########################################################################################
        private async Task DecodeSharperAsync(ComponentPaneBoardPictureRow row, int wanted)
        {
            if (this.thisSession is not { } session || !this.thisBoardPictures.Contains(row) || row.File.Length == 0)
            {
                return;
            }

            int width = PicturePreviewGeometry.SharpWidth(wanted, row.ImageSize.Width);

            if (width <= (row.ShownThumbnail?.PixelSize.Width ?? 0))
            {
                return;
            }

            string file = row.File;
            int generation = this.thisImages.Generation;
            string dataRoot = DataManager.DataRoot;
            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);
            string identity = session.ExcelDataFile;

            Bitmap? sharp = await Task.Run(() =>
            {
                string? path = DraftFileResolver.Resolve(dataRoot, draftFolder, file, identity);

                if (path is null || !ContributionPackaging.IsDisplayableImageFile(path))
                {
                    return null;
                }

                try
                {
                    using FileStream stream = File.OpenRead(path);
                    return Bitmap.DecodeToWidth(stream, width, BitmapInterpolationMode.HighQuality);
                }
                catch (Exception ex)
                {
                    // A preview stays as it was - never a pane that fails.
                    Logger.Warning($"Could not read the schematic image [{path}] sharper for its preview - [{ex.Message}]");
                    return null;
                }
            });

            if (sharp is null)
            {
                return;
            }

            if (generation != this.thisImages.Generation || !this.thisBoardPictures.Contains(row) ||
                !string.Equals(row.File, file, StringComparison.Ordinal) ||
                sharp.PixelSize.Width <= (row.ShownThumbnail?.PixelSize.Width ?? 0) ||
                (row.SharpWantedNext > 0 && PicturePreviewGeometry.SharpWidth(row.SharpWantedNext, row.ImageSize.Width) > sharp.PixelSize.Width))
            {
                sharp.Dispose();
                return;
            }

            if (this.thisSharpRow is { } previous && !ReferenceEquals(previous, row))
            {
                previous.SharpThumbnail = null;
            }

            row.SharpThumbnail = sharp;
            this.thisSharpRow = row;
        }

        // ---- the traces of the image's two sides ------------------------------------------------------------

        // ###########################################################################################
        // A tick changed: kept for the image's name and file; the traces read when they are not yet
        // (one read for both sides), else what the line under the ticks says follows the ticks.
        // ###########################################################################################
        private void OnTracesTickChanged(ComponentPaneBoardPictureRow row)
        {
            if (!row.ShowsAnyTraces)
            {
                this.thisTracesTicked.Remove(ComponentPane.TicksKey(row));
                row.TracesNote = string.Empty;
                return;
            }

            this.thisTracesTicked[ComponentPane.TicksKey(row)] = (row.ShowThisSideTraces, row.ShowOtherSideTraces);

            if (row.Traces is null && !row.TracesReading)
            {
                this.LoadTraces(row);
            }
            else if (!row.TracesReading)
            {
                row.TracesNote = ComponentPane.TracesNoteOf(row);
            }
        }

        // The line under the ticks: why there are none, else the sides ticked with no copper.
        private static string TracesNoteOf(ComponentPaneBoardPictureRow row) =>
            row.TracesProblem ?? ComponentPaneWording.TracesNote(
                row.ShowThisSideTraces, row.Traces?.ThisSide.Count ?? 0,
                row.ShowOtherSideTraces, row.Traces?.OtherSide.Count ?? 0);

        // ###########################################################################################
        // The traces of `row`'s image read - again even while a read runs: a read started before
        // another KiCad view was picked, or before a rename, is of what the image no longer is (code
        // review, 2026-10-30: it was kept, and the newer read dropped).
        // ###########################################################################################
        private void LoadTraces(ComponentPaneBoardPictureRow row) => this.TrackPreviewLoad(this.LoadTracesAsync(row));

        // ###########################################################################################
        // The traces of `row`'s image read again - when one is ticked, its card made again, its KiCad
        // view picked or its name changed. A read answers only while it is the LATEST asked for the
        // card (TracesAsk), and only a card still ticked; the latest one ending says what came of it.
        // ###########################################################################################
        private async Task LoadTracesAsync(ComponentPaneBoardPictureRow row)
        {
            if (this.ReadPreviewTraces is not { } read)
            {
                row.TracesNote = PictureToolWording.KiCadNotRead;
                return;
            }

            int ask = ++row.TracesAsk;
            row.TracesReading = true;
            row.TracesNote = ComponentPaneWording.ReadingTraces;

            try
            {
                (PreviewTraces traces, string? problem) = await read(row.Name, row.File);

                if (ask != row.TracesAsk || !this.thisBoardPictures.Contains(row) || !row.ShowsAnyTraces)
                {
                    return;
                }

                row.TracesProblem = problem;
                row.Traces = traces;
                row.TracesNote = ComponentPane.TracesNoteOf(row);
            }
            catch (Exception ex)
            {
                // A read that failed is said as one that found nothing to read - never "Reading..."
                // for good (code review, 2026-10-30).
                Logger.Warning($"Could not read the KiCad traces of [{row.Name}] for its preview - [{ex.Message}]");

                if (ask == row.TracesAsk && row.ShowsAnyTraces)
                {
                    row.TracesNote = PictureToolWording.KiCadNotRead;
                }
            }
            finally
            {
                if (ask == row.TracesAsk)
                {
                    row.TracesReading = false;
                }
            }
        }

        // The ticks let go of with the part - another board's images share no names with these.
        private void ForgetPreviews()
        {
            this.thisTracesTicked.Clear();
            this.thisSharpRow = null;
            this.thisPreviewDrag = null;
        }
    }
}
