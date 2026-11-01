using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "MARK ALL AUTOMATICALLY" (CRT 4.0 step 3, slice 3c - the plan: "Mark: on aligned PCB pictures,
    // Mark all automatically draws each component's rectangle from its footprint").
    //
    // A picture card's button in the component pane's Pictures part: the draft opened as every change
    // opens it (DraftEdits.OpenForChangeAsync - what the editors hold handed over, a file written
    // outside read), the board's KiCad data read (KiCadProjectLoader keeps the parse) and the image's
    // size read from its header, then KiCadFootprintMarks.Plan decides - a side of the board, aligned,
    // or the reason it cannot. Every component the board has, seen from that side and not marked on
    // the picture yet, gets its footprint's rectangle as ONE step of the draft's history, worked out
    // on the draft as the change finds it, saved at once - all under the BusyOverlay. What came of it
    // is said in the part; the board on screen follows as it does the pane's saves.
    //
    // The same plan places both sides' traces on a card's preview (PreviewTracesAsync,
    // 2026-10-29) - its "Traces on the preview".
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Stands in for reading a picture's size from its image - a test has no images.
        internal Func<string, (int Width, int Height)?>? PictureSizeOverrideForTests { get; set; }

        // Stands in for whether a picture tool is on on the Schematics tab - a test has no Main.
        internal Func<bool>? PictureToolOnOverrideForTests { get; set; }

        // `imageFile` tells the card's own picture from another of the same name (code review,
        // 2026-10-30: the first of the name was planned from - its view and its size).
        internal Task MarkAutomaticallyAsync(string picture, string imageFile = "") =>
            this.OnCurrentDraftAsync(entry => this.MarkFromKiCadAsync(entry, picture, imageFile, afterAligning: false));

        // ###########################################################################################
        // *** AFTER "ALIGN KiCad"'s DONE, THE PICTURE IS MARKED BY ITSELF (owner request, 2026-11-01: "it
        // is not clear that I must first do the Align KiCad, before I can do the Mark all automatically
        // from KiCad" - Done marks without asking, Undo takes it back). *** Main asks once the tool has
        // handed back to this tab, for every picture aligned with Done on the way (Main.PictureTools.cs).
        // As "Mark all automatically", with three differences: a picture with ANY mark on it is left
        // alone - its marks were drawn, by hand or from KiCad, and another round is the contributor's to
        // ask for; a schematic sheet is passed over without a word (nothing to mark from KiCad there);
        // and what came of it says the alignment is in too.
        //
        // *** ALIGNED AGAIN, THE MARKS KiCad DREW FOLLOW (owner request, 2026-11-01: "if I then again do the
        // Align KiCad then it seems it will not readjust the components, which I think it should"). ***
        // `alignedFrom` is the alignment the picture had - null for its first. KiCadFootprintMarks.Realigned
        // tells the marks KiCad drew from those moved or drawn by hand, which stay.
        // ###########################################################################################
        internal Task MarkAfterAligningAsync(string picture, string imageFile = "", KiCadCalibrationEntry? alignedFrom = null) =>
            this.OnCurrentDraftAsync(entry => this.MarkFromKiCadAsync(entry, picture, imageFile, afterAligning: true, alignedFrom));

        private async Task MarkFromKiCadAsync(HardwareBoardEntry entry, string picture, string imageFile, bool afterAligning, KiCadCalibrationEntry? alignedFrom = null)
        {
            // A picture tool's Done writes its picture's rectangles as it holds them - marks put in
            // meanwhile would go without a word (code review, 2026-10-26). As starting a tool shows the
            // one that is on, this waits for it.
            if (this.PictureToolOnOverrideForTests?.Invoke() ?? this.thisMainWindow?.IsPictureToolOn == true)
            {
                this.ComponentPane.ShowPicturesNote(PictureToolWording.PictureToolOn);
                return;
            }

            string identity = entry.ExcelDataFile;
            bool changed = false;

            string? note = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.MarkingFromKiCad, async () =>
            {
                if (await DraftEdits.OpenForChangeAsync(identity) is not DraftSession session)
                {
                    return DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.NoDraft);
                }

                bool hasMarks = afterAligning && KiCadFootprintMarks.HasMarks(session.Content.Board, picture);

                if (hasMarks && alignedFrom is null)
                {
                    return null;
                }

                KiCadProjectRoot? root = await this.ReadKiCadForReadinessAsync(identity, "marking its components");
                (int Width, int Height)? size = await this.PictureSizeAsync(session, picture, imageFile);
                BoardContent content = session.Content;
                KiCadMarkPlan plan = await Task.Run(() => KiCadFootprintMarks.Plan(content, picture, root, size, imageFile));

                if (plan.Problem is { } problem)
                {
                    return afterAligning && plan.NotABoardSide ? null : problem;
                }

                if (hasMarks)
                {
                    KiCadCalibrationValues alignedBefore = KiCadFootprintMarks.ValuesOf(alignedFrom!);
                    KiCadFootprintMarks.Realigned(session.Content.Board, plan, alignedBefore, out int toFollow);

                    if (toFollow == 0)
                    {
                        return null;
                    }

                    BoardContent was = session.Content;
                    int followed = 0;

                    DraftWorkbookEditOutcome moved = await DraftEdits.ApplyAndSaveAsync(
                        identity,
                        DraftStepWording.ComponentsFollowedAlignment(plan.Picture, toFollow),
                        board =>
                        {
                            BoardData result = KiCadFootprintMarks.Realigned(board, plan, alignedBefore, out int counted);
                            followed = counted;
                            return result;
                        });

                    if (moved != DraftWorkbookEditOutcome.Saved)
                    {
                        return DraftSaveWording.EditNotSaved(moved);
                    }

                    changed = !ReferenceEquals(session.Content, was);

                    return changed ? PictureToolWording.MarksFollowedAlignment(followed, plan.Picture) : null;
                }

                int count = plan.Marks(session.Content.Board).Count;

                if (count == 0)
                {
                    return plan.WhyNothing(session.Content.Board);
                }

                BoardContent before = session.Content;
                int marked = 0;

                DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(
                    identity,
                    DraftStepWording.ComponentsMarkedFromKiCad(plan.Picture, count),
                    board =>
                    {
                        List<ComponentHighlightEntry> marks = plan.Marks(board);
                        marked = marks.Count;
                        return KiCadFootprintMarks.WithMarks(board, marks);
                    });

                if (outcome != DraftWorkbookEditOutcome.Saved)
                {
                    return DraftSaveWording.EditNotSaved(outcome);
                }

                changed = !ReferenceEquals(session.Content, before);

                return !changed ? plan.WhyNothing(session.Content.Board)
                    : afterAligning ? PictureToolWording.MarkedAfterAligning(marked, plan.Picture)
                    : PictureToolWording.MarkedAutomatically(marked, plan.Picture);
            });

            if (note is not null)
            {
                this.ComponentPane.ShowPicturesNote(note);
            }

            if (changed)
            {
                // The draft's row and the board on screen follow, as they do the pane's own saves.
                this.OnComponentPaneSaved(identity);
            }
        }

        // ###########################################################################################
        // The traces of `picture`'s two sides, for its card's "Traces on the preview" (owner requests,
        // 2026-10-29): what "Mark all automatically" works with - the board's KiCad data as the
        // Schematics tab draws it, the image's size from its header, its alignment in the draft
        // (KiCadFootprintMarks.Plan) - and each side's copper placed on the image (KiCadPreviewTraces).
        // Or why there are none, in the contributor's words. `imageFile` tells the card's own picture
        // from another of the same name (code review, 2026-10-30).
        // ###########################################################################################
        internal async Task<(PreviewTraces Traces, string? Problem)> PreviewTracesAsync(string picture, string imageFile = "")
        {
            if (this.thisComponentEntry is not { } entry || this.ComponentPane.Session is not { } session)
            {
                return (PreviewTraces.None, PictureToolWording.KiCadNotRead);
            }

            KiCadProjectRoot? root = await this.ReadKiCadForPreviewsAsync(entry.ExcelDataFile);
            (int Width, int Height)? size = await this.PictureSizeAsync(session, picture, imageFile);
            BoardContent content = session.Content;

            return await Task.Run(() =>
            {
                KiCadMarkPlan plan = KiCadFootprintMarks.Plan(content, picture, root, size, imageFile);

                return plan.Problem is { } problem
                    ? (PreviewTraces.None, (string?)problem)
                    : (KiCadPreviewTraces.For(plan), null);
            });
        }

        // ###########################################################################################
        // The board's KiCad project for the cards' traces - ONE read shared by every card asking while
        // it runs (code review, 2026-10-30: the cards are made again at every undo, and four cards
        // ticked read the project four times at once). A read done is not kept: the next ask reads
        // again, the draft's KiCad files may have changed (KiCadProjectLoader keeps the parse).
        // ###########################################################################################
        private (string Identity, Task<KiCadProjectRoot?> Read)? thisPreviewKiCadRead;

        private Task<KiCadProjectRoot?> ReadKiCadForPreviewsAsync(string identity)
        {
            if (this.thisPreviewKiCadRead is { } running && !running.Read.IsCompleted &&
                string.Equals(running.Identity, identity, StringComparison.OrdinalIgnoreCase))
            {
                return running.Read;
            }

            Task<KiCadProjectRoot?> read = this.ReadKiCadForReadinessAsync(identity, "its schematic images' traces");
            this.thisPreviewKiCadRead = (identity, read);
            return read;
        }

        // ###########################################################################################
        // The size of `picture`'s image in pixels, from its header (ImagePixelSize) - found as the
        // pane's previews find it (DraftFileResolver), off the UI thread. Null when it cannot be read.
        // The picture's row is KiCadFootprintMarks.PictureRow's - the one with `imageFile` among two
        // of a name.
        // ###########################################################################################
        private Task<(int Width, int Height)?> PictureSizeAsync(DraftSession session, string picture, string imageFile = "")
        {
            string file = KiCadFootprintMarks.PictureRow(session.Content.Board, picture, imageFile)?.SchematicImageFile?.Trim() ?? string.Empty;

            if (this.PictureSizeOverrideForTests is { } read)
            {
                return Task.FromResult(read(file));
            }

            string dataRoot = DataManager.DataRoot;
            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);
            string identity = session.ExcelDataFile;

            return Task.Run(() => file.Length == 0
                ? null
                : ImagePixelSize.TryRead(DraftFileResolver.Resolve(dataRoot, draftFolder, file, identity)));
        }
    }
}
