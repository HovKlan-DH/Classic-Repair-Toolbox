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
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Stands in for reading a picture's size from its image - a test has no images.
        internal Func<string, (int Width, int Height)?>? PictureSizeOverrideForTests { get; set; }

        // Stands in for whether a picture tool is on on the Schematics tab - a test has no Main.
        internal Func<bool>? PictureToolOnOverrideForTests { get; set; }

        internal Task MarkAutomaticallyAsync(string picture) =>
            this.OnCurrentDraftAsync(async entry =>
            {
                // A picture tool's Done writes its picture's rectangles as it holds them - marks put in
                // meanwhile would go without a word (code review, 2026-10-26). As starting a tool shows
                // the one that is on, this waits for it.
                if (this.PictureToolOnOverrideForTests?.Invoke() ?? this.thisMainWindow?.IsPictureToolOn == true)
                {
                    this.ComponentPane.ShowPicturesNote(PictureToolWording.PictureToolOn);
                    return;
                }

                string identity = entry.ExcelDataFile;
                bool changed = false;

                string note = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.MarkingFromKiCad, async () =>
                {
                    if (await DraftEdits.OpenForChangeAsync(identity) is not DraftSession session)
                    {
                        return DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.NoDraft);
                    }

                    KiCadProjectRoot? root = await this.ReadKiCadForReadinessAsync(identity, "marking its components");
                    (int Width, int Height)? size = await this.PictureSizeAsync(session, picture);
                    BoardContent content = session.Content;
                    KiCadMarkPlan plan = await Task.Run(() => KiCadFootprintMarks.Plan(content, picture, root, size));

                    if (plan.Problem is { } problem)
                    {
                        return problem;
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

                    return changed
                        ? PictureToolWording.MarkedAutomatically(marked, plan.Picture)
                        : plan.WhyNothing(session.Content.Board);
                });

                this.ComponentPane.ShowPicturesNote(note);

                if (changed)
                {
                    // The draft's row and the board on screen follow, as they do the pane's own saves.
                    this.OnComponentPaneSaved(identity);
                }
            });

        // ###########################################################################################
        // The size of `picture`'s image in pixels, from its header (ImagePixelSize) - found as the
        // pane's previews find it (DraftFileResolver), off the UI thread. Null when it cannot be read.
        // ###########################################################################################
        private Task<(int Width, int Height)?> PictureSizeAsync(DraftSession session, string picture)
        {
            string file = session.Content.Board.Schematics
                .FirstOrDefault(row => string.Equals(row.SchematicName?.Trim(), picture?.Trim(), StringComparison.OrdinalIgnoreCase))
                ?.SchematicImageFile?.Trim() ?? string.Empty;

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
