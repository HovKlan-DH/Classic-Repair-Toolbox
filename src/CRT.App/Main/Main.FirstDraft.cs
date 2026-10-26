using Handlers.DataHandling;
using Handlers.Online;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // A BOARD'S FIRST DRAFT, MADE WHEN THE FIRST CHANGE IS STARTED (CRT 4.0 step 2, 2026-10-25). See
    // Main.axaml.cs for the file map of the whole class.
    //
    // The Contribute tab's left column - its board parts, "Table view" - works on the board's draft,
    // and a board with none gets one first: the published file as it is on disk
    // (DraftSeeder.SeedFromPublishedFile, the one seeding path every save shares), then the board read
    // again so every surface shows the draft. This replaces "Edit board as draft" (owner request,
    // 2026-10-09), whose own button made a draft on purpose and said so in a popup when one existed:
    // in the workspace a board simply has its draft, and nobody has to start one.
    //
    // *** NOTHING IS MADE UNDER "CRT HAS TO BE UPDATED" (code review, 2026-10-09). *** The tab's work
    // is out of reach under its cover, so the first draft of a run asks the server first
    // (AskApiRevisionBeforeFirstDraftAsync, code review 2026-10-10) and nothing is made when it
    // answers that this CRT is too old - the tab says why.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // True when `entry` has a draft - it had one, or it has one now. False, with the reason on the
        // Contribute tab's own line, when none could be made.
        // ###########################################################################################
        internal async Task<bool> EnsureDraftAsync(HardwareBoardEntry? entry)
        {
            if (entry is null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return false;
            }

            string draftsRoot = DraftManager.DraftsRoot;
            string dataRoot = DataManager.DataRoot;
            string excelDataFile = entry.ExcelDataFile;

            this.TabDrafts.ShowDraftProblem(null);

            if (DraftBoardSource.HasDraft(draftsRoot, excelDataFile))
            {
                return true;
            }

            // The first draft of a run: the server is asked first, so the guard below can fire.
            await this.AskApiRevisionBeforeFirstDraftAsync();

            if (this.TabDrafts.IsUpdateRequiredShown)
            {
                this.TabDrafts.ShowDraftProblem(AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered);
                return false;
            }

            DraftSeedResult seeded = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.MakingDraft,
                () => Task.Run(() => DraftSeeder.SeedFromPublishedFile(draftsRoot, dataRoot, excelDataFile)));

            if (!seeded.Created)
            {
                // Made meanwhile - by a save in another window - is a draft all the same.
                if (DraftBoardSource.HasDraft(draftsRoot, excelDataFile))
                {
                    return true;
                }

                Logger.Warning($"First draft: none made for [{excelDataFile}] - [{seeded.Reason}]");
                this.TabDrafts.ShowDraftProblem($"No draft could be made of this board: {seeded.Reason}");
                return false;
            }

            Logger.Info($"First draft: made a draft of [{excelDataFile}]");

            // The same funnel every save that seeds goes through: the board is read again (it now
            // reads the draft), and the Contribute tab lists the draft.
            this.ReloadCurrentBoardFromDisk(this.TabSchematicsControl.GetCurrentSchematicName());
            this.ApplyContributeTabState();

            return true;
        }
    }
}
