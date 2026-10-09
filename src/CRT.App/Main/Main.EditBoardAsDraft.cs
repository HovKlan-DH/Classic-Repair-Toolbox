using Handlers.DataHandling;
using Handlers.Online;
using System;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "Edit board as draft", the Contribute tab's button between "Add new component" and "Add a new
    // board" (owner request, 2026-10-09: "should create a new draft of the selected board, if it
    // does not already exists. If it exits already, it should be informed via popup, that there
    // cannot be two draft for same board"). See Main.axaml.cs for the file map of the whole class.
    //
    // A draft is the contributor's own copy of a board, edited in the Drafts tab's table and sent
    // from there. Until now one appeared only as a side effect of saving a component or the labels;
    // this makes one on purpose, for the whole board.
    //
    //   - No draft yet: seeded from the published file as it is on disk (DraftSeeder.
    //     SeedFromPublishedFile, the one seeding path every save shares), the board re-read so every
    //     surface shows the draft, and the Drafts tab opened on the draft's table.
    //   - A draft already: DraftExistsWindow says so and offers it - "Open the draft" lands on the
    //     same table. Nothing is written.
    //   - The Drafts tab says "CRT has to be updated" (code review, 2026-10-09): nothing is made or
    //     opened - its table is out of reach under the cover - and the Contribute tab says why.
    // ###########################################################################################
    public partial class Main
    {
        // For tests: the board shown, in place of the drop-downs' (which need a main Excel data file),
        // and DraftExistsWindow's answer for the board it names - ShowDialog blocks headlessly.
        internal HardwareBoardEntry? EditAsDraftBoardOverrideForTests { get; set; }

        internal Func<string, bool>? ExistingDraftAnswerForTests { get; set; }

        internal async Task EditBoardAsDraftAsync()
        {
            HardwareBoardEntry? entry = this.EditAsDraftBoardOverrideForTests ?? this.GetCurrentBoardEntry();

            if (entry is null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return;
            }

            string draftsRoot = DraftManager.DraftsRoot;
            string dataRoot = DataManager.DataRoot;
            string excelDataFile = entry.ExcelDataFile;

            this.TabContribute.ShowDraftProblem(null);

            if (this.TabDrafts.IsUpdateRequiredShown)
            {
                this.TabContribute.ShowDraftProblem(AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered);
                return;
            }

            if (DraftBoardSource.HasDraft(draftsRoot, excelDataFile))
            {
                await this.OfferExistingDraftAsync(entry);
                return;
            }

            DraftSeedResult seeded = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.MakingDraft,
                () => Task.Run(() => DraftSeeder.SeedFromPublishedFile(draftsRoot, dataRoot, excelDataFile)));

            if (!seeded.Created)
            {
                // Made meanwhile - by a save in another window - is the case above after all.
                if (DraftBoardSource.HasDraft(draftsRoot, excelDataFile))
                {
                    await this.OfferExistingDraftAsync(entry);
                    return;
                }

                Logger.Warning($"Edit board as draft: no draft made for [{excelDataFile}] - [{seeded.Reason}]");
                this.TabContribute.ShowDraftProblem($"No draft could be made of this board: {seeded.Reason}");
                return;
            }

            Logger.Info($"Edit board as draft: made a draft of [{excelDataFile}]");

            // The same funnel every save that seeds goes through: the board is read again (it now
            // reads the draft) and the Drafts tab appears for the first draft.
            this.ReloadCurrentBoardFromDisk(this.TabSchematicsControl.GetCurrentSchematicName());
            await this.OpenDraftTableAsync(entry);
        }

        private async Task OfferExistingDraftAsync(HardwareBoardEntry entry)
        {
            string name = $"{entry.HardwareName} / {entry.BoardName}";
            bool open;

            if (this.ExistingDraftAnswerForTests is not null)
            {
                open = this.ExistingDraftAnswerForTests(name);
            }
            else
            {
                var window = new DraftExistsWindow();
                window.Initialize(name);
                open = await window.ShowDialog<bool?>(this) == true;
            }

            if (open)
            {
                await this.OpenDraftTableAsync(entry);
            }
        }

        // The Drafts tab, with this board's draft open in its table - where the editing happens.
        private async Task OpenDraftTableAsync(HardwareBoardEntry entry)
        {
            this.ApplyDraftsTabVisibility();
            this.SwitchToDraftsTab();
            await this.TabDrafts.ShowTableAsync(entry);
        }
    }
}
