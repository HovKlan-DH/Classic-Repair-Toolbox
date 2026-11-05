using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "KiCad DATA" ON THE RIGHT OF THE TAB (owner request, 2026-10-30: "I do not understand why the
    // 'KiCad data' opens in a new window? I think this should just show up in the 'normal' content
    // area, and not in a new window").
    //
    // The left column's "KiCad data" shows a KiCadDataPanel where a component or a board part shows -
    // ONE thing on the right at a time: opening it closes the table and the pane (what they hold
    // saved), and opening either of them, or "Components", closes it. A panel is made for each opening,
    // as the window was, so nothing of one opening is left in the next.
    //
    //   - EACH CHANGE IS CAUGHT UP WITH AS IT ENDS (KiCadDataPanel.DraftChanged): the draft's row and
    //     the left column's count, and the board on screen read again - new KiCad files are new
    //     traces. The window did that once it closed; a panel can stay open while the contributor goes
    //     to the Schematics tab.
    //   - CLOSING WAITS FOR ITS CHANGES (WaitForChangesAsync), under CRT's "please wait", so nothing it
    //     writes lands with nothing on screen saying so - except when another board is shown or the
    //     draft goes, which cannot be called off: the panel goes at once and its changes end by
    //     themselves, still caught up with.
    //   - AN IMPORT ADDS THE KiCad BOARD'S COMPONENTS TOO (owner request, 2026-11-01: "it should add all
    //     components when importing the KiCad data. Then you can delete this option from 'Board
    //     status'"), while the panel's check box is ticked - the tab's AddFromKiCadToAsync, the one
    //     rule, handed to the panel; its summary offers the rest when it was not.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // The panel open on the right, and the board it is open on - none while it is closed.
        private KiCadDataPanel? thisKiCadPanel;
        private HardwareBoardEntry? thisKiCadEntry;

        internal bool IsKiCadOpen => this.thisKiCadPanel is not null;

        private bool IsKiCadOpenFor(HardwareBoardEntry entry) =>
            this.thisKiCadEntry is { } open && TabDrafts.SameBoard(open, entry);

        internal KiCadDataPanel? KiCadPanelForTests => this.thisKiCadPanel;

        // Whether a KiCad import adds the components too, and remembering it - handed in by Initialize
        // (UserSettings), so a test's tab never touches the settings; none is "yes".
        internal Func<bool>? KiCadImportAddsComponents { get; set; }

        internal Action<bool>? RememberKiCadImportAddsComponents { get; set; }

        // ###########################################################################################
        // Opens "KiCad data" for one drafted board (session 2c, task 9) - its pictures are the Pictures
        // part's since step 3.
        //
        // The board labels and the "Important signals" rows handed over - what the KiCad report judges
        // the KiCad data against - are the DRAFT's, read from its board file on disk
        // (BoardForFilesWindowAsync), after the table and the pane have saved what they hold.
        // ###########################################################################################
        private async Task ManageFilesAsync(HardwareBoardEntry entry)
        {
            // What closes on the way hands the draft over - never lets it go (OpeningOn).
            using (this.OpeningOn(entry))
            {
                await this.OpenKiCadOnDraftAsync(entry);
            }
        }

        private async Task OpenKiCadOnDraftAsync(HardwareBoardEntry entry)
        {
            if (this.IsKiCadOpenFor(entry))
            {
                return;
            }

            // A picture's name the Pictures part cannot take keeps it open (OpenPaneAsync has why).
            if (this.IsComponentOpen && this.ComponentPane.KeepsRefusedPictureName())
            {
                return;
            }

            // ###########################################################################################
            // *** WHAT THE DRAFT HOLDS IS SAVED FIRST, OR ASKED ABOUT ONCE (code review, 2026-11-01). ***
            // Every change the panel makes is saved with what the draft holds, so one that cannot be
            // saved would refuse each of them too, in words of its own - and an import's files would stay
            // unlisted without a word. Asked while everything is still open: Cancel leaves it all as it
            // was. Moving this from a window to a panel (2026-10-30) had dropped it.
            // ###########################################################################################
            if (!await this.SettleTableEditsAsync(entry))
            {
                return;
            }

            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                return;
            }

            if (this.IsComponentOpen)
            {
                this.CloseComponent();

                Task paneSave = this.ComponentPane.LastSave;

                if (!paneSave.IsCompleted)
                {
                    await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SavingDraft, () => paneSave);
                }
            }

            await this.CloseKiCadAsync();

            // Under CRT's "please wait" (owner request, 2026-10-10): reading the board took a few
            // seconds with nothing on screen. The panel says what it reads itself (KiCadDataPanel.Opening).
            BoardData? boardData = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.OpeningBoardFiles,
                () => TabDrafts.BoardForFilesWindowAsync(entry));

            // Another board shown meanwhile, or something else opened: nothing opens.
            if (this.thisCurrentEntry is not { } shown || !TabDrafts.SameBoard(shown, entry) ||
                this.IsTableOpen || this.IsComponentOpen || this.IsKiCadOpen)
            {
                return;
            }

            var boardLabels = boardData?.Components
                .Select(component => component.BoardLabel)
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .ToList() ?? new List<string>();

            var panel = new KiCadDataPanel
            {
                // An import adds the KiCad board's components too, while its check box is ticked - the
                // choice remembered (owner request, 2026-11-01); the panel catches up with the change.
                AddKiCadComponents = () => this.AddFromKiCadToAsync(entry),
                AddsComponentsOnImport = this.KiCadImportAddsComponents?.Invoke() ?? true,
                AddsComponentsOnImportChanged = value => this.RememberKiCadImportAddsComponents?.Invoke(value),
            };

            panel.DraftChanged += (_, _) => this.CatchUpAfterBoardFiles(entry, changedTheDraft: true);

            this.thisKiCadPanel = panel;
            this.thisKiCadEntry = entry;
            this.thisNoComponentPage = null;
            this.KiCadPanelHost.Content = panel;
            this.ShowWorkspace();

            // In the tab before it reads, so its "please wait" is CRT's own (BusyOverlay.For).
            panel.Initialize(entry.ExcelDataFile, boardLabels, boardData?.KiCadImportantSignals);
        }

        // The left column's "KiCad data", and Board status's KiCad panel.
        internal Task OpenKiCadAsync() => this.OnCurrentDraftAsync(entry => this.ManageFilesAsync(entry));

        // ###########################################################################################
        // Closes the panel once its changes have ended - "Close", or something else opened in its place.
        // ###########################################################################################
        internal async Task CloseKiCadAsync()
        {
            if (this.thisKiCadPanel is not { } panel)
            {
                return;
            }

            // Each change holds CRT's "please wait" itself while it runs.
            await panel.WaitForChangesAsync();

            // Closed meanwhile - another board shown.
            if (ReferenceEquals(this.thisKiCadPanel, panel))
            {
                this.CloseKiCadNow();
            }
        }

        // ###########################################################################################
        // The panel taken away at once - another board shown, or its draft going, neither of which can
        // be called off. Its changes still being written end by themselves and are still caught up with;
        // a draft made only to open it goes again, when nothing is still being written into it.
        // ###########################################################################################
        private void CloseKiCadNow()
        {
            if (this.thisKiCadPanel is not { } panel)
            {
                return;
            }

            HardwareBoardEntry? entry = this.thisKiCadEntry;

            panel.Close();
            this.thisKiCadPanel = null;
            this.thisKiCadEntry = null;
            this.KiCadPanelHost.Content = null;
            this.ShowWorkspace();

            if (entry is not null && panel.DraftChanging.IsCompleted)
            {
                this.LetGoOfDraftMadeOnlyToLook(entry);
            }
        }

        // ###########################################################################################
        // After a change "KiCad data" made to the draft - and ONLY then (owner report, 2026-10-10: the
        // board on screen was reloaded every time, changed or not). The row's count follows, and the
        // board on screen is read again if it is this one - new KiCad files are new traces (the
        // pictures' own changes are the Pictures part's, saved as it makes them). An open table needs
        // nothing: it showed each change as the panel made it (the draft held in memory, CRT 4.0 step 1).
        // ###########################################################################################
        internal void CatchUpAfterBoardFiles(HardwareBoardEntry entry, bool changedTheDraft)
        {
            if (!changedTheDraft)
            {
                return;
            }

            // Its files are counted again - one copied in may not be listed (ShowKiCadCount).
            this.ForgetKiCadCount();
            this.RefreshDrafts();

            if (this.ReloadBoardOverrideForTests is { } reload)
            {
                reload(entry.ExcelDataFile);
            }
            else
            {
                this.thisMainWindow?.ReloadCurrentBoardAfterDraftFileImport(entry.ExcelDataFile);
            }
        }

        // Stands in for Main's reload of the board on screen in a test, which has no Main.
        internal Action<string>? ReloadBoardOverrideForTests { get; set; }
    }
}
