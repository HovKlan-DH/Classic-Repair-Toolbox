using Handlers.DataHandling;
using System;
using System.IO;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "Add a new system" (NewContributeStrategy.md Phase 2, session 2c, task 9) - opens
    // NewSystemWindow, creates the local draft the answer describes, and navigates to it. See
    // Main.axaml.cs for the file map of the whole partial class.
    //
    // This replaces a manual procedure documented in Assets/Wiki/Add-new-board-with-KiCad-data.md:
    // copying an existing board folder by hand, emptying a copied .xlsx while keeping every sheet
    // name exact, and creating a version-named "_UserContribution" workbook to register the board
    // in. Every one of those steps was a place to get it silently wrong.
    //
    // NOTHING HERE WRITES AN .xlsx, and that is the design rather than an omission. The board
    // "workbook" for a drafted system is its draft.json, whose sections ARE BoardData's schema
    // (see DraftDataStore's DraftJsonRoot). Writing an empty Excel file under Drafts/ would be the
    // same second-mechanism mistake the label editor's redirect already rejected.
    //
    // The navigation at the end deliberately drives the ORDINARY board-selection path - the same
    // combo box assignments any other board switch makes - so "from that point everything is
    // ordinary CRT" is true structurally rather than by discipline. There is no second flow for a
    // drafted system to drift away from.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Asks for the new system's identity, creates its draft, and selects it. Does nothing at all
        // when the dialog is cancelled.
        // ###########################################################################################
        internal async Task OpenNewSystemWindowAsync()
        {
            var window = new NewSystemWindow();
            window.Initialize(DataManager.HardwareBoards);

            var registration = await window.ShowDialog<NewSystemRegistration?>(this);
            if (registration == null)
            {
                return;
            }

            if (!TryCreateNewSystem(registration))
            {
                return;
            }

            // The new system exists on disk but nothing in memory knows about it yet. This is the
            // narrow re-scan rather than a full main-workbook reload - see its own comment.
            DataManager.RefreshDraftOnlySystems();

            this.RefreshHardwareAndBoardSelectionsAfterMainExcelSync();
            this.ApplyDraftsTabVisibility();
            this.SelectBoardForNewSystem(registration);

            // ###########################################################################################
            // *** AND LAND ON THE DRAFTS TAB, because that is where the next step is (owner
            // request, 2026-09-24). ***
            //
            // Creating a system is never the end of a task: the contributor now has to add schematic
            // images, KiCad data and component rows, and every one of those actions lives on the
            // Drafts tab. Selecting the board without switching left them on whichever tab they
            // happened to start from - usually Contribute, which has nothing further to offer - with
            // no indication that a new tab had just appeared to hold their work.
            //
            // AFTER SelectBoardForNewSystem, not before: that call drives the ordinary board-selection
            // path, and the Drafts tab reads the selected board when it is shown.
            // ###########################################################################################
            this.SwitchToDraftsTab();
        }

        // ###########################################################################################
        // Moves to the Drafts tab, if it is showing. The twin of SwitchToSchematicsTab, and
        // null-guarded for the same reason: it is callable from another tab, so it cannot assume the
        // tab control is up.
        //
        // Silently does nothing when the tab is hidden. A caller has just created a draft, so
        // ApplyDraftsTabVisibility will have shown it - but a hidden tab cannot be selected, and
        // throwing here would turn a cosmetic miss into a failed system creation.
        // ###########################################################################################
        internal void SwitchToDraftsTab()
        {
            if (this.MainTabControl == null || this.DraftsTabItem == null || !this.DraftsTabItem.IsVisible)
            {
                return;
            }

            if (!ReferenceEquals(this.MainTabControl.SelectedItem, this.DraftsTabItem))
            {
                this.MainTabControl.SelectedItem = this.DraftsTabItem;
            }
        }

        // ###########################################################################################
        // Rebuilds the hardware/board drop-downs after a draft-only system appeared or disappeared,
        // preserving the current selection where it still exists. A named wrapper rather than
        // widening RefreshHardwareAndBoardSelectionsAfterMainExcelSync's own visibility: the two
        // callers have nothing to do with a main-Excel sync, and the method name is the only place
        // that reason is recorded.
        // ###########################################################################################
        internal void RefreshHardwareAndBoardSelectionsAfterDraftChange() =>
            this.RefreshHardwareAndBoardSelectionsAfterMainExcelSync();

        // ###########################################################################################
        // Reloads the board on screen after board images or KiCad data were imported into a draft,
        // but ONLY when the imported system is the one currently selected - importing into some
        // other system must not yank the user's current board out from under them.
        //
        // A reload is genuinely needed rather than a refresh: a newly imported image adds a
        // Schematics row (so the thumbnail list changes) and newly imported KiCad files are only
        // discovered when the board's KiCad paths are resolved again.
        // ###########################################################################################
        internal void ReloadCurrentBoardAfterDraftFileImport(string excelDataFile)
        {
            var current = this.GetCurrentBoardEntry();

            if (current == null ||
                !string.Equals(current.ExcelDataFile, excelDataFile, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            this.ReloadCurrentBoardFromDisk(string.Empty);
        }

        // ###########################################################################################
        // Creates the system's draft folder, with a real board workbook inside it. Returns false and
        // logs when anything fails - a half-created system is worse than none, so the caller stops
        // rather than navigating to something that may not load.
        //
        // *** IT IS A BOARD FOLDER NOW, NOT A Files/ FOLDER PLUS draft.json (Phase 6, 2026-09-23).
        // *** This used to create the folder, an empty Files/ subfolder for images to land in, and
        // a draft.json holding a registration and eleven empty sections. The contributor opening it
        // by hand found a shape that looked nothing like the board it was a draft of - and could
        // not open the board at all, because there was no workbook.
        //
        // DraftSeeder.CreateNewSystem writes the workbook (headers and sheet names, every section
        // empty) plus the marker carrying the registration. That workbook is hand-editable from the
        // moment it exists, which is the whole point of the change.
        //
        // NOTE Phase 2 explicitly rejected writing an .xlsx under Drafts/. That reasoning was about
        // an EMPTY placeholder sitting alongside draft.json - two mechanisms for one job. This
        // replaces draft.json rather than joining it; see DraftSeeder's own header.
        // ###########################################################################################
        private static bool TryCreateNewSystem(NewSystemRegistration registration)
        {
            DraftSeedResult result = DraftSeeder.CreateNewSystem(DraftManager.DraftsRoot, registration);

            if (!result.Created)
            {
                Logger.Warning(
                    $"Could not create the new system [{registration.ExcelDataFile}] - [{result.Reason}]");

                return false;
            }

            Logger.Info(
                $"Created new system [{registration.HardwareName}] / [{registration.BoardName}] " +
                $"at [{result.SystemFolder}]");

            return true;
        }

        // ###########################################################################################
        // Selects the newly created system in the hardware/board drop-downs, which runs the ordinary
        // board load for it.
        //
        // Hardware is switched first so OnHardwareSelectionChanged has repopulated BoardComboBox's
        // items before the board is set, and _pendingBoardSelectionOverride makes that switch land
        // on the new board directly - exactly the sequence ActivateWorkbook already uses, and for
        // the same reasons its own comment gives.
        // ###########################################################################################
        private void SelectBoardForNewSystem(NewSystemRegistration registration)
        {
            var currentHardware = this.HardwareComboBox.SelectedItem as string;

            if (!string.Equals(currentHardware, registration.HardwareName, StringComparison.OrdinalIgnoreCase))
            {
                this._pendingBoardSelectionOverride = registration.BoardName;
                this.HardwareComboBox.SelectedItem = registration.HardwareName;
            }

            this.BoardComboBox.SelectedItem = registration.BoardName;
        }
    }
}
