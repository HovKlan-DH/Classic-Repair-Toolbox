using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT PANE on the right of the Contribute tab (CRT 4.0 step 2, 2026-10-25 - the plan:
    // "A component opens in a pane, never in a maximized window").
    //
    // Owns: which component of which draft is open (thisComponentEntry), opening and closing the pane,
    // and what follows its changes - the left column reads the draft as the pane left it, the draft's
    // row follows each save, and the board on screen catches up once the pane closes or the tab is
    // left (CatchUpBoardOnScreen, as for the table: reading the whole board after every typed field
    // would make typing lag).
    //
    // The right side shows ONE of the table, the pane and the start panel (ApplyTableMode). A
    // component opened closes the table, and the table opened closes the pane - each having saved
    // what it holds.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // The draft the pane is open on - null while it is closed.
        private HardwareBoardEntry? thisComponentEntry;

        // ###########################################################################################
        // What the pane changed, per label - "Changed" in the left column follows the pane at once,
        // while the board on screen (where the marks come from) catches up later. Forgotten whenever
        // Main shows the board again: its marks are then up to date.
        // ###########################################################################################
        private readonly Dictionary<string, bool> thisChangedHere = new(StringComparer.OrdinalIgnoreCase);

        internal bool IsComponentOpen => this.thisComponentEntry is not null;

        private bool IsComponentOpenFor(HardwareBoardEntry entry) =>
            this.thisComponentEntry is not null &&
            string.Equals(this.thisComponentEntry.ExcelDataFile, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase);

        internal ComponentPane ComponentPaneForTests => this.ComponentPane;

        private void WireComponentPane()
        {
            this.ComponentPane.CloseRequested += (_, _) => this.CloseComponent();
            this.ComponentPane.ComponentEditorRequested += (_, _) => this.OpenComponentEditorWindow();
            this.ComponentPane.ContentChanged += (_, _) => this.OnComponentPaneChanged();
            this.ComponentPane.Saved += (_, excelDataFile) => this.OnComponentPaneSaved(excelDataFile);
        }

        // ###########################################################################################
        // Opens `label` of the board shown in the pane - its draft made first when it has none
        // (OnCurrentDraftAsync), the table closed first when it is open (what it holds saved).
        // ###########################################################################################
        internal Task OpenComponentAsync(string label) =>
            this.OnCurrentDraftAsync(entry => this.ShowComponentAsync(entry, label));

        private async Task ShowComponentAsync(HardwareBoardEntry entry, string label)
        {
            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                return;
            }

            DraftStatus? status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile, entry.IsPublished);

            BoardData? published = this.PublishedBoardOverrideForTests is not null
                ? this.PublishedBoardOverrideForTests(entry)
                : await BusyOverlay.RunLocalAsync(this, CrtWaitWording.OpeningComponent, () => TabDrafts.LoadPublishedBoardAsync(entry, status));

            // The one session of this draft, as the table opens it - something may have written the
            // file meanwhile: read again first, when nothing of the session's waits to be saved.
            DraftSession? session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, entry.ExcelDataFile);

            if (session is null)
            {
                Logger.Warning($"Could not open [{label}] of the draft of [{entry.ExcelDataFile}] - its board file could not be read");
                return;
            }

            session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep);

            // Its list of files too, so a picture replaced under its own name reads as a change.
            string publishedBoardFile = published is null ? string.Empty : DraftBoardSource.PublishedPathOf(DataManager.DataRoot, entry.ExcelDataFile);

            this.ComponentPane.Show(session, label, this.thisCurrentRegion, published, publishedBoardFile);
            this.thisComponentEntry = entry;

            // The draft's header follows the session now open.
            this.RefreshDrafts();
        }

        // ###########################################################################################
        // Closes the pane - what it holds goes into the draft, saved as it closes. The draft's row and
        // the board on screen follow that save (OnComponentPaneSaved): read now, from the file the
        // save is replacing, they kept it from being replaced (found by a test that failed in some
        // orders only - "WriteFailed").
        // ###########################################################################################
        internal void CloseComponent()
        {
            if (this.thisComponentEntry is null)
            {
                return;
            }

            this.ComponentPane.Close();
            this.thisComponentEntry = null;
            this.ApplyTableMode();
            this.ShowWorkspace();

            if (this.ComponentPane.LastSave.IsCompleted)
            {
                this.CatchUpBoardOnScreen();
            }
        }

        // The pane let go of a draft that is going - nothing handed over, nothing saved.
        private void ForgetComponent()
        {
            this.ComponentPane.Forget();
            this.thisComponentEntry = null;
        }

        // ###########################################################################################
        // Everything the pane holds saved now - under CRT's "please wait", since the contributor
        // pressed for it. True when nothing is left unsaved (or the pane is closed).
        // ###########################################################################################
        internal async Task<bool> SaveComponentNowAsync()
        {
            if (!this.IsComponentOpen)
            {
                return true;
            }

            this.ComponentPane.CommitHeld();

            if (!this.ComponentPane.HasUnsavedChanges)
            {
                return true;
            }

            return await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SavingDraft, this.ComponentPane.SaveNowAsync);
        }

        // The pane's change: the left column reads the draft as it is now.
        private void OnComponentPaneChanged()
        {
            if (this.ComponentPane.Label.Length > 0)
            {
                this.thisChangedHere[this.ComponentPane.Label] = this.ComponentPane.IsComponentChanged;
            }

            this.ShowBoardParts();
            this.ShowComponents();
        }

        // ###########################################################################################
        // The pane's draft saved: its row follows at once; the board on screen once the pane closes or
        // the tab is left - at once when the tab is not on screen (a save the pane started as it went).
        // ###########################################################################################
        private void OnComponentPaneSaved(string excelDataFile)
        {
            this.thisBoardOnScreenBehind = excelDataFile;

            if (this.thisMainWindow is not null)
            {
                this.thisMainWindow.ApplyContributeTabState();
            }
            else
            {
                this.RefreshDrafts();
            }

            if (this.thisActivationWindow is null || !this.IsComponentOpen)
            {
                this.CatchUpBoardOnScreen();
            }
        }

        // ###########################################################################################
        // The component editor window on the pane's component - its pictures, baselines and files,
        // until they move into the pane. It is handed the draft AS THE PANE LEFT IT: Main's board may
        // not have caught up with the pane yet, and a window opened on that would save it back.
        // ###########################################################################################
        private void OpenComponentEditorWindow()
        {
            if (this.ComponentPane.Session is not { } session || this.ComponentPane.Label.Length == 0)
            {
                return;
            }

            if (this.ComponentEditorOverrideForTests is { } open)
            {
                open(this.ComponentPane.Label);
                return;
            }

            this.thisMainWindow?.OpenComponentContributionWindow(this.ComponentPane.Label, session.Content.Board);
        }

        // For tests, in place of the component editor window.
        internal Action<string>? ComponentEditorOverrideForTests { get; set; }

        // The board the left column shows: the draft as the pane or the table left it, when one is open
        // on the board shown - else the board Main showed.
        private BoardData? BoardForList()
        {
            if (this.thisCurrentEntry is { } entry &&
                (this.IsComponentOpenFor(entry) || this.IsTableOpenFor(entry)) &&
                DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is { } session)
            {
                return session.Content.Board;
            }

            return this.thisCurrentBoard;
        }

        // Main's marks, with what the pane changed since the board was last shown.
        private IReadOnlySet<string>? DraftedLabelsForList()
        {
            if (this.thisChangedHere.Count == 0)
            {
                return this.thisDraftedLabels;
            }

            var labels = new HashSet<string>(this.thisDraftedLabels ?? new HashSet<string>(), StringComparer.OrdinalIgnoreCase);

            foreach ((string label, bool changed) in this.thisChangedHere)
            {
                if (changed)
                {
                    labels.Add(label);
                }
                else
                {
                    labels.Remove(label);
                }
            }

            return labels;
        }
    }
}
