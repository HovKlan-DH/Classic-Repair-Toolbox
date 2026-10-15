using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "What changed in the published data" - one drafted board against the published data as it
    // is now, and the one action that settles it (NewContributeStrategy.md Phase 2, session 2d;
    // rebuilt around the draft's base on 2026-10-13).
    //
    // The report is built in CRT.Data (DraftChangeReport.For) and the words are DraftDriftWording;
    // this window renders them and owns the action:
    //
    //   - a draft WITH a base: "Bring my draft up to date" - DraftRebaser carries the contributor's
    //     changes onto the published update. Every change of theirs is kept; rows both changed are
    //     listed first, before the button is pressed, so nothing is decided blind;
    //   - a draft WITHOUT one (seeded before 2026-10-13): "I have looked at this" quiets the warning
    //     and changes nothing else - not even the revision the server is told
    //     (DraftWorkbookStore.MarkDriftSeen).
    //
    // *** IT USED TO SAY "YOUR EDITS ARE STILL APPLIED ON TOP OF IT" (code review, 2026-10-13). ***
    // That was untrue since drafts became complete copies (Phase 6): a stale draft sent as it was
    // asked the maintainer to remove everything published since. See DraftRebase.
    // ###########################################################################################
    public partial class DraftDriftWindow : Window
    {
        private string thisExcelDataFile = string.Empty;
        private string thisOfficialRevision = string.Empty;
        private bool thisCanBringUpToDate;

        // Whether the Drafts tab's table holds unsaved edits of this draft - then nothing may write
        // the draft (CLAUDE.md: "No other editor writes a draft under a table with unsaved edits").
        private Func<string, UnsavedTableEditsPrompt?>? thisSavingBlockedFor;

        public ObservableCollection<DriftRowItem> AttentionRows { get; } = new();
        public ObservableCollection<DriftRowItem> OtherRows { get; } = new();
        public ObservableCollection<DriftRowItem> UpdateRows { get; } = new();

        // True once "Bring my draft up to date" changed the draft - the caller reloads the board on
        // screen and the Drafts list.
        public bool BroughtUpToDate { get; private set; }

        public DraftDriftWindow()
        {
            this.InitializeComponent();

            this.AttentionItemsControl.ItemsSource = this.AttentionRows;
            this.OtherItemsControl.ItemsSource = this.OtherRows;
            this.UpdateItemsControl.ItemsSource = this.UpdateRows;

            this.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);
        }

        public void Initialize(
            string displayName,
            string excelDataFile,
            DraftChangeReport report,
            Func<string, UnsavedTableEditsPrompt?>? savingBlockedFor = null)
        {
            ArgumentNullException.ThrowIfNull(report);

            this.thisExcelDataFile = excelDataFile;
            this.thisOfficialRevision = report.OfficialRevision;
            this.thisCanBringUpToDate = report.CanBringUpToDate;
            this.thisSavingBlockedFor = savingBlockedFor;

            this.HeaderText.Text = displayName;
            this.SummaryText.Text = DraftDriftWording.Summary(report.State, report.CanBringUpToDate);
            this.ComparedText.Text = DraftDriftWording.Compared(report.ComparedWithBase);

            // "Not recorded" rather than an empty line: a draft written before every save path
            // stamped a base revision genuinely has none, and a blank here would read as a bug.
            this.BaseRevisionText.Text = string.IsNullOrWhiteSpace(report.BaseRevision) ? "Not recorded" : report.BaseRevision;
            this.OfficialRevisionText.Text = string.IsNullOrWhiteSpace(report.OfficialRevision) ? "Not recorded" : report.OfficialRevision;

            this.AttentionRows.Clear();
            this.OtherRows.Clear();
            this.UpdateRows.Clear();

            foreach (DraftRebaseConflict conflict in report.Conflicts)
            {
                this.AttentionRows.Add(new DriftRowItem { Title = $"{conflict.Section}: {conflict.Row}", Explanation = conflict.What });
            }

            foreach (BoardRowChange row in report.Rows)
            {
                this.OtherRows.Add(DriftRowItem.From(row, DraftDriftWording.OwnRow(row, report.ComparedWithBase)));
            }

            foreach (BoardRowChange row in report.UpdateRows)
            {
                this.UpdateRows.Add(DriftRowItem.From(row, DraftDriftWording.UpdateRow(row)));
            }

            this.AttentionPanel.IsVisible = this.AttentionRows.Count > 0;
            this.AttentionHeaderText.Text = DraftDriftWording.ConflictHeader(this.AttentionRows.Count);

            this.OtherPanel.IsVisible = this.OtherRows.Count > 0;
            this.OtherHeaderText.Text = DraftDriftWording.OwnHeader(this.OtherRows.Count, report.ComparedWithBase);

            this.UpdatePanel.IsVisible = this.UpdateRows.Count > 0;
            this.UpdateHeaderText.Text = DraftDriftWording.UpdateHeader(this.UpdateRows.Count);

            this.NoRowsText.IsVisible = report.Rows.Count == 0;

            // Nothing to act on when there is no warning to begin with.
            this.RebaseButton.IsVisible = DraftDriftWording.IsWarning(report.State);
            this.RebaseButton.Content = report.CanBringUpToDate ? DraftDriftWording.BringUpToDateButton : DraftDriftWording.SeenButton;
            ToolTip.SetTip(this.RebaseButton, report.CanBringUpToDate ? DraftDriftWording.BringUpToDateTip : DraftDriftWording.SeenTip);
        }

        // ###########################################################################################
        // The one action - see the class header for which it is. No confirmation: bringing a draft
        // up to date keeps every change of the contributor's (and lists the rows both changed
        // first), and "I have looked at this" changes nothing in the draft at all.
        // ###########################################################################################
        private async void OnRebaseClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.RebaseAsync();
            }
            catch (Exception ex)
            {
                // An async void handler: nothing may leave it (App lets such an exception end CRT).
                Logger.Critical($"Bringing the draft for [{this.thisExcelDataFile}] up to date failed - [{ex}]");
                this.ShowStatus("Something unexpected stopped this. Nothing was changed.", isError: true);
            }
        }

        internal async Task RebaseAsync()
        {
            if (!this.thisCanBringUpToDate)
            {
                // With the published content as it is now, which a draft judged by content is quiet
                // for (code review, 2026-10-13: the revision alone was never read for such a draft).
                string publishedContent = DraftBaseStore.PublishedContentHash(
                    DraftBoardSource.PublishedPathOf(DataManager.DataRoot, this.thisExcelDataFile));

                if (!DraftWorkbookStore.MarkDriftSeen(DraftManager.DraftsRoot, this.thisExcelDataFile, this.thisOfficialRevision, publishedContent))
                {
                    this.ShowStatus("Could not read this board's draft.", isError: true);
                    return;
                }

                this.RebaseButton.IsVisible = false;
                this.ShowStatus(DraftDriftWording.SeenDone);
                return;
            }

            if (this.thisSavingBlockedFor?.Invoke(this.thisExcelDataFile) is UnsavedTableEditsPrompt blocked)
            {
                await (this.ShowSavingBlockedOverrideForTests?.Invoke(blocked) ?? UnsavedTableEditsWindow.ShowSavingBlockedAsync(this, blocked));
                return;
            }

            string draftsRoot = DraftManager.DraftsRoot;
            string dataRoot = DataManager.DataRoot;
            string excelDataFile = this.thisExcelDataFile;

            DraftRebaseReport report = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.BringingDraftUpToDate,
                () => Task.Run(() => DraftRebaser.Rebase(draftsRoot, dataRoot, excelDataFile)));

            bool success = DraftDriftWording.IsSuccess(report);

            // The board reader's cache answers only while a workbook is as it was read, so the
            // rewritten draft is read afresh by whoever asks next - nothing to clear here.
            // NotNoted wrote the draft too - only its new base could not be kept.
            this.BroughtUpToDate = report.Outcome is DraftRebaseOutcome.BroughtUpToDate or DraftRebaseOutcome.NotNoted;

            if (success)
            {
                this.RebaseButton.IsVisible = false;
                this.SummaryText.Text = DraftDriftWording.Summary(DraftDriftState.InSync, canBringUpToDate: false);
                this.UpdatePanel.IsVisible = false;
            }

            this.ShowStatus(DraftDriftWording.Outcome(report), isError: !success);
        }

        private void ShowStatus(string message, bool isError = false)
        {
            this.StatusText.Text = message;
            this.StatusText.Foreground = ThemeResources.ResolveBrush(isError ? "Text_Fail_Fg" : "Text_Success_Fg");
        }

        internal Func<UnsavedTableEditsPrompt, Task>? ShowSavingBlockedOverrideForTests { get; set; }

        internal string StatusTextForTests => this.StatusText.Text ?? string.Empty;

        internal bool RebaseButtonVisibleForTests => this.RebaseButton.IsVisible;

        internal string RebaseButtonTextForTests => this.RebaseButton.Content as string ?? string.Empty;

        // The action is a Click handler, and a window never attached to a visual tree cannot be
        // clicked - so the tests drive the same method the handler awaits.
        internal void RaiseRebaseForTests() => this.RebaseAsync().GetAwaiter().GetResult();

        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();
    }

    // ###########################################################################################
    // One row as this window shows it: what it is, and what happened to it (DraftDriftWording).
    // ###########################################################################################
    public sealed class DriftRowItem
    {
        public string Title { get; init; } = string.Empty;
        public string Explanation { get; init; } = string.Empty;

        public static DriftRowItem From(BoardRowChange row, string explanation) => new()
        {
            Title = $"{row.Section}: {row.DisplayLabel}",
            Explanation = explanation,
        };
    }
}
