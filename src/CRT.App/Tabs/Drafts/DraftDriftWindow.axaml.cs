using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Theming;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // "What changed officially" - shows how one board's drafted rows line up against the official
    // data as it stands now, and lets the contributor dismiss the warning once they have looked
    // (NewContributeStrategy.md Phase 2, session 2d).
    //
    // The report itself is built by DraftDriftDetector in CRT.Data; this window only renders it and
    // owns the one action. See that class for why a genuine official-versus-official diff is
    // impossible - the old official workbook is gone once sync has run, so what can honestly be
    // shown is how each of the contributor's OWN rows stands against the current official data.
    // ###########################################################################################
    public partial class DraftDriftWindow : Window
    {
        private string thisExcelDataFile = string.Empty;
        private string thisOfficialRevision = string.Empty;

        public ObservableCollection<DriftRowItem> AttentionRows { get; } = new();
        public ObservableCollection<DriftRowItem> OtherRows { get; } = new();

        public DraftDriftWindow()
        {
            this.InitializeComponent();

            this.AttentionItemsControl.ItemsSource = this.AttentionRows;
            this.OtherItemsControl.ItemsSource = this.OtherRows;

            this.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Renders one report. Rows that carry a consequence are shown FIRST and separately, because
        // they are the only ones the contributor may need to act on - the rest are there so the
        // view is complete rather than because anything is wrong with them.
        // ###########################################################################################
        public void Initialize(string displayName, string excelDataFile, DraftChangeReport report)
        {
            this.thisExcelDataFile = excelDataFile;
            this.thisOfficialRevision = report.OfficialRevision;

            this.HeaderText.Text = displayName;

            this.SummaryText.Text = report.State switch
            {
                DraftDriftState.OfficialIsNewer =>
                    "The official data for this board has been updated since you started your draft. Your edits are still applied on top of it.",
                DraftDriftState.Changed =>
                    "The official data for this board has changed since you started your draft. Your edits are still applied on top of it.",
                DraftDriftState.InSync =>
                    "The official data for this board has not changed since you started your draft.",
                _ =>
                    "CRT does not know which revision this draft was started from, so it cannot tell whether the official data has changed.",
            };

            // "Not recorded" rather than an empty line: a draft written before every save path
            // stamped a base revision genuinely has none, and a blank here would read as a bug.
            this.BaseRevisionText.Text = string.IsNullOrWhiteSpace(report.BaseRevision)
                ? "Not recorded"
                : report.BaseRevision;

            this.OfficialRevisionText.Text = string.IsNullOrWhiteSpace(report.OfficialRevision)
                ? "Not recorded"
                : report.OfficialRevision;

            this.AttentionRows.Clear();
            this.OtherRows.Clear();

            // ###########################################################################################
            // *** ONE LIST NOW, NOT TWO (Phase 6, 2026-09-23). ***
            //
            // The "needs attention" band held rows whose edit had been silently defeated by the
            // merge - a change to a row that had since vanished officially, or an addition the
            // official data had since made too. Neither can happen without a merge, so nothing
            // would ever populate that band; leaving it in place would leave a heading on screen
            // that is permanently empty.
            //
            // The panel itself is kept (and stays hidden) rather than torn out of the markup, so
            // the day drift detection grows a sharper answer it has somewhere to go.
            // ###########################################################################################
            foreach (BoardRowChange row in report.Rows)
            {
                this.OtherRows.Add(DriftRowItem.From(row));
            }

            this.AttentionPanel.IsVisible = false;

            this.OtherPanel.IsVisible = this.OtherRows.Count > 0;
            this.OtherHeaderText.Text = this.OtherRows.Count == 1
                ? "1 change of yours"
                : $"{this.OtherRows.Count} changes of yours";

            this.NoRowsText.IsVisible = report.Rows.Count == 0;

            // Nothing to dismiss when there is no warning to begin with.
            this.RebaseButton.IsVisible =
                report.State is DraftDriftState.OfficialIsNewer or DraftDriftState.Changed;
        }

        // ###########################################################################################
        // Records that the contributor has seen this, by moving the draft's base revision to the
        // current official one. It writes ONE STRING and touches no drafted row - see
        // DraftBaseRevision.Rebase for the full list of what it deliberately does not do.
        //
        // No confirmation dialog: nothing is destroyed and it is reversible in effect, since the
        // next sync that moves the official revision raises the warning again. DiscardDraftWindow
        // exists because discarding is permanent; this is not.
        // ###########################################################################################
        private void OnRebaseClick(object? sender, RoutedEventArgs e)
        {
            // *** REBASING WRITES THE MARKER, NOT THE BOARD (Phase 6, 2026-09-23). *** It is a
            // statement about WHICH published revision the draft is measured against, and it
            // deliberately changes not one row - which is exactly why the status line below can
            // promise the contributor their work is untouched.
            if (!DraftWorkbookStore.SetBaseRevision(
                    DraftManager.DraftsRoot,
                    this.thisExcelDataFile,
                    this.thisOfficialRevision))
            {
                this.ShowStatus("Could not read this board's draft.", isError: true);
                return;
            }

            this.RebaseButton.IsVisible = false;
            this.ShowStatus("Noted. Your drafted rows are unchanged.");
        }

        private void ShowStatus(string message, bool isError = false)
        {
            this.StatusText.Text = message;
            this.StatusText.Foreground = ThemeResources.ResolveBrush(isError ? "Text_Fail_Fg" : "Text_Success_Fg");
        }

        internal string StatusTextForTests => this.StatusText.Text ?? string.Empty;

        internal bool RebaseButtonVisibleForTests => this.RebaseButton.IsVisible;

        // The rebase is a Click handler rather than a Command, and a window never attached to a
        // visual tree cannot be clicked - so the tests drive the handler directly. It is the same
        // method the button invokes, so this exercises the shipped path rather than a copy of it.
        internal void RaiseRebaseForTests() => this.OnRebaseClick(this, new RoutedEventArgs());

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
    // One drafted row as this window shows it: what it is, and what has happened to it.
    // ###########################################################################################
    public sealed class DriftRowItem
    {
        public string Title { get; init; } = string.Empty;
        public string Explanation { get; init; } = string.Empty;

        public static DriftRowItem From(BoardRowChange row) => new()
        {
            Title = $"{row.Section}: {row.DisplayLabel}",
            Explanation = ExplanationFor(row),
        };

        // ###########################################################################################
        // Plain statements of what the contributor did, not labels.
        //
        // *** THESE NO LONGER DESCRIBE A CONSEQUENCE, and the change is deliberate. *** The old
        // wordings said what the MERGE had done with each row ("yours is not applied while both
        // exist"), which was the only way to surface an outcome the contributor could not
        // otherwise see. There is no merge: the draft workbook is the board, so a change is
        // simply in effect and saying anything more would be inventing a caveat.
        //
        // A Modified row names the fields that differ, because "U8 changed" is far less useful
        // than "U8: Description, Part number" and the information is free.
        // ###########################################################################################
        private static string ExplanationFor(BoardRowChange row) => row.Kind switch
        {
            BoardRowChangeKind.Added =>
                "You added this. It is not in the published data.",
            BoardRowChangeKind.Deleted =>
                "You removed this. It is still in the published data, and your submission would remove it.",
            _ => row.ChangedFields.Count > 0
                ? "You changed: " + string.Join(", ", row.ChangedFields)
                : "You changed this.",
        };
    }
}
