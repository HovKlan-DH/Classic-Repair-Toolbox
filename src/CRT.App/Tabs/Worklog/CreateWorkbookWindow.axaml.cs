using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Modal dialog collecting the two fields a workbook needs (a description and an optional
    // note). Serves two purposes with the same layout, matching WorklogAddLinkWindow's own
    // add/edit split:
    //   - creating a new workbook (the default): the id shown is only a preview of what
    //     WorklogManager.CreateWorkbook will allocate, and OnCreateClick calls that. Returns the
    //     created WorkbookRecord via ShowDialog, or null when cancelled.
    //   - editing an existing one, via InitializeForEdit: the real id is shown, the fields are
    //     pre-filled, and OnCreateClick calls WorklogManager.UpdateWorkbook instead. Returns the
    //     UPDATED WorkbookRecord (title/note only - id, board key, status, start date and entry
    //     count come from thisEditingWorkbook untouched), or null when cancelled.
    // ###########################################################################################
    public partial class CreateWorkbookWindow : Window
    {
        private string thisBoardKey = string.Empty;

        // Non-null in edit mode, set by InitializeForEdit. Carries every field OnCreateClick does
        // not collect (id, board key, status, start date, entry count) so the returned record is
        // complete rather than a partial one the caller has to patch up.
        private WorkbookRecord? thisEditingWorkbook;

        public CreateWorkbookWindow()
        {
            this.InitializeComponent();

            this.WorkbookIdPreviewText.Text = $"#{WorklogManager.PeekNextId()}";

            // Starts disabled: the description is blank until the user types (or InitializeForEdit
            // pre-fills it), matching the same "no title, no save" gate the worklog entry card's
            // own save button uses (TabSchematics.Worklog.cs's UpdateWorklogEntryCardSaveEnabled).
            this.UpdateCreateButtonEnabled();

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.TitleTextBox.Focus(), DispatcherPriority.Background);

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Must be called before showing the dialog in CREATE mode: sets the board the new workbook
        // belongs to. Not used in edit mode - the workbook being edited already has a board.
        // ###########################################################################################
        public void Initialize(string boardKey)
        {
            this.thisBoardKey = boardKey;
        }

        // ###########################################################################################
        // Switches the dialog into "edit" mode: shows the real id instead of the next-id preview,
        // pre-fills the existing description/note, and relabels the title/submit button - the same
        // idea as WorklogAddLinkWindow.InitializeForEdit, so "Edit workbook" opens exactly the
        // dialog "Create new workbook" does rather than a second, separately-maintained one.
        // ###########################################################################################
        public void InitializeForEdit(WorkbookRecord workbook)
        {
            this.thisEditingWorkbook = workbook;
            this.thisBoardKey = workbook.BoardKey;

            this.Title = "Edit workbook";
            this.WorkbookIdPreviewText.Text = $"#{workbook.Id}";
            this.TitleTextBox.Text = workbook.Title;
            this.NoteTextBox.Text = workbook.Note;
            this.CreateButton.Content = "Update workbook";

            // TitleTextBox.Text above already fires OnTitleTextChanged (the control is part of the
            // visual tree by the time InitializeForEdit runs, called after InitializeComponent), so
            // this is only needed for the edge case of an existing workbook whose title is blank -
            // otherwise redundant, but cheap and keeps this method correct on its own rather than
            // relying on that ordering.
            this.UpdateCreateButtonEnabled();
        }

        // ###########################################################################################
        // Gates the submit button on a non-blank description, the same rule OnCreateClick itself
        // enforces (and the same pattern TabSchematics.Worklog.cs's own save button uses) - so
        // clearing the field (in either Create or Edit mode) disables the button instead of letting
        // the user click through to the "A description is required" validation message.
        // ###########################################################################################
        private void OnTitleTextChanged(object? sender, TextChangedEventArgs e)
        {
            this.UpdateCreateButtonEnabled();
        }

        private void UpdateCreateButtonEnabled()
        {
            this.CreateButton.IsEnabled = !string.IsNullOrWhiteSpace(this.TitleTextBox.Text);
        }

        // ###########################################################################################
        // Escape cancels, same as every other worklog dialog. Plain Enter is deliberately left alone -
        // NoteTextBox is multi-line (AcceptsReturn) and Enter is how a user adds a line to it - so
        // submitting from the keyboard needs Ctrl+Enter instead (see the hint text under NoteTextBox),
        // same as WorklogAddCommentWindow/WorklogAddWorkDoneWindow/WorklogAddLinkWindow. Handled on
        // the Tunnel route so it fires before NoteTextBox's own AcceptsReturn handling would otherwise
        // insert a newline instead - a bubbling KeyDown handler would run too late to stop that.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter && e.KeyModifiers.HasFlag(KeyModifiers.Control))
            {
                this.OnCreateClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close(null);
        }

        private void OnCreateClick(object? sender, RoutedEventArgs e)
        {
            string title = this.TitleTextBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(title))
            {
                this.TitleValidationText.Text = "A description is required.";
                this.TitleValidationText.IsVisible = true;
                this.TitleTextBox.Focus();
                return;
            }

            this.TitleValidationText.IsVisible = false;

            string note = this.NoteTextBox.Text?.Trim() ?? string.Empty;

            if (this.thisEditingWorkbook != null)
            {
                // The record that actually reached disk is what closes this dialog - no patching up
                // of the caller's own copy, which would apply the trim rules a second time and could
                // hand back something the file does not say.
                var updated = WorklogManager.UpdateWorkbook(this.thisEditingWorkbook.Id, title, note);

                if (updated == null)
                {
                    this.TitleValidationText.Text = "Could not save the workbook - see the log for details.";
                    this.TitleValidationText.IsVisible = true;
                    return;
                }

                this.Close(updated);
                return;
            }

            var record = WorklogManager.CreateWorkbook(this.thisBoardKey, title, note);

            if (record == null)
            {
                // The workbook root was unusable or the write failed - closing with null here would
                // just make the dialog vanish as if the user had cancelled, so say so instead and
                // keep what they typed. The log carries the underlying reason.
                this.TitleValidationText.Text = "Could not create the workbook - see the log for details.";
                this.TitleValidationText.IsVisible = true;
                return;
            }

            this.Close(record);
        }
    }
}
