using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Modal dialog collecting the two fields a workbook needs (a description and an optional
    // note). Serves two purposes with the same layout, matching WorklogAddLinkWindow's own
    // add/edit split:
    //   - creating a new workbook (the default): the id shown is only a preview of what
    //     WorklogManager.CreateWorkbook will allocate, and OnCreateClick calls that - on the device
    //     picked under "On my device", which can be a new one (CRT 4.0 step 4). Returns the
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
            this.thisCreateTitle = this.Title;

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
        //
        // It asks which of the user's devices the repair is on (CRT 4.0 step 4 - the plan: "one of
        // yours on this board, a new one, or none"): "None" and the devices with this board, and
        // beside them "Add device" - a button, never an item of the list (code review, 2026-10-29:
        // reaching "A new device" with the Down key opened the dialog). `ownedDeviceId` preselected -
        // and so named in the title ("New repair on Breadbin #2", ShowDeviceInTitle).
        // `devices`: the user's devices as the caller read them (Main's refresh, the Workbooks tab's)
        // - else they are read here.
        // ###########################################################################################
        public void Initialize(string boardKey, int? ownedDeviceId = null, IReadOnlyList<OwnedDeviceRecord>? devices = null)
        {
            this.thisBoardKey = boardKey;

            List<OwnedDeviceRecord> onThisBoard = (devices ?? OwnedDeviceStore.GetAll())
                .Where(device => device.Id == ownedDeviceId ||
                                 string.Equals(device.BoardKey, boardKey, StringComparison.OrdinalIgnoreCase))
                .ToList();

            this.FillDeviceChoices(onThisBoard.Select(device => new DeviceChoice(device.Id, device.Name)), ownedDeviceId);
            this.DevicePanel.IsVisible = true;
        }

        // One entry of the device drop-down - its label is what the list shows.
        internal sealed record DeviceChoice(int? Id, string Label)
        {
            public override string ToString() => this.Label;
        }

        // Stands in for "Add device" in a test, which cannot show a dialog - null in the app.
        internal Func<Task<OwnedDeviceRecord?>>? AddDeviceOverrideForTests { get; set; }

        // What "Add device" last started, for a test to wait on.
        internal Task LastNewDeviceForTests { get; private set; } = Task.CompletedTask;

        // Whether a device was added here - it is on disk whether or not a workbook is created.
        internal bool DeviceAdded { get; private set; }

        // "None", then `devices` - `selectId` selected, else "None".
        private void FillDeviceChoices(IEnumerable<DeviceChoice> devices, int? selectId)
        {
            var choices = new List<DeviceChoice> { new(null, OwnedDeviceWording.NoDeviceChoice) };
            choices.AddRange(devices);

            this.DeviceComboBox.ItemsSource = choices;
            this.DeviceComboBox.SelectedItem = choices.FirstOrDefault(choice => choice.Id == selectId) ?? choices[0];
            this.ShowDeviceInTitle();
        }

        private void OnDeviceSelectionChanged(object? sender, SelectionChangedEventArgs e) => this.ShowDeviceInTitle();

        // ###########################################################################################
        // The title names the device the repair goes on - "New repair on Breadbin #2" - else it is the
        // form's own (code review, 2026-10-29: it stayed on the device while the repair went on none).
        // ###########################################################################################
        private void ShowDeviceInTitle()
        {
            if (this.thisEditingWorkbook is null)
            {
                this.Title = this.DeviceComboBox.SelectedItem is DeviceChoice { Id: not null } choice
                    ? OwnedDeviceWording.NewRepairTitle(choice.Label)
                    : this.thisCreateTitle;
            }
        }

        // The form's title when the repair is on no device - the markup's own, read as the window is
        // made, never a second copy of the words (code review, 2026-10-29).
        private readonly string? thisCreateTitle;

        private void OnAddDeviceClick(object? sender, RoutedEventArgs e) => this.LastNewDeviceForTests = this.AddNewDeviceAsync();

        // ###########################################################################################
        // "Add device": the form on THIS board, not to be changed there - and the device it adds picked
        // for the repair. One that comes back on another board is no device for this repair (code
        // review, 2026-10-29: a C64 repair was filed on a C128 device) and is not listed; nothing
        // added, the choice stays as it was. Either way a device added is on disk (DeviceAdded).
        // ###########################################################################################
        private async Task AddNewDeviceAsync()
        {
            OwnedDeviceRecord? added = null;

            try
            {
                if (this.AddDeviceOverrideForTests is { } add)
                {
                    added = await add();
                }
                else
                {
                    var dialog = new OwnedDeviceWindow();
                    dialog.Initialize(DataManager.HardwareBoards, this.thisBoardKey, boardFixed: true);
                    added = await dialog.ShowDialog<OwnedDeviceRecord?>(this);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave the click: App lets such an exception end CRT.
                Logger.Critical($"Adding a device for a new workbook failed - [{ex}]");
            }

            if (added is null)
            {
                return;
            }

            this.DeviceAdded = true;

            if (!string.Equals(added.BoardKey.Trim(), this.thisBoardKey, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            List<DeviceChoice> devices = (this.DeviceComboBox.ItemsSource as IEnumerable<DeviceChoice> ?? [])
                .Where(choice => choice.Id is not null)
                .ToList();

            devices.Add(new DeviceChoice(added.Id, added.Name));
            this.FillDeviceChoices(devices, added.Id);
        }

        // The device picked, or null for none.
        internal int? ChosenDeviceId =>
            this.DevicePanel.IsVisible ? (this.DeviceComboBox.SelectedItem as DeviceChoice)?.Id : null;

        // The record the dialog closed with - for the tests, which cannot read a ShowDialog result.
        internal WorkbookRecord? Created { get; private set; }

        // Types a description and presses the submit button, as a test cannot through ShowDialog.
        internal void CreateForTests(string title)
        {
            this.TitleTextBox.Text = title;
            this.OnCreateClick(null, new RoutedEventArgs());
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

            var record = WorklogManager.CreateWorkbook(this.thisBoardKey, title, note, this.ChosenDeviceId);

            if (record == null)
            {
                // The workbook root was unusable or the write failed - closing with null here would
                // just make the dialog vanish as if the user had cancelled, so say so instead and
                // keep what they typed. The log carries the underlying reason.
                this.TitleValidationText.Text = "Could not create the workbook - see the log for details.";
                this.TitleValidationText.IsVisible = true;
                return;
            }

            this.Created = record;
            this.Close(record);
        }
    }
}
