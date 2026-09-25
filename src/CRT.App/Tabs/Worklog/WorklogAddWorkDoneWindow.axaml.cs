using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Tiny modal collecting one "Work done" row: a description plus hours spent and cost. Returns
    // the triple via ShowDialog, or null when cancelled. The date/time is stamped by the caller
    // (DateTime.Now at the moment of Add), not entered here.
    // ###########################################################################################
    public partial class WorklogAddWorkDoneWindow : Window
    {
        public WorklogAddWorkDoneWindow()
        {
            this.InitializeComponent();

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.DescriptionTextBox.Focus(), DispatcherPriority.Background);

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            // The one field in the app where a cost is TYPED, so it names the currency it will be
            // recorded in - "Cost (DKK)". Every surface that later displays the figure appends the
            // same code, and a number entered without knowing which currency was assumed is the
            // thing that makes an exported invoice wrong.
            //
            // Set from code rather than in the markup because the code is a user setting: the
            // markup would have to carry a literal, which would then be right for one user only.
            this.CostLabelText.Text = $"Cost ({UserSettings.WorklogCurrencyCode})";

            // The dialog opens on zero, so this collapses the line rather than showing anything -
            // but it is called anyway so the control's state is decided in ONE place, and an
            // edit-mode open (InitializeForEdit, which sets Value before the window is shown) does
            // not depend on ValueChanged having fired.
            this.UpdateHoursReadback();
        }

        // ###########################################################################################
        // Switches the dialog into "edit" mode: pre-fills the existing description/hours/cost and
        // relabels the title/submit button, so the same modal serves both "Add work" and the Work
        // done row's click-to-edit behavior.
        // ###########################################################################################
        public void InitializeForEdit(string text, double hoursSpent, double cost)
        {
            this.Title = "Edit work done";
            this.HeaderText.Text = "Edit work done";
            this.AddButton.Content = "Update work";
            this.DescriptionTextBox.Text = text;
            this.HoursNumericUpDown.Value = (decimal)hoursSpent;
            this.CostNumericUpDown.Value = (decimal)cost;

            // Setting Value above raises ValueChanged, which refreshes this already - but only when
            // the value actually MOVED. Editing a row recorded as 0 hours assigns 0 over 0, which
            // raises nothing, so the readback is refreshed explicitly here too.
            this.UpdateHoursReadback();
        }

        // ###########################################################################################
        // Echoes the decimal hours back in words as they are typed. NumericUpDown raises this on
        // every accepted value - the spinner buttons, typing, and the programmatic assignment in
        // InitializeForEdit - so the line can never disagree with the number above it.
        // ###########################################################################################
        private void OnHoursValueChanged(object? sender, NumericUpDownValueChangedEventArgs e)
        {
            this.UpdateHoursReadback();
        }

        // ###########################################################################################
        // Rewrites the "1 hour and 15 minutes" line under the hours field, and hides it outright
        // when there is nothing to say (an untouched field, or a value under half a minute).
        //
        // The NUMBERS are bold and the words are not, which is why this walks
        // WorklogDurationFormatter's parts rather than taking its finished string - a TextBlock
        // cannot mix weights within one Text. Note that this leaves Text null with the content in
        // Inlines: a block carrying both renders the Text and silently ignores the Inlines, the
        // same trap TextLinkRenderer and the Workbooks summary strip document.
        // ###########################################################################################
        private void UpdateHoursReadback()
        {
            double hours = (double?)this.HoursNumericUpDown.Value ?? 0.0;

            var parts = WorklogDurationFormatter.BuildParts(hours);

            this.HoursReadbackText.Inlines?.Clear();

            if (parts.Count == 0)
            {
                // Collapsed rather than blanked: an empty line here would push the Cancel/Update
                // row down by its own height for no reason, so the dialog would jump as soon as a
                // value was typed and back again when it was cleared.
                this.HoursReadbackText.IsVisible = false;
                return;
            }

            foreach (var part in parts)
            {
                this.HoursReadbackText.Inlines!.Add(new Run(part.Number) { FontWeight = FontWeight.Bold });
                this.HoursReadbackText.Inlines!.Add(new Run(part.Words));
            }

            this.HoursReadbackText.IsVisible = true;
        }

        // ###########################################################################################
        // Escape cancels, same as the comment dialog. Plain Enter is deliberately left alone since the
        // description field is multi-line - Ctrl+Enter submits instead (see the hint text under the
        // textarea). Handled on the Tunnel route so it fires before DescriptionTextBox's own
        // AcceptsReturn handling inserts a newline - a bubbling KeyDown handler would run too late.
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
                this.OnAddClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close(null);
        }

        private void OnAddClick(object? sender, RoutedEventArgs e)
        {
            string text = this.DescriptionTextBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(text))
            {
                this.DescriptionValidationText.IsVisible = true;
                return;
            }

            double hoursSpent = (double?)this.HoursNumericUpDown.Value ?? 0.0;
            double cost = (double?)this.CostNumericUpDown.Value ?? 0.0;

            this.Close(((string Text, double HoursSpent, double Cost)?)(text, hoursSpent, cost));
        }
    }
}
