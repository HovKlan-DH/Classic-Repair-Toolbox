using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace CRT
{
    // ###########################################################################################
    // Tiny modal collecting a headline + URL for one "Links of interest" row. Returns the pair via
    // ShowDialog, or null when cancelled. Same shape as CreateWorkbookWindow.
    // ###########################################################################################
    public partial class WorklogAddLinkWindow : Window
    {
        public WorklogAddLinkWindow()
        {
            this.InitializeComponent();

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.HeadlineTextBox.Focus(), DispatcherPriority.Background);

            this.KeyDown += this.OnWindowKeyDown;
        }

        // ###########################################################################################
        // Escape cancels, Enter submits from either field - same as pressing the Cancel/Add(Save)
        // buttons directly.
        // ###########################################################################################
        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                this.OnAddClick(sender, e);
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Switches the dialog into "edit" mode: pre-fills the existing headline/URL and relabels the
        // title/submit button, so the same modal serves both "Add link" and the Links row's edit icon.
        // ###########################################################################################
        public void InitializeForEdit(string headline, string url)
        {
            this.Title = "Edit link";
            this.HeaderText.Text = "Edit link";
            this.AddButton.Content = "Update link";
            this.HeadlineTextBox.Text = headline;
            this.UrlTextBox.Text = url;
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close(null);
        }

        private void OnAddClick(object? sender, RoutedEventArgs e)
        {
            string headline = this.HeadlineTextBox.Text?.Trim() ?? string.Empty;
            string url = this.UrlTextBox.Text?.Trim() ?? string.Empty;

            bool isValid = true;

            if (string.IsNullOrWhiteSpace(headline))
            {
                this.HeadlineValidationText.IsVisible = true;
                isValid = false;
            }

            if (string.IsNullOrWhiteSpace(url))
            {
                this.UrlValidationText.IsVisible = true;
                isValid = false;
            }

            if (!isValid)
                return;

            this.Close(((string Headline, string Url)?)(headline, url));
        }
    }
}
