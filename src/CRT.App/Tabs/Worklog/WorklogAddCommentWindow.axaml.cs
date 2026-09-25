using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace CRT
{
    // ###########################################################################################
    // Tiny modal collecting the text for one "Comments" row. Returns the trimmed text via
    // ShowDialog, or null when cancelled. The date/time is stamped by the caller (DateTime.Now at
    // the moment of Add), not entered here.
    // ###########################################################################################
    public partial class WorklogAddCommentWindow : Window
    {
        public WorklogAddCommentWindow()
        {
            this.InitializeComponent();

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.CommentTextBox.Focus(), DispatcherPriority.Background);

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Switches the dialog into "edit" mode: pre-fills the existing text and relabels the title/
        // submit button, so the same modal serves both "Add comment" and the Comments row's click-
        // to-edit behavior.
        // ###########################################################################################
        public void InitializeForEdit(string text)
        {
            this.Title = "Edit comment";
            this.HeaderText.Text = "Edit comment";
            this.AddButton.Content = "Update comment";
            this.CommentTextBox.Text = text;
        }

        // ###########################################################################################
        // Escape cancels, same as the link dialog. Plain Enter is deliberately left alone (unlike the
        // link dialog) since the comment field is multi-line and Enter is how a user adds a new line -
        // Ctrl+Enter submits instead (see the hint text under the textarea). Handled on the Tunnel
        // route so it fires before CommentTextBox's own AcceptsReturn handling inserts a newline -
        // a bubbling KeyDown handler would run too late to stop that.
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
            string text = this.CommentTextBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(text))
            {
                this.CommentValidationText.IsVisible = true;
                return;
            }

            this.Close(text);
        }
    }
}
