using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Confirmation modal for "Delete workbook" on the Workbooks tab - deleting a workbook removes
    // its entries, photos and files permanently (see WorklogManager.DeleteWorkbook), so this asks
    // before it happens rather than acting on a single click the way the entry editor's own
    // sub-item deletes (a link, a comment, a photo) do. Returns true via ShowDialog when the user
    // confirms, or null when cancelled.
    // ###########################################################################################
    public partial class DeleteWorkbookWindow : Window
    {
        public DeleteWorkbookWindow()
        {
            this.InitializeComponent();

            // Tunnel, NOT a plain KeyDown subscription. A focused Button handles Enter itself and
            // marks the event handled, so a bubbling handler never runs - which meant tabbing to (or
            // clicking, and so focusing) "Delete workbook" and pressing Enter DELETED the workbook,
            // the exact opposite of what this window promises below. On the tunnel route this sees
            // the key first and can stop it. Same reasoning, and the same fix, as the worklog
            // dialogs that must beat a multi-line TextBox to Enter.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Names the workbook being deleted, on its own bold line, so a user with several cards open
        // at once cannot mistake which one they are about to lose.
        // ###########################################################################################
        public void Initialize(WorkbookRecord workbook)
        {
            string title = string.IsNullOrWhiteSpace(workbook.Title) ? "(untitled)" : workbook.Title;
            this.WorkbookNameText.Text = $"#{workbook.Id} · {title}";
        }

        // ###########################################################################################
        // Escape AND Enter both cancel - deliberately unlike every other modal in the app (which
        // submits on Enter), because this one's "submit" is a permanent delete. Defaulting Enter to
        // Cancel means a reflexive keypress (dismissing what looks like "just another dialog") can
        // never destroy a workbook; deleting stays a click on the button itself.
        //
        // Runs on the Tunnel route (see the constructor) so it beats a focused Button's own Enter
        // handling. Subscribed the ordinary bubbling way, this never fired once the Delete button had
        // focus and Enter deleted the workbook.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close(null);
        }

        private void OnDeleteClick(object? sender, RoutedEventArgs e)
        {
            this.Close(true);
        }
    }
}
