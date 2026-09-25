using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Confirmation modal for "Delete worklog" on an entry card in the Workbooks tab - deleting a
    // worklog removes its row from entries.json AND its whole attachment folder, photos and files
    // included (see WorklogManager.DeleteEntry), so this asks before it happens rather than acting
    // on a single click the way the entry editor's own sub-item deletes (a link, a comment, a
    // photo) do. Returns true via ShowDialog when the user confirms, or null when cancelled.
    //
    // A near-twin of DeleteWorkbookWindow, deliberately kept as its own window rather than merged
    // into a shared "confirm a delete" dialog with swappable strings: the two say different things
    // about different objects, and the point of the copy is that a change to one cannot silently
    // change what the other promises about a permanent delete.
    // ###########################################################################################
    public partial class DeleteWorklogWindow : Window
    {
        public DeleteWorklogWindow()
        {
            this.InitializeComponent();

            // Tunnel, NOT a plain KeyDown subscription. A focused Button handles Enter itself and
            // marks the event handled, so a bubbling handler never runs - which would mean tabbing
            // to (or clicking, and so focusing) "Delete worklog" and pressing Enter DELETED the
            // worklog, the exact opposite of what this window promises below. On the tunnel route
            // this sees the key first and can stop it. Same reasoning, and the same fix, as
            // DeleteWorkbookWindow and the worklog dialogs that must beat a multi-line TextBox to
            // Enter.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Names the worklog being deleted, on its own bold line, as "#{N} · {Title}" - the same
        // "#N" the entry's card, its board pill and the exported PDF all show, so the thing named
        // here is recognisably the thing the user clicked. Several cards are on screen at once, and
        // "are you sure?" alone does not say which one is about to be lost.
        // ###########################################################################################
        public void Initialize(WorklogEntryRecord entry)
        {
            string title = string.IsNullOrWhiteSpace(entry.Title) ? "(untitled)" : entry.Title;
            this.WorklogNameText.Text = $"#{entry.Id} · {title}";
        }

        // ###########################################################################################
        // Escape AND Enter both cancel - deliberately unlike every other modal in the app (which
        // submits on Enter), because this one's "submit" is a permanent delete. Defaulting Enter to
        // Cancel means a reflexive keypress (dismissing what looks like "just another dialog") can
        // never destroy a worklog; deleting stays a click on the button itself.
        //
        // Runs on the Tunnel route (see the constructor) so it beats a focused Button's own Enter
        // handling. Subscribed the ordinary bubbling way, this would never fire once the Delete
        // button had focus and Enter would delete the worklog.
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
