using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace CRT
{
    // ###########################################################################################
    // Confirmation modal for "Discard draft" on the Drafts tab - discarding a draft removes every
    // local, unpublished edit to a system permanently (see DraftManager.DiscardDraft), the same
    // shape and the same reasoning as DeleteWorkbookWindow. Returns true via ShowDialog when the
    // user confirms, or null when cancelled.
    // ###########################################################################################
    public partial class DiscardDraftWindow : Window
    {
        public DiscardDraftWindow()
        {
            this.InitializeComponent();

            // Tunnel, NOT a plain KeyDown subscription - see DeleteWorkbookWindow's own comment on
            // this exact trap: a focused Button handles Enter itself and marks the event handled,
            // so a bubbling handler never runs, which would let a reflexive Enter after clicking
            // (or tabbing to) "Discard draft" actually discard it.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Names the system being discarded, on its own bold line, so a user with several boards'
        // worth of drafts cannot mistake which one they are about to lose.
        // ###########################################################################################
        public void Initialize(string systemDisplayName)
        {
            this.SystemNameText.Text = systemDisplayName;
        }

        // ###########################################################################################
        // Escape AND Enter both cancel - deliberately unlike every other modal in the app, because
        // this one's "submit" is a permanent discard. See DeleteWorkbookWindow's own comment: this
        // must run on the Tunnel route to beat a focused Button's own Enter handling.
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

        private void OnDiscardClick(object? sender, RoutedEventArgs e)
        {
            this.Close(true);
        }
    }
}
