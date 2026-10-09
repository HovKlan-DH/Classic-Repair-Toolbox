using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace CRT
{
    // ###########################################################################################
    // "Edit board as draft" on a board that already has a draft (owner request, 2026-10-09: "If it
    // exists already, it should be informed via popup, that there cannot be two draft for same
    // board"). Says so, names the board on its own bold line, and offers the draft there is.
    //
    // Returns true via ShowDialog for "Open the draft", or null for Close, Escape or the title bar.
    // ###########################################################################################
    public partial class DraftExistsWindow : Window
    {
        public DraftExistsWindow()
        {
            this.InitializeComponent();

            // Escape closes from anywhere - tunnel, since a focused Button would handle the key first.
            // Enter is the buttons' own: "Open the draft" is the default, and a focused "Close"
            // answers Enter itself (code review, 2026-10-09: a tunnel handler made Enter open the
            // draft even with Close focused).
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        public void Initialize(string boardDisplayName)
        {
            this.BoardNameText.Text = boardDisplayName;
        }

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCloseClick(sender, e);
                e.Handled = true;
            }
        }

        // Whether "Open the draft" was chosen - what ShowDialog returns, readable by a test that
        // only shows the window (ShowDialog blocks headlessly).
        internal bool OpenChosen { get; private set; }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnOpenDraftClick(object? sender, RoutedEventArgs e)
        {
            this.OpenChosen = true;
            this.Close(true);
        }
    }
}
