using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace CRT
{
    // ###########################################################################################
    // The maintainer agreement shown when "Create system" is clicked in NewSystemWindow
    // (maintainer request, 2026-09-24). A system submitted to the community needs someone to
    // review the changes others later submit for it, and the person who created it is the one who
    // knows it - so creating a new system requires accepting that role up front. Returns true via
    // ShowDialog when accepted, null when declined.
    //
    // *** ENTER AND ESCAPE BOTH DECLINE, on the Tunnel route. *** The contributor very likely
    // pressed Enter in the form a moment ago to get here, and a second reflexive Enter must not
    // accept an agreement they have not read. Declining is harmless - it returns to the form with
    // everything they typed still in it - so it is the safe thing for a key to do. Tunnel for the
    // same reason DiscardDraftWindow and DeleteWorkbookWindow use it: a focused Button handles
    // Enter itself on the bubbling route, which would let Enter on a focused "Accept and create"
    // accept after all.
    //
    // The words "system maintainer" match the server's own role (the `maintainers` table - the
    // people who may approve work on a system); see "Vocabulary the user reads" in CLAUDE.md
    // before renaming it.
    // ###########################################################################################
    public partial class NewSystemMaintainerWindow : Window
    {
        public NewSystemMaintainerWindow()
        {
            this.InitializeComponent();

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // Names the system the agreement is about, on its own bold line.
        public void Initialize(string systemDisplayName)
        {
            this.SystemNameText.Text = systemDisplayName;
        }

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.OnDeclineClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnDeclineClick(object? sender, RoutedEventArgs e)
        {
            this.Close(null);
        }

        private void OnAcceptClick(object? sender, RoutedEventArgs e)
        {
            this.Close(true);
        }
    }
}
