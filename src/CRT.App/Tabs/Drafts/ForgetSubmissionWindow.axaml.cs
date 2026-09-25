using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace CRT
{
    // ###########################################################################################
    // Confirmation modal for "Remove" in the My submissions window - the same shape and the same
    // reasoning as DiscardDraftWindow, and kept as its own window rather than merged into a shared
    // "confirm something permanent" dialog for exactly the reason DeleteWorklogWindow documents:
    // the two promise different things, and the point of the copy is that changing one cannot
    // silently change what the other says.
    //
    // WHAT IS PERMANENT HERE IS SMALLER THAN IT LOOKS, and the copy has to make that clear.
    // Removing a receipt does NOT withdraw the contribution - it is already on the server and will
    // still be reviewed. What is lost is the capability token, and with it the ability to check on
    // that submission from inside CRT on this computer. The server stores only the token's hash,
    // so nobody can hand it back, which is why this asks at all rather than just doing it.
    //
    // Returns true via ShowDialog when confirmed, or null when cancelled.
    // ###########################################################################################
    public partial class ForgetSubmissionWindow : Window
    {
        public ForgetSubmissionWindow()
        {
            this.InitializeComponent();

            // Tunnel, not a plain KeyDown - a focused Button handles Enter itself and marks the
            // event handled, so a bubbling handler never runs. See DiscardDraftWindow.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Names the system, on its own bold line - someone with several contributions in the list
        // must not have to work out which row they pressed.
        // ###########################################################################################
        public void Initialize(string systemDisplayName)
        {
            this.SystemNameText.Text = systemDisplayName;
        }

        // ###########################################################################################
        // Escape AND Enter both cancel, like every other irreversible confirmation in this app.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnForgetClick(object? sender, RoutedEventArgs e) => this.Close(true);
    }
}
