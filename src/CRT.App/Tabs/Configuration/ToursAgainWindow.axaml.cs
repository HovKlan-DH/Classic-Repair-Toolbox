using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.Tours;

namespace CRT
{
    // ###########################################################################################
    // Said as "Take the introduction tours again" (the "Configuration" tab) is pressed (owner request, 2026-11-15: "show a
    // modal that the use now should navigate to any tab to get started on the intro. Only button is "OK"" - nothing was
    // said until then, the tab staying as it was since 2026-11-14). Main.TakeToursAgain opens it. OK, Enter and Escape
    // close it; nothing else happens there.
    // ###########################################################################################
    public partial class ToursAgainWindow : Window
    {
        public ToursAgainWindow()
        {
            this.InitializeComponent();

            this.Title = TourWording.ToursAgainTitle;
            this.HeadingText.Text = TourWording.ToursAgainTitle;
            this.BodyText.Text = TourWording.ToursAgainBody;

            // Tunnel, not a plain KeyDown - see DiscardDraftWindow.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            this.Opened += (_, _) => this.OkButton.Focus();
        }

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key is Key.Escape or Key.Enter)
            {
                this.Close();
                e.Handled = true;
            }
        }

        private void OnOkClick(object? sender, RoutedEventArgs e) => this.Close();
    }
}
