using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Asked before "Mark all components automatically from KiCad" on a picture that has marks already
    // (owner request, 2026-11-01; TabDrafts.AskBeforeMarkingAsync). Its own window for the reason
    // DeleteWorklogWindow documents: each says exactly what it does.
    //
    // Returns true via ShowDialog when confirmed, or null when cancelled.
    // ###########################################################################################
    public partial class MarkOverMarksWindow : Window
    {
        public MarkOverMarksWindow()
        {
            this.InitializeComponent();

            // Tunnel, not a plain KeyDown - see DiscardDraftWindow.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        public void Initialize(string picture, int marked, int toMark)
        {
            this.Title = PictureToolWording.MarkOverMarksTitle;
            this.HeadingText.Text = PictureToolWording.MarkOverMarksTitle;
            this.QuestionText.Text = PictureToolWording.MarkOverMarksQuestion(picture, marked);
            this.DetailText.Text = PictureToolWording.MarkOverMarksDetail(toMark);
            this.MarkButton.Content = PictureToolWording.MarkOverMarksButton;
        }

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnMarkClick(object? sender, RoutedEventArgs e) => this.Close(true);
    }
}
