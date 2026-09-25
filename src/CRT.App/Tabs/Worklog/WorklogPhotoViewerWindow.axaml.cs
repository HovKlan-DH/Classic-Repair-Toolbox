using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using System;
using System.IO;

namespace CRT
{
    // ###########################################################################################
    // Read-only viewer for one worklog photo, opened by clicking a photo row's thumbnail in the
    // entry editor. Shows the image at full size with its file name and comment; editing and
    // deleting stay in the editor's row buttons, so this window has nothing to save.
    // ###########################################################################################
    public partial class WorklogPhotoViewerWindow : Window
    {
        public WorklogPhotoViewerWindow()
        {
            this.InitializeComponent();

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            // Full-resolution decode is the point of this window, which makes the bitmap large -
            // tens of megabytes for a phone photo - so it is released when the window closes rather
            // than left waiting on a finalizer. Opening the viewer repeatedly would otherwise
            // accumulate a decoded copy per visit.
            this.Closed += (_, _) => (this.PhotoImage.Source as IDisposable)?.Dispose();
        }

        // ###########################################################################################
        // Loads the photo to display. A missing or undecodable file leaves the frame empty with an
        // explanation rather than throwing - the bytes live beside entries.json and can be removed
        // or corrupted outside the app, and that should not take a window down.
        // ###########################################################################################
        public void Initialize(string fileName, string comment, string? imagePath)
        {
            this.Title = string.IsNullOrWhiteSpace(fileName) ? "Photo" : fileName;
            this.FileNameText.Text = fileName;

            this.CommentText.Text = comment;
            this.CommentText.IsVisible = !string.IsNullOrWhiteSpace(comment);

            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
            {
                this.ShowUnavailable("This photo's file could not be found.");
                return;
            }

            try
            {
                using var stream = File.OpenRead(imagePath);
                this.PhotoImage.Source = new Bitmap(stream);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open worklog photo [{imagePath}]: {ex.Message}");
                this.ShowUnavailable("This photo could not be displayed.");
            }
        }

        private void ShowUnavailable(string message)
        {
            this.PhotoImage.Source = null;
            this.PhotoUnavailableText.Text = message;
            this.PhotoUnavailableText.IsVisible = true;
        }

        // ###########################################################################################
        // Escape closes, matching every other worklog modal. There is nothing to submit here, so
        // Ctrl+Enter is deliberately not handled.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
