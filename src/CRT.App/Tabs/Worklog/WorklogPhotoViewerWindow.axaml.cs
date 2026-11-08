using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.IO;

namespace CRT
{
    // ###########################################################################################
    // Read-only viewer for one photo, opened by clicking a photo's thumbnail in the worklog editor,
    // in the device form or on a device's page. Shows the image whole with its file name and comment -
    // the wheel zooms in at the pointer (owner request, 2026-11-06; PhotoZoomView); editing and
    // deleting stay with the rows' buttons, so this window has nothing to save.
    // ###########################################################################################
    public partial class WorklogPhotoViewerWindow : Window
    {
        public WorklogPhotoViewerWindow()
        {
            this.InitializeComponent();

            this.ZoomHintText.Text = PanGesture.Hint + " Double-click to see it whole.";

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

            // The comment is a note with formatting (RichNote): drawn as it reads, its links clickable - never its
            // marks as typed (code review, 2026-11-08: "**dead** SID" showed with its stars).
            TextLinkRenderer.SetLinkText(this.CommentText, RichNote.TrimMarkup(comment));
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
