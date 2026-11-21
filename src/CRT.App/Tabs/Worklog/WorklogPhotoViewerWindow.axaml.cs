using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.IO;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Read-only viewer for one photo, opened by clicking a photo's thumbnail in the worklog editor,
    // in the device form or on a device's page - and any picture in "Capture from oscilloscope pin-by-pin",
    // where a click outside closes it (ClosesWhenLeft). Shows the image whole with its file name and comment -
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
            // accumulate a decoded copy per visit. Only the window's OWN decode: the small one shown
            // meanwhile is DevicePhotoThumbnails', which other screens show too (2026-11-16).
            this.Closed += (_, _) =>
            {
                this.thisIsClosed = true;
                this.thisWholePhoto?.Dispose();
                this.thisWholePhoto = null;
            };

            this.Deactivated += (_, _) =>
            {
                if (this.ClosesWhenLeft && !this.thisIsClosed)
                {
                    this.Close();
                }
            };
        }

        // ###########################################################################################
        // Closed by a click anywhere outside it - its window no longer the active one - as Escape and "Close" close it ("Capture from
        // oscilloscope pin-by-pin", owner request, 2026-11-20: "when I click with mouse outside it, then it should also close the
        // window"). Only shown with Show, never ShowDialog: a dialog refuses the click outside. Off elsewhere.
        // ###########################################################################################
        internal bool ClosesWhenLeft { get; set; }

        // The photo decoded whole, this window's own - null until it is in.
        private Bitmap? thisWholePhoto;

        internal Bitmap? WholePhotoForTests => this.thisWholePhoto;

        private bool thisIsClosed;

        // ###########################################################################################
        // Loads the photo to display. A missing or undecodable file leaves the frame empty with an
        // explanation rather than throwing - the bytes live beside entries.json and can be removed
        // or corrupted outside the app, and that should not take a window down.
        // ###########################################################################################
        public void Initialize(string fileName, string comment, string? imagePath)
        {
            this.Title = string.IsNullOrWhiteSpace(fileName) ? "Photo" : fileName;
            this.FileNameText.Text = fileName;
            this.CommentForTests = comment;

            // The comment is a note with formatting (RichNote): drawn as it reads, its links clickable - never its
            // marks as typed (code review, 2026-11-08: "**dead** SID" showed with its stars).
            TextLinkRenderer.SetLinkText(this.CommentText, RichNote.TrimMarkup(comment));
            this.CommentText.IsVisible = !string.IsNullOrWhiteSpace(comment);

            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
            {
                this.ShowUnavailable("This photo's file could not be found.");
                return;
            }

            this.ShowPhotoAsync(imagePath);
        }

        // ###########################################################################################
        // A picture already in memory - one no file holds, or holds no longer: an oscilloscope baseline replaced in the draft, or a
        // capture not yet accepted (BaselineCaptureWindow, 2026-11-20). The caller's: never disposed here.
        // ###########################################################################################
        public void Initialize(string title, string comment, Bitmap picture)
        {
            ArgumentNullException.ThrowIfNull(picture);

            this.Title = string.IsNullOrWhiteSpace(title) ? "Photo" : title;
            this.FileNameText.Text = title;
            this.CommentForTests = comment;
            TextLinkRenderer.SetLinkText(this.CommentText, RichNote.TrimMarkup(comment));
            this.CommentText.IsVisible = !string.IsNullOrWhiteSpace(comment);
            this.PhotoImage.Source = picture;
        }

        internal Bitmap? ShownForTests => this.PhotoImage.Source;

        internal string CommentForTests { get; private set; } = string.Empty;

        internal string? ShownTitleForTests => this.FileNameText.Text;

        // ###########################################################################################
        // *** THE WINDOW SHOWS AT ONCE (2026-11-16, performance walkthrough). *** It waited for the photo's whole
        // decode on the UI thread - 100-200 ms for a phone photo - before it appeared. Now the small read the photo
        // rows already hold (DevicePhotoThumbnails) shows at once when it is there, and the whole photo takes its
        // place once decoded on the pool. One decoded after the window closed is let go of at once.
        // ###########################################################################################
        private async void ShowPhotoAsync(string imagePath)
        {
            try
            {
                // Only a small read made already: one begun here would come after the whole photo (code review, 2026-11-16).
                bool smallMade = DevicePhotoThumbnails.TryGetReadMadeAlready(imagePath, out Task<Bitmap?> small);
                Task<Bitmap?> whole = Task.Run(() => WorklogPhotoViewerWindow.Decode(imagePath));
                bool previewShown = false;

                if (smallMade && small is { IsCompletedSuccessfully: true, Result: { } preview } && !whole.IsCompleted)
                {
                    this.PhotoImage.Source = preview;
                    previewShown = true;
                }

                Bitmap? decoded = await whole;

                if (this.thisIsClosed)
                {
                    decoded?.Dispose();
                    return;
                }

                if (decoded is null)
                {
                    this.ShowUnavailable("This photo could not be displayed.");
                    return;
                }

                this.thisWholePhoto = decoded;

                // In the preview's place, the zoom made on it kept.
                if (previewShown)
                    this.PhotoImage.ShowSharper(decoded);
                else
                    this.PhotoImage.Source = decoded;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open worklog photo [{imagePath}]: {ex.Message}");
            }
        }

        // Never throws - a photo that cannot be decoded is said on the window.
        private static Bitmap? Decode(string imagePath)
        {
            try
            {
                using var stream = File.OpenRead(imagePath);
                return new Bitmap(stream);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open worklog photo [{imagePath}]: {ex.Message}");
                return null;
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
