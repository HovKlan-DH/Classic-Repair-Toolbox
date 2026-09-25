using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Tiny modal collecting one photo attachment: an image file plus a comment. Returns the pair
    // via ShowDialog, or null when cancelled - the same shape as the Add comment/work done modals,
    // and it keeps the same keyboard contract (Escape cancels, Ctrl+Enter saves).
    //
    // The image can be chosen with Browse or dropped onto the box; both go through the same
    // ApplySelectedFile so a dropped file is vetted exactly like a picked one. The window only
    // resolves and validates a source path - copying the bytes into the entry's attachments folder
    // is the caller's job (WorklogEntryEditorWindow), because only the caller knows the entry.
    //
    // In edit mode SelectedFilePath is left null when the user does not pick a replacement, which
    // is how the caller tells "keep the existing image" from "swap it".
    // ###########################################################################################
    public partial class WorklogAddPhotoWindow : Window
    {
        // ###########################################################################################
        // What the dialog returns: the chosen image (null when an existing one is being kept) and
        // the comment, which may be empty - a photo is allowed to speak for itself.
        // ###########################################################################################
        public sealed record PhotoResult(string? SourcePath, string Comment);

        private string? thisSelectedFilePath;

        private bool thisIsEditMode;

        // ###########################################################################################
        // Which list the attachment is destined for. Photos are restricted to what the app can draw
        // and get a thumbnail preview; Files take the wider openable-document set and show no
        // preview, since there is nothing to render for a PDF. Everything else about the dialog -
        // the drop target, the comment box, the keyboard contract - is identical, which is why one
        // window serves both rather than a near-copy existing for each.
        // ###########################################################################################
        private WorklogAttachmentStorage.AttachmentKind thisKind = WorklogAttachmentStorage.AttachmentKind.Photo;

        private bool IsFileKind => this.thisKind == WorklogAttachmentStorage.AttachmentKind.File;

        public WorklogAddPhotoWindow()
        {
            this.InitializeComponent();

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.CommentTextBox.Focus(), DispatcherPriority.Background);

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            DragDrop.SetAllowDrop(this.ImageDropBorder, true);
            this.ImageDropBorder.AddHandler(DragDrop.DragOverEvent, this.OnImageDragOver);
            this.ImageDropBorder.AddHandler(DragDrop.DropEvent, this.OnImageDrop);

            // The preview bitmap holds an unmanaged surface and would otherwise outlive the dialog.
            this.Closed += (_, _) => (this.ImagePreview.Source as IDisposable)?.Dispose();
        }

        // ###########################################################################################
        // Switches the dialog to collecting a File rather than a Photo: wider accepted types, no
        // image preview, and wording that says "file" throughout. Call before showing.
        //
        // InitializeForEdit takes the kind itself and assigns it unconditionally, rather than
        // relying on this having been called first - the two used to be order-dependent, and
        // calling them the other way round produced a dialog worded for photos that nonetheless
        // accepted PDFs.
        // ###########################################################################################
        public void InitializeForFileKind()
        {
            this.thisKind = WorklogAttachmentStorage.AttachmentKind.File;

            this.Title = "Add file";
            this.HeaderText.Text = "Add file";
            this.AddButton.Content = "Add file";
            this.SectionLabelText.Text = "File";
            this.AcceptedTypesText.Text = "Documents, images and data files - not programs, scripts or shortcuts";
            this.SelectedFileText.Text = "Drag a file here, or use Browse";
        }

        // ###########################################################################################
        // Switches the dialog into "edit" mode: shows the attachment already there and its comment,
        // and relabels the title/submit button, so the same modal serves both Add and the row's
        // click-to-edit behavior - matching the comment and work-done dialogs.
        //
        // existingImagePath is only for the preview (Photos only); leaving the picker untouched
        // returns a null SourcePath, meaning the stored file stays as it is.
        // ###########################################################################################
        public void InitializeForEdit(
            string fileName,
            string comment,
            string? existingImagePath,
            WorklogAttachmentStorage.AttachmentKind kind = WorklogAttachmentStorage.AttachmentKind.Photo)
        {
            this.thisIsEditMode = true;

            // The kind is set from the ARGUMENT in both directions, never inferred from whatever a
            // previous call left behind. Applying it only for File made the claim below true one
            // way round and false the other: an instance switched to File and then edited as a
            // Photo kept File's wider validation while being worded for photos, accepting a PDF
            // where the caller asked for an image. Assigning unconditionally is what actually makes
            // this method own the whole appearance regardless of what the caller called before.
            this.thisKind = kind;

            if (kind == WorklogAttachmentStorage.AttachmentKind.File)
            {
                this.InitializeForFileKind();
            }

            bool isFile = this.IsFileKind;

            this.Title = isFile ? "Edit file" : "Edit photo";
            this.HeaderText.Text = this.Title;
            this.AddButton.Content = isFile ? "Update file" : "Update photo";
            this.CommentTextBox.Text = comment;

            this.SelectedFileText.Text = fileName;
            this.SelectedFilePathText.Text = isFile
                ? "Drop or browse to replace this file"
                : "Drop or browse to replace this image";
            this.SelectedFilePathText.IsVisible = true;

            this.TryShowPreview(existingImagePath);
        }

        // ###########################################################################################
        // Escape cancels and Ctrl+Enter saves, identical to the comment and work-done dialogs. Plain
        // Enter is deliberately left alone because the comment field is multi-line. Handled on the
        // Tunnel route so it fires before CommentTextBox's own AcceptsReturn handling inserts a
        // newline - a bubbling handler would run too late.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter && e.KeyModifiers.HasFlag(KeyModifiers.Control))
            {
                this.OnAddClick(sender, e);
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Only offers the copy cursor for a drag actually carrying files - without this the box
        // appears to accept dragged text or anything else and then silently does nothing on drop.
        // ###########################################################################################
        private void OnImageDragOver(object? sender, DragEventArgs e)
        {
            e.DragEffects = e.DataTransfer.Contains(DataFormat.File) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        // ###########################################################################################
        // Takes the first dropped file. Multi-file drops use only the first rather than being
        // refused: the dialog collects exactly one photo, and quietly taking one of several is
        // friendlier than rejecting the whole drop.
        // ###########################################################################################
        private void OnImageDrop(object? sender, DragEventArgs e)
        {
            e.Handled = true;

            var files = e.DataTransfer.TryGetFiles();
            if (files == null)
            {
                return;
            }

            var first = files.FirstOrDefault();
            if (first == null)
            {
                return;
            }

            string? path = first.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(path))
            {
                // A drag from a source with no local file (a virtual item from an archive or a
                // remote location) cannot be copied, so it is refused rather than half-accepted.
                this.ShowValidationMessage("That item is not a file on this computer.");
                return;
            }

            this.ApplySelectedFile(path);
        }

        private async void OnBrowseClick(object? sender, RoutedEventArgs e)
        {
            await this.PickImageAsync();
        }

        // ###########################################################################################
        // Offers only the formats the application can draw, the same filter the contribution editor
        // uses. The filter is a suggestion the user can type past, so ApplySelectedFile re-checks.
        // ###########################################################################################
        private async Task PickImageAsync()
        {
            var storageProvider = this.StorageProvider;
            if (storageProvider == null)
            {
                return;
            }

            // The Files filter is built from the launcher's own openable set, so the picker can only
            // offer what the app will later be able to open. ApplySelectedFile re-checks regardless,
            // since a filter is only a suggestion a typed name can get past.
            var fileType = this.IsFileKind
                ? new FilePickerFileType("Documents, images and data")
                {
                    Patterns = ExternalTargetLauncher.OpenableFileExtensions.Select(extension => "*" + extension).ToArray()
                }
                : new FilePickerFileType("Image files")
                {
                    Patterns = ContributionPackaging.DisplayableImageExtensions.Select(extension => "*" + extension).ToArray(),
                    MimeTypes = new[] { "image/*" },
                    AppleUniformTypeIdentifiers = new[] { "public.image" }
                };

            var files = await storageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = this.IsFileKind ? "Select file" : "Select image",
                AllowMultiple = false,
                FileTypeFilter = new[] { fileType }
            });

            if (files == null || files.Count == 0)
            {
                return;
            }

            string? path = files[0].TryGetLocalPath();
            if (!string.IsNullOrWhiteSpace(path))
            {
                this.ApplySelectedFile(path);
            }
        }

        // ###########################################################################################
        // The single gate every chosen file passes through, whether picked or dropped: vet it, and
        // only show it as selected once it is known to be usable. A refused file leaves any previous
        // selection intact rather than clearing it.
        // ###########################################################################################
        private void ApplySelectedFile(string path)
        {
            var problem = WorklogAttachmentStorage.ValidateSourceFile(path, this.thisKind);
            if (problem != WorklogAttachmentStorage.AttachmentProblem.None)
            {
                this.ShowValidationMessage(WorklogAttachmentStorage.DescribeProblem(problem));
                return;
            }

            this.thisSelectedFilePath = path;

            this.ImageValidationText.IsVisible = false;
            this.SelectedFileText.Text = Path.GetFileName(path);
            this.SelectedFilePathText.Text = path;
            this.SelectedFilePathText.IsVisible = true;
            this.ClearButton.IsVisible = true;

            this.TryShowPreview(path);
        }

        // ###########################################################################################
        // Drops the pending selection. In edit mode this reverts to keeping the stored image rather
        // than removing it - a photo row without an image would have nothing to show.
        // ###########################################################################################
        private void OnClearClick(object? sender, RoutedEventArgs e)
        {
            this.thisSelectedFilePath = null;

            this.ImageValidationText.IsVisible = false;
            this.ClearButton.IsVisible = false;

            var previous = this.ImagePreview.Source as IDisposable;
            this.ImagePreview.Source = null;
            this.ImagePreview.IsVisible = false;
            previous?.Dispose();

            this.SelectedFileText.Text = (this.thisIsEditMode, this.IsFileKind) switch
            {
                (true, true) => "Keeping the current file",
                (true, false) => "Keeping the current image",
                (false, true) => "Drag a file here, or use Browse",
                (false, false) => "Drag an image here, or use Browse"
            };
            this.SelectedFilePathText.IsVisible = false;
        }

        // ###########################################################################################
        // Shows a thumbnail of the chosen image. Decoding is best-effort: a file with a valid
        // extension can still be corrupt or truncated, and that must not take the dialog down, so a
        // failure just leaves the preview empty while the file itself stays selected.
        // ###########################################################################################
        private void TryShowPreview(string? path)
        {
            // Files show no preview - a PDF or CSV has nothing to render, and an image attached as a
            // File is still being treated as a document. Falls through to the clearing branch rather
            // than returning outright, so switching to File kind after a preview was shown still
            // releases that bitmap instead of leaving it on screen and undisposed.
            if (this.IsFileKind)
            {
                path = null;
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                var discarded = this.ImagePreview.Source as IDisposable;
                this.ImagePreview.Source = null;
                this.ImagePreview.IsVisible = false;
                discarded?.Dispose();
                return;
            }

            // Decoded down to the preview's own size rather than at full resolution: a phone photo
            // is several thousand pixels wide and this box is 72, so new Bitmap(stream) spent tens
            // of megabytes of unmanaged memory to draw a thumbnail. Same pattern the editor's row
            // thumbnails use.
            var previous = this.ImagePreview.Source as IDisposable;

            try
            {
                using var stream = File.OpenRead(path);
                this.ImagePreview.Source = Bitmap.DecodeToWidth(stream, 144);
                this.ImagePreview.IsVisible = true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to preview worklog photo [{path}]: {ex.Message}");
                this.ImagePreview.Source = null;
                this.ImagePreview.IsVisible = false;
            }

            // Disposed only after the control has been repointed, so nothing renders against a
            // freed surface. Browsing through several photos would otherwise orphan each one.
            previous?.Dispose();
        }

        private void ShowValidationMessage(string message)
        {
            this.ImageValidationText.Text = message;
            this.ImageValidationText.IsVisible = true;
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close(null);
        }

        // ###########################################################################################
        // An image is required when adding. When editing, keeping the existing one is fine, so only
        // a comment change is needed - the caller reads a null SourcePath as "leave the image".
        // ###########################################################################################
        private void OnAddClick(object? sender, RoutedEventArgs e)
        {
            if (!this.thisIsEditMode && string.IsNullOrWhiteSpace(this.thisSelectedFilePath))
            {
                this.ShowValidationMessage(
                    WorklogAttachmentStorage.DescribeProblem(WorklogAttachmentStorage.AttachmentProblem.NoFileSelected));
                return;
            }

            // Re-vetted at submit: the file was verified when it was chosen, but the dialog can sit
            // open for a while and the file can be moved or deleted in the meantime.
            if (!string.IsNullOrWhiteSpace(this.thisSelectedFilePath))
            {
                var problem = WorklogAttachmentStorage.ValidateSourceFile(this.thisSelectedFilePath, this.thisKind);
                if (problem != WorklogAttachmentStorage.AttachmentProblem.None)
                {
                    this.ShowValidationMessage(WorklogAttachmentStorage.DescribeProblem(problem));
                    return;
                }
            }

            this.Close(new PhotoResult(this.thisSelectedFilePath, this.CommentTextBox.Text?.Trim() ?? string.Empty));
        }
    }
}
