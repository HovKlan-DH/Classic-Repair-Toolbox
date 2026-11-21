using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Oscilloscope;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // BaselineCaptureWindow's PICTURES (see BaselineCaptureWindow.axaml.cs for the file map): the pin shown's pictures as tall as
    // the window allows - its baseline alone, or the old baseline framed beside the new capture - read off the UI thread, and every
    // picture clicked to see it large, "Click to see it larger" following the pointer.
    // ###########################################################################################
    public partial class BaselineCaptureWindow
    {
        // ---- the pictures -----------------------------------------------------------------------

        // Big pictures of the baselines, by their file - read once while the window is open - and the reads under way.
        private readonly Dictionary<string, Bitmap?> thisImages = new(StringComparer.Ordinal);
        private readonly Queue<string> thisImageOrder = new();
        private readonly Dictionary<string, Task> thisImagesBeingRead = new(StringComparer.Ordinal);
        private const int ImagesKept = 24;

        // ###########################################################################################
        // The pictures of the pin shown, as tall as the window allows: its baseline alone - or, captured again, the old baseline framed
        // beside the new capture, a strip over each saying which it is (owner requests, 2026-11-20: "have it more clear/visible what is
        // the old and new image", "a red'ish background color") - the old red, a capture amber while not yet accepted and green once
        // accepted. A capture not yet accepted is its own picture, read from where it waits.
        // ###########################################################################################
        private void ShowPictures(BaselineCaptureTarget target)
        {
            BaselineRunStep step = target.Step;
            ComponentPaneBaselineRow? row = target.Row;
            bool waiting = target.Waiting is not null;
            bool compare = target.IsComparable && row is not null;
            ComponentImageEntry? shown = target.WaitingEntry(row, this.MachineStateBox.Text ?? string.Empty, this.CapturedBaselineRegion) ?? row?.Committed;
            string settings = shown is null ? string.Empty : ComponentPaneWording.RunSettings(shown);

            this.SingleFrame.IsVisible = !compare;
            this.ComparePanel.IsVisible = compare;

            if (!compare)
            {
                this.NoImageText.IsVisible = shown is null;
                this.SingleHeader.IsVisible = shown is not null;
                BoldRuns.ShowMarked(this.SingleHeaderText, ComponentPaneWording.RunPictureHeader(
                    step.IsCaptured ? ComponentPaneWording.PictureNewTitle(step.IsAccepted) : ComponentPaneWording.PictureBaseline,
                    settings));
                BaselineCaptureWindow.ShowCaptureState(this.SingleFrame, this.SingleHeader, step);

                if (waiting)
                {
                    BaselineCaptureWindow.ShowWaitingPicture(this.CurrentImage, target);
                }
                else
                {
                    this.ShowImage(this.CurrentImage, row);
                }

                this.ShowImage(this.NowImage, null);
                this.OldImage.Source = null;
                return;
            }

            this.ShowImage(this.CurrentImage, null);

            if (waiting)
            {
                BaselineCaptureWindow.ShowWaitingPicture(this.NowImage, target);
            }
            else
            {
                this.ShowImage(this.NowImage, row);
            }

            BoldRuns.ShowMarked(this.OldHeaderText, ComponentPaneWording.RunPictureHeader(
                ComponentPaneWording.PictureOld,
                target.BeforeEntry is null ? string.Empty : ComponentPaneWording.RunSettings(target.BeforeEntry)));
            BoldRuns.ShowMarked(this.NewHeaderText, ComponentPaneWording.RunPictureHeader(ComponentPaneWording.PictureNewTitle(step.IsAccepted), settings));
            BaselineCaptureWindow.ShowCaptureState(this.NewFrame, this.NewHeader, step);

            this.ShowOldPicture(target);
        }

        // A capture's frame and strip: amber while it is not yet accepted, green once accepted - a baseline not captured again quiet.
        private static void ShowCaptureState(Border frame, Border header, BaselineRunStep step)
        {
            frame.Classes.Set("ToCheck", step.IsWaiting);
            frame.Classes.Set("Accepted", step.IsAccepted);
            header.Classes.Set("ToCheck", step.IsWaiting);
            header.Classes.Set("Accepted", step.IsAccepted);
        }

        // A capture not yet accepted, as read from where it waits - nothing until it is read.
        private static void ShowWaitingPicture(Image image, BaselineCaptureTarget target)
        {
            image.Tag = string.Empty;
            image.Source = target.WaitingPicture;
        }

        // ###########################################################################################
        // The old baseline's picture - read off the UI thread from the bytes kept before a capture could replace it, once - or why
        // there is none.
        // ###########################################################################################
        private void ShowOldPicture(BaselineCaptureTarget target)
        {
            if (target.BeforeBytes is not { } bytes)
            {
                this.OldImage.Source = null;
                this.NoOldImageText.IsVisible = true;
                return;
            }

            if (target.OldPictureRead is null)
            {
                target.OldPictureRead = Task.Run(() => BaselineCaptureWindow.DecodeBytes(bytes), CancellationToken.None);
                this.thisPictureReads.Add(this.ShowOnceReadAsync(target, target.OldPictureRead));
            }

            Bitmap? picture = target.OldPicture;

            this.OldImage.Source = picture;
            this.NoOldImageText.IsVisible = target.OldPictureRead.IsCompleted && picture is null;
        }

        // A picture of `target` read - the old baseline's, or a capture's - shown if it is the pin shown still, and in its line.
        private async Task ShowOnceReadAsync(BaselineCaptureTarget target, Task<Bitmap?> read)
        {
            try
            {
                await read;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read a picture of pin [{target.Step.PinText}] - [{ex.Message}]");
            }

            target.ShowPictures();

            if (this.CurrentForTests is { } shown && ReferenceEquals(shown, target))
            {
                this.ShowPictures(target);
            }
        }

        private void ShowCase()
        {
            static List<int> Numbers(IEnumerable<BaselineRunStep> steps) =>
                steps.Where(step => step.Pin.Number is not null).Select(step => step.Pin.Number!.Value).Distinct().ToList();

            this.CasePinout.ShowChosen(Numbers(this.Steps().Where(step => step.IsPicked)), Numbers(this.Steps().Where(step => step.IsAccepted)));
            this.CasePinout.MarkPin(this.thisCurrentIndex >= 0 ? this.thisTargets[this.thisCurrentIndex].Step.Pin.Number : null);
        }

        private bool IsNextTarget(BaselineCaptureTarget target) =>
            this.NextIndex() is int next and >= 0 && ReferenceEquals(this.thisTargets[next], target);

        private int NextIndex() => this.thisCurrentIndex < 0 ? -1 : BaselineRunPins.NextToAccept(this.Steps(), this.thisCurrentIndex);

        private void ShowNext()
        {
            int next = this.NextIndex();

            // With nothing picked there is no next pin to name - the line over the list says what to do.
            this.NextBorder.IsVisible = this.thisTargets.Any(target => target.Step.IsPicked);

            BaselineRunStep? shown = this.thisCurrentIndex >= 0 && this.thisCurrentIndex < this.thisTargets.Count
                ? this.thisTargets[this.thisCurrentIndex].Step
                : null;

            BoldRuns.ShowMarked(this.NextText, next >= 0
                ? ComponentPaneWording.RunNext(this.thisTargets[next].Step)
                : shown is { IsPicked: true, IsAccepted: false }
                    ? ComponentPaneWording.RunShownIsLast
                    : ComponentPaneWording.RunNext(null));

            foreach (BaselineCaptureTarget target in this.thisTargets)
            {
                target.ShowState(next >= 0 && ReferenceEquals(this.thisTargets[next], target));
            }
        }

        private void ShowProgress()
        {
            List<BaselineRunStep> picked = this.Steps().Where(step => step.IsPicked).ToList();
            BoldRuns.ShowMarked(this.ProgressText, ComponentPaneWording.RunProgress(picked.Count, picked.Count(step => step.IsAccepted)));
        }

        private void ShowButtons()
        {
            bool any = this.thisTargets.Count > 0;

            this.CaptureButton.IsEnabled = this.thisConnected && any;
            this.PreviousButton.IsEnabled = any;
            this.AcceptButton.IsEnabled = any;
        }

        // A new baseline's signal name as typed - kept with its line.
        private void CommitName()
        {
            if (this.thisCurrentIndex >= 0 && this.thisCurrentIndex < this.thisTargets.Count && this.NamePanel.IsVisible)
            {
                this.thisTargets[this.thisCurrentIndex].Step.Name = this.NameBox.Text?.Trim() ?? string.Empty;
                this.thisTargets[this.thisCurrentIndex].ShowLabel();
            }
        }

        // ###########################################################################################
        // A baseline's picture, large - read off the UI thread from where the draft keeps it, at most once while the window is
        // open; its thumbnail stands in until then. Never disposed by hand: one still drawn would take the render thread down.
        // ###########################################################################################
        private void ShowImage(Image image, ComponentPaneBaselineRow? row)
        {
            string file = row?.Committed.File?.Trim() ?? string.Empty;

            image.Tag = file;

            if (row is null || file.Length == 0)
            {
                image.Source = null;
                return;
            }

            if (this.thisImages.TryGetValue(file, out Bitmap? kept))
            {
                image.Source = kept ?? row.Thumbnail;
                return;
            }

            image.Source = row.Thumbnail;

            // One read per file at a time: two images may show the same one.
            if (!this.thisImagesBeingRead.TryGetValue(file, out Task? reading) || reading.IsCompleted)
            {
                this.thisImagesBeingRead[file] = this.ReadImageAsync(file);
            }
        }

        // `file`'s big picture read, kept, and shown wherever it is still wanted.
        private async Task ReadImageAsync(string file)
        {
            try
            {
                string? path = this.thisPane?.ResolveFileOfBoard(file);
                Bitmap? bitmap = await Task.Run(() => BaselineCaptureWindow.Decode(path), CancellationToken.None);

                if (!this.thisImages.ContainsKey(file))
                {
                    this.thisImages[file] = bitmap;
                    this.thisImageOrder.Enqueue(file);

                    while (this.thisImageOrder.Count > ImagesKept)
                    {
                        this.thisImages.Remove(this.thisImageOrder.Dequeue());
                    }
                }

                if (bitmap is not null)
                {
                    foreach (Image image in new[] { this.CurrentImage, this.NowImage }.Where(image => string.Equals(image.Tag as string, file, StringComparison.Ordinal)))
                    {
                        image.Source = bitmap;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not show the oscilloscope baseline [{file}] - [{ex.Message}]");
            }
        }

        private static Bitmap? Decode(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path) || !ContributionPackaging.IsDisplayableImageFile(path))
            {
                return null;
            }

            try
            {
                using FileStream stream = File.OpenRead(path);

                // An oscilloscope's screen is small; a photo given as a baseline is read no wider than the window can show it.
                return stream.Length > 4_000_000 ? Bitmap.DecodeToWidth(stream, 1600, BitmapInterpolationMode.MediumQuality) : new Bitmap(stream);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read the picture [{path}] - [{ex.Message}]");
                return null;
            }
        }

        // As Decode reads a file: a large picture no wider than the window can show it (code review, 2026-11-21: a 20 MB photo given as
        // a baseline was decoded whole, well over 100 MB, for as long as the window stayed open).
        private static Bitmap? DecodeBytes(byte[] bytes)
        {
            try
            {
                using var stream = new MemoryStream(bytes);
                return bytes.Length > 4_000_000 ? Bitmap.DecodeToWidth(stream, 1600, BitmapInterpolationMode.MediumQuality) : new Bitmap(stream);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read the old baseline's picture - [{ex.Message}]");
                return null;
            }
        }

        // The old baseline's picture, read as the run started from it - before a capture accepted replaces it in the draft, whose own
        // file is then listed no more. Null when it is nowhere on this machine.
        private async Task<byte[]?> ReadOldPictureBytesAsync(string file)
        {
            string? path = this.thisPane?.ResolveFileOfBoard(file);

            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return null;
            }

            try
            {
                return await File.ReadAllBytesAsync(path, CancellationToken.None);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not read the old baseline [{path}] - [{ex.Message}]");
                return null;
            }
        }

        // ---- every picture, large ---------------------------------------------------------------

        private void OnThumbnailPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Control { DataContext: BaselineCaptureTarget target } control || !e.GetCurrentPoint(control).Properties.IsLeftButtonPressed)
            {
                return;
            }

            // The picture's press alone - the line is not picked by it, nor the oscilloscope set.
            e.Handled = true;
            this.OpenPicture(target);
        }

        private void OnShownPicturePressed(object? sender, PointerPressedEventArgs e)
        {
            if (BaselineCaptureWindow.IsLeftPress(sender, e) && this.CurrentForTests is { } target)
            {
                this.OpenPicture(target);
            }
        }

        private void OnOldPicturePressed(object? sender, PointerPressedEventArgs e)
        {
            if (BaselineCaptureWindow.IsLeftPress(sender, e))
            {
                this.OpenOldPicture();
            }
        }

        private static bool IsLeftPress(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Visual visual || !e.GetCurrentPoint(visual).Properties.IsLeftButtonPressed)
            {
                return false;
            }

            e.Handled = true;
            return true;
        }

        // ###########################################################################################
        // "Click to see it larger" beside the pointer while it is over a picture there is (owner request, 2026-11-20: "have the
        // hover/helper text follow the mouse") - drawn in the window and taking no click, never a tooltip (PointerHintGeometry).
        // ###########################################################################################
        private void OnPicturePointerMoved(object? sender, PointerEventArgs e)
        {
            if (sender is not Panel picture || !picture.Children.OfType<Image>().Any(image => image.Source is not null && image.IsVisible))
            {
                this.PictureHint.IsVisible = false;
                return;
            }

            this.PictureHint.IsVisible = true;
            this.PictureHint.Measure(Size.Infinity);

            Point place = PointerHintGeometry.Place(e.GetPosition(this.HintLayer), this.PictureHint.DesiredSize, this.HintLayer.Bounds.Size);

            Canvas.SetLeft(this.PictureHint, place.X);
            Canvas.SetTop(this.PictureHint, place.Y);
        }

        private void OnPicturePointerExited(object? sender, PointerEventArgs e) => this.PictureHint.IsVisible = false;

        // ###########################################################################################
        // `target`'s picture large - a capture not yet accepted, else its baseline as it now is; nothing for one with no picture. A
        // capture waiting is ALWAYS its own picture: not read yet, it is read from where it waits - never the old baseline under the
        // new capture's title (code review, 2026-11-21).
        // ###########################################################################################
        internal void OpenPicture(BaselineCaptureTarget target)
        {
            ArgumentNullException.ThrowIfNull(target);

            ComponentPaneBaselineRow? row = target.Row;
            string title = ComponentPaneWording.RunPictureTitle(target.Step, target.Step.IsCaptured ? ComponentPaneWording.PictureNew : null);

            if (target.Waiting is { } capture && target.WaitingEntry(row, this.MachineStateBox.Text ?? string.Empty, this.CapturedBaselineRegion) is { } entry)
            {
                var viewer = new WorklogPhotoViewerWindow();
                string settings = ComponentPaneWording.RunSettingsNamed(entry);

                if (target.WaitingPicture is { } waiting)
                {
                    viewer.Initialize(title, settings, waiting);
                }
                else
                {
                    viewer.Initialize(title, settings, capture.Path);
                }

                this.ShowPicture(viewer);
                return;
            }

            if (row?.Committed.File?.Trim() is not { Length: > 0 } file)
            {
                return;
            }

            var shown = new WorklogPhotoViewerWindow();
            shown.Initialize(title, ComponentPaneWording.RunSettingsNamed(row.Committed), this.thisPane?.ResolveFileOfBoard(file));
            this.ShowPicture(shown);
        }

        // The old baseline of the pin shown, large.
        internal void OpenOldPicture()
        {
            if (this.CurrentForTests is { OldPicture: { } old } target)
            {
                var viewer = new WorklogPhotoViewerWindow();
                viewer.Initialize(
                    ComponentPaneWording.RunPictureTitle(target.Step, ComponentPaneWording.PictureOld),
                    target.BeforeEntry is null ? string.Empty : ComponentPaneWording.RunSettingsNamed(target.BeforeEntry),
                    old);
                this.ShowPicture(viewer);
            }
        }

        // ###########################################################################################
        // A picture large, over this window - and gone at a click anywhere outside it, as Escape and its "Close" do (owner request,
        // 2026-11-20: "when I click with mouse outside it, then it should also close the window"): shown beside this window, never
        // as a dialog, which would refuse that click.
        // ###########################################################################################
        private void ShowPicture(WorklogPhotoViewerWindow viewer)
        {
            this.PictureHint.IsVisible = false;
            viewer.ClosesWhenLeft = true;

            if (this.ShowPictureOverrideForTests is { } show)
            {
                show(viewer);
                return;
            }

            viewer.Show(this);
        }
    }
}
