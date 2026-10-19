// FeedbackComposer, the input half: the keyboard across the parts (Up and Down over a picture,
// Backspace and Delete taking one away, Ctrl+A then Delete clearing the field), a click on the
// empty space, paste (files, the text the TextBox would have pasted itself, or a picture - in that
// order) and drag and drop. The parts and pictures themselves are in FeedbackComposer.axaml.cs.
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Platform;
using Handlers.Geometry;
using System.Runtime.InteropServices;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    public partial class FeedbackComposer
    {
        // The caret as it was before the TextBox saw the key: Up or Down that left it where it was
        // was on the part's first or last line, and goes to the part above or below.
        private int thisCaretBeforeKey;

        // Ctrl+A with more than one part selects the text of every part; the next Backspace or
        // Delete then clears the whole field. Any other key, a click, the focus leaving a part, a
        // picture put in or taken out, or Clear is the end of it - and it holds only while every
        // part's text is still selected whole (code review, 2026-10-16: left set after the focus
        // went elsewhere and came back by Tab, a Backspace meant for one character erased it all).
        private bool thisAllSelected;

        // Set while the composer calls a part's own Paste() for plain text, so that the
        // PastingFromClipboard it raises is let through to the TextBox instead of back here.
        private bool thisPassingPasteThrough;

        // Run on the pool thread just before a pasted picture's PNG is encoded - a test holds the
        // encoding there to act while it is under way.
        internal Action? BeforeEncodingForTests { get; set; }

        // ###########################################################################################
        // *** PASTES AND DROPS STILL PUTTING THINGS IN (code review, 2026-10-16). *** A dropped
        // photo is decoded before it is shown, one after another, and files that are no pictures
        // reach the attachments list when the last is done - so Send, clicked meanwhile, read the
        // field without them, and its Clear after a successful send threw them away unsent. Each
        // is kept here until done (UI thread only); WhenSettledAsync is what Send waits for.
        // ###########################################################################################
        private readonly HashSet<Task> thisUnderWay = new();

        private void InitializeInput()
        {
            this.AddHandler(KeyDownEvent, this.OnKeyDownBeforePart, RoutingStrategies.Tunnel);
            this.AddHandler(KeyDownEvent, this.OnKeyDownAfterPart, RoutingStrategies.Bubble, handledEventsToo: true);
            this.AddHandler(PointerPressedEvent, (_, _) => this.thisAllSelected = false, RoutingStrategies.Tunnel);
            this.Frame.PointerPressed += this.OnFramePressed;

            // The composer's border shows the focus of whichever part has it; a part losing it ends
            // a Ctrl+A.
            this.AddHandler(GotFocusEvent, (_, _) => this.Frame.Classes.Set("Focused", true));
            this.AddHandler(LostFocusEvent, (_, _) =>
            {
                this.thisAllSelected = false;
                this.Frame.Classes.Set("Focused", this.thisParts.Any(part => part.IsFocused));
            });

            DragDrop.SetAllowDrop(this, true);
            this.AddHandler(DragDrop.DragOverEvent, this.OnDragOver);
            this.AddHandler(DragDrop.DropEvent, this.OnDrop);
        }

        private TextBox? FocusedPart() => this.thisParts.FirstOrDefault(part => part.IsFocused);

        // Every part's text still selected whole, as Ctrl+A left it.
        private bool EveryPartSelectedWhole() =>
            this.thisParts.All(part =>
                Math.Min(part.SelectionStart, part.SelectionEnd) == 0 &&
                Math.Max(part.SelectionStart, part.SelectionEnd) == (part.Text?.Length ?? 0));

        // ###########################################################################################
        // Done once every paste and drop under way has put in what it brings - and any that began
        // while waiting. A failure among them was logged where it happened; it does not fail this.
        // ###########################################################################################
        public async Task WhenSettledAsync()
        {
            while (true)
            {
                this.thisUnderWay.RemoveWhere(work => work.IsCompleted);

                if (this.thisUnderWay.Count == 0)
                    return;

                try
                {
                    await Task.WhenAll(this.thisUnderWay.ToArray());
                }
                catch (Exception)
                {
                    // Logged by KeepUnderWay.
                }
            }
        }

        // A paste or a drop kept until done (WhenSettledAsync), its failure logged.
        private void KeepUnderWay(Task work, string what)
        {
            this.thisUnderWay.RemoveWhere(done => done.IsCompleted);

            if (work.IsCompleted)
            {
                if (work.IsFaulted)
                    Logger.Warning($"{what} failed - [{work.Exception}]");

                return;
            }

            this.thisUnderWay.Add(work);
            work.ContinueWith(failed => Logger.Warning($"{what} failed - [{failed.Exception}]"), TaskContinuationOptions.OnlyOnFaulted);
        }

        // ###########################################################################################
        // Before the part's own TextBox sees the key. Backspace at the start of the part under a
        // picture, and Delete at the end of the part above one, take the picture away - the TextBox
        // would do nothing with either there. Ctrl+A (the platform's select-all gesture) with more
        // than one part selects every part.
        // ###########################################################################################
        private void OnKeyDownBeforePart(object? sender, KeyEventArgs e)
        {
            TextBox? part = this.FocusedPart();

            if (part is null)
                return;

            int index = this.thisParts.IndexOf(part);
            int length = part.Text?.Length ?? 0;
            bool nothingSelected = part.SelectionStart == part.SelectionEnd;
            this.thisCaretBeforeKey = part.CaretIndex;

            if (this.thisAllSelected)
            {
                if ((e.Key is Key.Back or Key.Delete) && this.EveryPartSelectedWhole())
                {
                    this.Clear();
                    e.Handled = true;
                    return;
                }

                // A modifier on its own keeps the selection - Ctrl is still down after Ctrl+A.
                this.thisAllSelected = e.Key is Key.LeftCtrl or Key.RightCtrl or Key.LeftShift or Key.RightShift
                    or Key.LeftAlt or Key.RightAlt or Key.LWin or Key.RWin;
            }

            if (this.thisParts.Count > 1 && FeedbackComposer.IsSelectAll(e))
            {
                foreach (TextBox each in this.thisParts)
                    each.SelectAll();

                this.thisAllSelected = true;
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Back && nothingSelected && part.CaretIndex == 0 && index > 0)
            {
                this.RemovePictureAt(index - 1);
                e.Handled = true;
            }
            else if (e.Key == Key.Delete && nothingSelected && part.CaretIndex == length && index < this.thisPictures.Count)
            {
                this.RemovePictureAt(index);
                e.Handled = true;
            }
        }

        // The platform's own select-all gesture (Ctrl+A, Cmd+A on a Mac), as the TextBox reads it.
        private static bool IsSelectAll(KeyEventArgs e)
        {
            PlatformHotkeyConfiguration? hotkeys = Application.Current?.PlatformSettings?.HotkeyConfiguration;

            if (hotkeys is null)
                return e.Key == Key.A && e.KeyModifiers == KeyModifiers.Control;

            return hotkeys.SelectAll.Any(gesture => gesture.Matches(e));
        }

        // ###########################################################################################
        // After the TextBox has had the key. Up that left the caret where it was had nothing above
        // it in the part - so the caret goes to the end of the part above; Down likewise to the
        // start of the part below. The TextBox marks both handled, hence handledEventsToo.
        // ###########################################################################################
        private void OnKeyDownAfterPart(object? sender, KeyEventArgs e)
        {
            if (e.Key is not (Key.Up or Key.Down))
                return;

            TextBox? part = this.FocusedPart();

            if (part is null || part.CaretIndex != this.thisCaretBeforeKey)
                return;

            int index = this.thisParts.IndexOf(part);

            if (e.Key == Key.Up && index > 0)
            {
                TextBox above = this.thisParts[index - 1];
                above.Focus();
                above.CaretIndex = above.Text?.Length ?? 0;
                e.Handled = true;
            }
            else if (e.Key == Key.Down && index < this.thisParts.Count - 1)
            {
                TextBox below = this.thisParts[index + 1];
                below.Focus();
                below.CaretIndex = 0;
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // A click on the field's empty space puts the caret where an ordinary text box would: below
        // the last part at the end of it; BESIDE a picture on the line under it - the start of the
        // part after that picture (code review, 2026-10-16: it went to the very end of the text).
        // FeedbackClickGeometry picks the row the click is on.
        // ###########################################################################################
        private void OnFramePressed(object? sender, PointerPressedEventArgs e)
        {
            if (!ReferenceEquals(e.Source, this.Frame) && e.Source is not ScrollContentPresenter && e.Source is not StackPanel)
                return;

            double y = e.GetPosition(this.Parts).Y;
            int row = FeedbackClickGeometry.RowAt(y, this.Parts.Children.Select(child => (child.Bounds.Top, child.Bounds.Bottom)).ToList());

            // Part i is child 2i, picture i child 2i + 1 (the parts' order, FeedbackComposer.axaml.cs).
            TextBox part = row % 2 == 1 ? this.thisParts[(row + 1) / 2] : this.thisParts[row / 2];
            bool atStart = row % 2 == 1 || y < part.Bounds.Top;

            part.Focus();
            part.CaretIndex = atStart ? 0 : part.Text?.Length ?? 0;
            e.Handled = true;
        }

        // ###########################################################################################
        // *** EVERY PASTE COMES HERE FIRST. *** The TextBox raises PastingFromClipboard before it
        // reads the clipboard - for Ctrl+V and for its menu's Paste alike - so an image-only
        // clipboard is seen. Files become pictures or attachments here; text is handed back to the
        // TextBox's own Paste(), with its own undo step (thisPassingPasteThrough lets that second
        // raise through); a picture becomes a picture only when there is no text.
        //
        // *** IN THAT ORDER (code review, 2026-10-16). *** A program often puts more than one thing
        // on the clipboard, and the first is what was meant: Excel, LibreOffice Calc and Numbers
        // put a picture of copied cells beside their text, and the cells' text was pasted as a
        // screenshot of them; Finder puts a copied file's name as text and its icon as a picture
        // beside the file.
        // ###########################################################################################
        private void OnPartPasting(object? sender, RoutedEventArgs e)
        {
            if (this.thisPassingPasteThrough || sender is not TextBox part)
                return;

            e.Handled = true;
            this.KeepUnderWay(this.PasteAsync(part), "Pasting into the feedback");
        }

        private async Task PasteAsync(TextBox part)
        {
            try
            {
                IClipboard? clipboard = TopLevel.GetTopLevel(this)?.Clipboard;

                if (clipboard is null)
                    return;

                List<string> paths = StorageItemPaths.LocalPaths(await clipboard.TryGetFilesAsync());

                if (paths.Count > 0)
                {
                    await this.AddFilesAsync(paths, part);
                    return;
                }

                if (string.IsNullOrWhiteSpace(await clipboard.TryGetTextAsync()))
                {
                    Bitmap? bitmap = await clipboard.TryGetBitmapAsync();

                    if (bitmap is not null)
                    {
                        await this.AddPastedBitmapAsync(bitmap, part);
                        return;
                    }
                }

                this.thisPassingPasteThrough = true;

                try
                {
                    // Raises PastingFromClipboard again, synchronously, before its first await.
                    part.Paste();
                }
                finally
                {
                    this.thisPassingPasteThrough = false;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Pasting into the feedback failed - [{ex}]");
            }
        }

        private void OnDragOver(object? sender, DragEventArgs e)
        {
            e.DragEffects = e.DataTransfer.Contains(DataFormat.File) || e.DataTransfer.Contains(DataFormat.Bitmap)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        // Dropped files go where the caret is, or at the end when no part has the focus - as from
        // Explorer or Finder, where CRT is not the window in front.
        private void OnDrop(object? sender, DragEventArgs e)
        {
            e.Handled = true;

            try
            {
                List<string> paths = StorageItemPaths.LocalPaths(e.DataTransfer.TryGetFiles());

                if (paths.Count > 0)
                {
                    this.KeepUnderWay(this.AddFilesAsync(paths), "Dropping onto the feedback");
                    return;
                }

                Bitmap? bitmap = e.DataTransfer.TryGetBitmap();

                if (bitmap is not null)
                    this.KeepUnderWay(this.AddPastedBitmapAsync(bitmap), "Dropping onto the feedback");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Dropping onto the feedback failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // A pasted screenshot: shown at once, as it came, under its pasted-image-N name (the next
        // number no picture's name has - a dropped file can have one); its PNG for the zip is
        // encoded in the background from this moment (0.9 s for a 4K screenshot, measured) - from
        // the very bitmap shown, so that is not let go before the encoding is done - and one wider
        // than the mail's column is then swapped for a bitmap decoded from that PNG at the column's
        // width, so the full bitmap (33 MB for 4K) is not kept. A failure there leaves the picture
        // as pasted.
        // ###########################################################################################
        internal async Task AddPastedBitmapAsync(Bitmap bitmap, TextBox? part = null)
        {
            ArgumentNullException.ThrowIfNull(bitmap);

            // Before it is shown: Windows' clipboard hands many pictures over as nothing to see.
            bitmap = await Task.Run(() => FeedbackComposer.WithOpaqueAlpha(bitmap));

            (string name, int number) = FeedbackComposition.NextPastedName(
                this.thisNextPastedNumber, this.thisPictures.Select(block => block.Picture.Name));
            this.thisNextPastedNumber = number + 1;
            Action? beforeEncoding = this.BeforeEncodingForTests;

            Task<byte[]> encoding = Task.Run(() =>
            {
                beforeEncoding?.Invoke();

                using var stream = new MemoryStream();
                bitmap.Save(stream, new PngBitmapEncoderOptions());
                return stream.ToArray();
            });

            PixelSize natural = bitmap.PixelSize;
            var picture = new FeedbackPicture(name, bitmap, natural, () => encoding, thumbnailReadBy: encoding);
            this.InsertPicture(picture, part);

            if (natural.Width <= FeedbackComposition.MaxPictureWidth)
                return;

            try
            {
                byte[] bytes = await encoding;

                Bitmap thumbnail = await Task.Run(() =>
                {
                    using var stream = new MemoryStream(bytes);
                    return Bitmap.DecodeToWidth(stream, FeedbackComposition.MaxPictureWidth, BitmapInterpolationMode.HighQuality);
                });

                this.ReplaceThumbnail(picture, thumbnail);
            }
            catch (Exception ex)
            {
                Logger.Warning($"The pasted picture [{name}] could not be encoded for sending - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // A pasted bitmap whose alpha is not what it says (PastedBitmapAlpha has why - Avalonia's
        // Windows clipboard reader labels a CF_DIB opaque with every alpha byte 0, and it was drawn
        // as nothing): its pixels copied out, every alpha set to 255, and a new opaque bitmap made
        // of them; the old one is disposed. Any other bitmap - another pixel format, or one whose
        // pixels cannot be read - is returned as it is.
        // ###########################################################################################
        internal static Bitmap WithOpaqueAlpha(Bitmap bitmap)
        {
            try
            {
                if (FeedbackComposer.ReadPixels(bitmap) is not { } pixels ||
                    !PastedBitmapAlpha.NeedsOpaqueAlpha(pixels, bitmap.AlphaFormat == AlphaFormat.Opaque))
                {
                    return bitmap;
                }

                PastedBitmapAlpha.SetOpaque(pixels);

                Bitmap mended = FeedbackComposer.FromPixels(pixels, bitmap.PixelSize, bitmap, AlphaFormat.Opaque);
                bitmap.Dispose();
                return mended;
            }
            catch (Exception ex)
            {
                // A mending that fails must not lose the paste: shown as it came, and said in the log.
                Logger.Warning($"A pasted picture's pixels could not be read to check its transparency - shown as it came [{ex.Message}]");
                return bitmap;
            }
        }

        // ###########################################################################################
        // A dropped photo turned upright by its EXIF orientation (JpegExifOrientation, code review
        // 2026-10-16), as the mail's reader shows it: its pixels turned (ExifOrientationGeometry)
        // into a new bitmap, the old one disposed. As stored - logged - when its pixels cannot be
        // read; the size it has in the mail is the composer's either way.
        // ###########################################################################################
        internal static Bitmap Upright(Bitmap bitmap, int orientation)
        {
            if (orientation is <= JpegExifOrientation.AsStored or > 8)
                return bitmap;

            try
            {
                if (FeedbackComposer.ReadPixels(bitmap) is not { } pixels)
                {
                    Logger.Warning($"A dropped photo's pixels could not be read to turn it upright - shown as stored [{bitmap.Format}]");
                    return bitmap;
                }

                PixelSize size = bitmap.PixelSize;
                (byte[] turned, int width, int height) = ExifOrientationGeometry.Apply(pixels, size.Width, size.Height, orientation);

                Bitmap upright = FeedbackComposer.FromPixels(turned, new PixelSize(width, height), bitmap, bitmap.AlphaFormat ?? AlphaFormat.Premul);
                bitmap.Dispose();
                return upright;
            }
            catch (Exception ex)
            {
                Logger.Warning($"A dropped photo could not be turned upright - shown as stored [{ex.Message}]");
                return bitmap;
            }
        }

        // A bitmap's pixels, 4 bytes each, row after row - null for any other pixel format, or for
        // more pixels than one array holds.
        private static byte[]? ReadPixels(Bitmap bitmap)
        {
            if (bitmap.Format is not { } format || (format != PixelFormats.Bgra8888 && format != PixelFormats.Rgba8888))
                return null;

            PixelSize size = bitmap.PixelSize;
            long length = (long)size.Width * 4 * size.Height;

            if (size.Width <= 0 || size.Height <= 0 || length > int.MaxValue)
                return null;

            byte[] pixels = new byte[length];
            GCHandle pin = GCHandle.Alloc(pixels, GCHandleType.Pinned);

            try
            {
                bitmap.CopyPixels(new PixelRect(size), pin.AddrOfPinnedObject(), pixels.Length, size.Width * 4);
                return pixels;
            }
            finally
            {
                pin.Free();
            }
        }

        // A new bitmap of such pixels, in `like`'s pixel format and resolution - the pixels copied.
        private static Bitmap FromPixels(byte[] pixels, PixelSize size, Bitmap like, AlphaFormat alpha)
        {
            GCHandle pin = GCHandle.Alloc(pixels, GCHandleType.Pinned);

            try
            {
                return new Bitmap(like.Format!.Value, alpha, pin.AddrOfPinnedObject(), size, like.Dpi, size.Width * 4);
            }
            finally
            {
                pin.Free();
            }
        }

        // ###########################################################################################
        // Files pasted or dropped: an image the application can display becomes a picture under its
        // own name (one its marker line can carry - FeedbackComposition.PictureNameFor - made unique
        // among the pictures), its bytes read from the file when sent; a file that is not one, or
        // will not decode, is handed to the tab for the attachments list. Each picture goes in after
        // the one before it - the focus has moved to the part below it.
        // ###########################################################################################
        internal async Task AddFilesAsync(IReadOnlyList<string> paths, TextBox? part = null)
        {
            ArgumentNullException.ThrowIfNull(paths);

            var others = new List<string>();

            foreach (string path in paths)
            {
                if (!ImageFileTypes.IsDisplayable(path) || !File.Exists(path))
                {
                    others.Add(path);
                    continue;
                }

                (Bitmap Bitmap, PixelSize NaturalSize)? decoded = await Task.Run(() => FeedbackComposer.DecodePicture(path));

                if (decoded is not { } shown)
                {
                    others.Add(path);
                    continue;
                }

                string name = FeedbackComposition.UniqueName(
                    FeedbackComposition.PictureNameFor(Path.GetFileName(path)), this.thisPictures.Select(block => block.Picture.Name));
                string source = path;

                this.InsertPicture(new FeedbackPicture(name, shown.Bitmap, shown.NaturalSize, () => File.ReadAllBytesAsync(source), filePath: source), part);
                part = null;
            }

            if (others.Count > 0)
                this.FilesForAttachment?.Invoke(this, others);
        }

        // ###########################################################################################
        // The image no wider than the mail's column - one narrower as it is (decoding TO a width
        // would scale it up) - with its own size in pixels; null for a file that is not the image its
        // name says. The size comes from the header where it can (PNG, JPEG), so a large picture is
        // decoded only to the width shown; any other format is decoded whole and then scaled, as the
        // schematic thumbnails are (SchematicThumbnailSet).
        //
        // A JPEG stored on its side (EXIF orientation 5-8, a phone's portrait photo) is as wide as
        // it stands upright - the mail's reader turns it so - and is decoded by its stored HEIGHT
        // to that width, then turned upright (code review, 2026-10-16: it was shown sideways, at a
        // size it would not have in the mail).
        // ###########################################################################################
        private static (Bitmap Bitmap, PixelSize NaturalSize)? DecodePicture(string path)
        {
            const int widest = FeedbackComposition.MaxPictureWidth;

            try
            {
                if (WorkbookPdfExporter.TryReadImageSize(path) is { } size)
                {
                    int orientation = JpegExifOrientation.Read(path);
                    (int width, int height) = ExifOrientationGeometry.OrientedSize(size.Width, size.Height, orientation);
                    var natural = new PixelSize(width, height);
                    Bitmap decoded;

                    if (width <= widest)
                    {
                        decoded = new Bitmap(path);
                    }
                    else
                    {
                        using var stream = File.OpenRead(path);
                        int across = ThumbnailDecodeGeometry.DecodeWidth(width, widest);

                        decoded = ExifOrientationGeometry.SwapsSides(orientation)
                            ? Bitmap.DecodeToHeight(stream, across, BitmapInterpolationMode.HighQuality)
                            : Bitmap.DecodeToWidth(stream, across, BitmapInterpolationMode.HighQuality);
                    }

                    return (FeedbackComposer.Upright(decoded, orientation), natural);
                }

                var full = new Bitmap(path);
                PixelSize fullSize = full.PixelSize;

                if (fullSize.Width <= widest)
                    return (full, fullSize);

                using (full)
                {
                    int width = ThumbnailDecodeGeometry.DecodeWidth(fullSize.Width, widest);
                    int height = ThumbnailDecodeGeometry.DecodeHeight(fullSize.Width, fullSize.Height, width);
                    return (full.CreateScaledBitmap(new PixelSize(width, height), BitmapInterpolationMode.HighQuality), fullSize);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read [{path}] as a picture - [{ex.Message}]");
                return null;
            }
        }
    }
}
