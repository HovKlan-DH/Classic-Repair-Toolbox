using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // *** THE FEEDBACK TEXT WITH PICTURES IN IT (owner request, 2026-10-15: pictures "directly in
    // the field, without having a huge rich-text interface"). *** Avalonia has no rich text box, so
    // this is a column of ordinary TextBoxes - the PARTS - with a picture between each pair, drawn
    // as one field: the composer's border is the box, the parts have none, one scroll bar serves
    // all. A TextBox lays out ALL its text on every key, so a keystroke here costs what the part
    // being typed in costs, whatever the number of pictures - measured 2026-10-15 in the render
    // harness: 6 ms a key with 19 pictures, and 11 ms against 45 ms for one box at 500 lines.
    //
    // What the tab SENDS is plain text with a marker line where each picture sat, and the pictures
    // themselves in the zip (CRT.Data's FeedbackComposition is the rule, and the server's too: the
    // mail shows each picture where its line sits).
    //
    // What stops at a picture, on purpose: a mouse selection, a copy across one, and Ctrl+Z
    // bringing a removed picture back (undo is each part's own). Backspace at the start of the part
    // under a picture, Delete at the end of the part above it, or its bin remove it; the parts
    // either side become one again.
    //
    // *** A PICTURE IS SHOWN AT THE SIZE IT WILL HAVE IN THE MAIL, AND ITS CORNERS RESIZE IT (owner
    // request, 2026-10-15). *** Its own width, no wider than the mail's column
    // (FeedbackComposition.MaxPictureWidth) - or than the field, which can be narrower. Pointing at
    // it shows a handle in each corner; dragging one changes the width it will have in the mail
    // (PictureResizeGeometry), with the size written on it meanwhile. The bitmap shown is decoded
    // no wider than the column (the full picture goes in the zip), so twenty of them cost tens of
    // MB, not hundreds.
    //
    // The pictures' bitmaps live until Clear or their removal - NOT until the control leaves the
    // visual tree, which a TabControl makes it do on every tab switch - and never past work still
    // reading one (FeedbackPicture.DisposeWhenFree).
    //
    // File map: this part owns the parts and the pictures, their resizing, what is sent, and Clear;
    //   FeedbackComposer.Input.cs - the keyboard across parts, clicks on the empty space, paste,
    //   drag and drop, and the pastes and drops still under way (WhenSettledAsync).
    // ###########################################################################################
    public partial class FeedbackComposer : UserControl
    {
        // The parts in order, and the pictures in order: picture i sits between parts i and i + 1.
        // In the StackPanel that is part 0, picture 0, part 1, picture 1, ... - part i at child 2i,
        // picture i at child 2i + 1.
        private readonly List<TextBox> thisParts = new();
        private readonly List<PictureBlock> thisPictures = new();

        // Pasted pictures are numbered, and a number is never handed out twice in one feedback -
        // the markers in the text stay unambiguous after a removal. Back to 1 on Clear.
        private int thisNextPastedNumber = 1;

        private string? thisPlaceholderText;

        public FeedbackComposer()
        {
            this.InitializeComponent();
            this.AddPart(string.Empty, 0);
            this.InitializeInput();
        }

        // The placeholder, shown while the field is one empty part and nothing else.
        public string? PlaceholderText
        {
            get => this.thisPlaceholderText;
            set
            {
                this.thisPlaceholderText = value;
                this.UpdatePlaceholder();
            }
        }

        public IReadOnlyList<FeedbackPicture> Pictures => this.thisPictures.Select(block => block.Picture).ToList();

        public bool IsEmpty =>
            this.thisParts.Count == 1 && this.thisPictures.Count == 0 && string.IsNullOrEmpty(this.thisParts[0].Text);

        // Raised by a drop or a paste of files that are not pictures: the tab lists them as attachments.
        public event EventHandler<IReadOnlyList<string>>? FilesForAttachment;

        internal IReadOnlyList<TextBox> PartsForTests => this.thisParts;

        // Each picture's image and its four corner handles, in order.
        internal IReadOnlyList<(Image Image, IReadOnlyList<Border> Handles)> PictureControlsForTests =>
            this.thisPictures.Select(block => (block.Image, (IReadOnlyList<Border>)block.Handles)).ToList();

        // What the tab sends beside the text: each picture's width in the mail.
        public string PictureWidthsText() =>
            FeedbackComposition.PictureWidthsText(this.thisPictures.Select(block => (block.Picture.Name, block.Picture.DisplayWidth)));

        // The text as it is sent: every part, with a marker line where each picture sits.
        public string ComposeText() =>
            FeedbackComposition.ComposeText(
                this.thisParts.Select(part => part.Text ?? string.Empty).ToList(),
                this.thisPictures.Select(block => block.Picture.Name).ToList());

        // ###########################################################################################
        // Back to one empty part, as after a send. The pictures' bitmaps go with them, and so does a
        // Ctrl+A. A field that had the keyboard keeps it, in that one part (code review,
        // 2026-10-16: Ctrl+A then Delete in a later part removed the part with the focus, and the
        // next keys went nowhere).
        // ###########################################################################################
        public void Clear()
        {
            bool hadFocus = this.thisParts.Any(part => part.IsFocused);

            foreach (PictureBlock block in this.thisPictures)
                block.Dispose();

            this.thisPictures.Clear();

            if (this.thisParts.Count > 1)
            {
                this.Parts.Children.RemoveRange(1, this.Parts.Children.Count - 1);
                this.thisParts.RemoveRange(1, this.thisParts.Count - 1);
            }

            this.thisParts[0].Text = string.Empty;
            this.thisNextPastedNumber = 1;
            this.thisAllSelected = false;
            this.UpdatePlaceholder();

            if (hadFocus)
                this.thisParts[0].Focus();
        }

        // ###########################################################################################
        // Puts a picture where the caret is - of `part` (the one pasted into), or of the focused
        // part, or at the very end: the part's text splits around it, and the part after it takes
        // the focus with the caret at its start, so typing goes on under the picture. Selected text
        // is REPLACED, as by any paste (FeedbackComposition.Split).
        //
        // The part pasted into goes by ITS caret, focused or not (code review, 2026-10-16): a paste
        // from the right-click menu may come while the menu has the focus, and the picture went to
        // the end of the paragraph. Only a drop with no part focused goes at the very end.
        // ###########################################################################################
        internal void InsertPicture(FeedbackPicture picture, TextBox? part = null)
        {
            ArgumentNullException.ThrowIfNull(picture);

            TextBox? target = part is not null && this.thisParts.Contains(part)
                ? part
                : this.thisParts.FirstOrDefault(candidate => candidate.IsFocused);

            bool atCaret = target is not null;
            part = target ?? this.thisParts[^1];

            int index = this.thisParts.IndexOf(part);
            string text = part.Text ?? string.Empty;

            // The picture takes the line the user gave it (an Enter before the paste), and gives it
            // back when removed - the block remembers which.
            PictureSplit split =
                !atCaret ? FeedbackComposition.Split(text, text.Length)
                : part.SelectionStart != part.SelectionEnd ? FeedbackComposition.Split(text, part.SelectionStart, part.SelectionEnd)
                : FeedbackComposition.Split(text, part.CaretIndex);

            // The parts change: what Ctrl+A selected is not the whole field any more.
            this.thisAllSelected = false;

            part.Text = split.Before;

            var block = new PictureBlock(picture, split, () => this.RemovePicture(picture), this);
            this.thisPictures.Insert(index, block);
            this.Parts.Children.Insert(2 * index + 1, block.Element);

            TextBox next = this.AddPart(split.After, index + 1);
            next.Focus();
            next.CaretIndex = 0;

            this.UpdatePlaceholder();
        }

        // The picture goes and the parts either side of it become one; the caret sits at the join.
        internal void RemovePicture(FeedbackPicture picture)
        {
            int index = this.thisPictures.FindIndex(block => ReferenceEquals(block.Picture, picture));

            if (index >= 0)
                this.RemovePictureAt(index);
        }

        private void RemovePictureAt(int index)
        {
            PictureBlock block = this.thisPictures[index];
            TextBox before = this.thisParts[index];
            TextBox after = this.thisParts[index + 1];
            (string text, int caret) = FeedbackComposition.Join(
                before.Text ?? string.Empty, after.Text ?? string.Empty, block.Split.LineBreakTaken, block.Split.TakenFromBefore);

            // The picture, then the part that followed it - both at the same child index.
            this.Parts.Children.RemoveAt(2 * index + 1);
            this.Parts.Children.RemoveAt(2 * index + 1);
            this.thisPictures.RemoveAt(index);
            this.thisParts.RemoveAt(index + 1);
            this.thisAllSelected = false;
            block.Dispose();

            before.Text = text;
            before.Focus();
            before.CaretIndex = caret;

            this.UpdatePlaceholder();
        }

        // ###########################################################################################
        // A pasted screenshot is shown as pasted first and swapped for its thumbnail once the
        // background work has it (FeedbackComposer.Input.cs). The one shown until now is let go on
        // the next dispatcher cycle, after the swap has been drawn and once nothing reads it
        // (FeedbackPicture.DisposeWhenFree); a picture removed meanwhile takes the new one with it.
        // ###########################################################################################
        internal void ReplaceThumbnail(FeedbackPicture picture, Bitmap thumbnail)
        {
            PictureBlock? block = this.thisPictures.FirstOrDefault(candidate => ReferenceEquals(candidate.Picture, picture));

            if (block is null)
            {
                thumbnail.Dispose();
                return;
            }

            Bitmap? shownUntilNow = picture.Thumbnail;
            Task? readBy = picture.ThumbnailReadBy;
            block.Image.Source = thumbnail;
            picture.Thumbnail = thumbnail;
            picture.ThumbnailReadBy = null;

            if (shownUntilNow is not null)
                FeedbackPicture.DisposeWhenFree(shownUntilNow, readBy);
        }

        private TextBox AddPart(string text, int index)
        {
            TextBox part = this.MakePart(text);
            this.thisParts.Insert(index, part);
            this.Parts.Children.Insert(2 * index, part);
            return part;
        }

        private TextBox MakePart(string text)
        {
            var part = new TextBox
            {
                Text = text,
                AcceptsReturn = true,
                AcceptsTab = false,
                TextWrapping = TextWrapping.Wrap,
            };

            part.Classes.Add("Part");

            // The part never scrolls by itself - the composer's one scroll bar does.
            ScrollViewer.SetVerticalScrollBarVisibility(part, ScrollBarVisibility.Disabled);
            ScrollViewer.SetHorizontalScrollBarVisibility(part, ScrollBarVisibility.Disabled);

            part.AddHandler(TextBox.PastingFromClipboardEvent, this.OnPartPasting);

            return part;
        }

        private void UpdatePlaceholder()
        {
            if (this.thisParts.Count == 0)
                return;

            this.thisParts[0].PlaceholderText = this.thisParts.Count == 1 && this.thisPictures.Count == 0 ? this.thisPlaceholderText : null;
        }

        // ###########################################################################################
        // A picture as shown: the image, left-aligned, at the width it will have in the mail (no
        // wider than the field); the attachments list's bin near its top-right corner and a resize
        // handle ON each corner, shown while the pointer is over it; and its size written on it
        // while a corner is dragged.
        // ###########################################################################################
        private sealed class PictureBlock : IDisposable
        {
            private const double HandleSize = 10;

            // The control the drag is measured against - one that does not move as the picture grows.
            private readonly Visual thisReference;
            private readonly Border thisSizeLabel;
            private readonly TextBlock thisSizeText;

            private PictureCorner? thisDragCorner;
            private Point thisDragStart;
            private int thisDragStartWidth;

            public PictureBlock(FeedbackPicture picture, PictureSplit split, Action remove, Visual reference)
            {
                this.Picture = picture;
                this.Split = split;
                this.thisReference = reference;

                // Uniform within a box the size it has in the mail: so it is that size, or the field's
                // width when that is narrower, whatever size the bitmap shown was decoded at.
                this.Image = new Image
                {
                    Source = picture.Thumbnail,
                    Stretch = Stretch.Uniform,
                    StretchDirection = StretchDirection.Both,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top,
                };

                var bin = new Button
                {
                    Content = "🗑️",
                    FontSize = 12,
                    Padding = new Thickness(2, 2),
                    Margin = new Thickness(0, 8, 8, 0),
                    Focusable = false,
                    IsTabStop = false,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Top,
                };

                bin.Classes.Add("RemovePicture");
                ToolTip.SetTip(bin, "Remove picture");
                bin.Click += (_, _) => remove();

                this.thisSizeText = new TextBlock { FontSize = 11 };
                this.thisSizeLabel = new Border
                {
                    Child = this.thisSizeText,
                    Padding = new Thickness(6, 2),
                    Margin = new Thickness(0, 0, 8, 8),
                    CornerRadius = new CornerRadius(3),
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Bottom,
                    IsHitTestVisible = false,
                    IsVisible = false,
                };

                this.thisSizeLabel.Classes.Add("PictureSize");

                this.Handles =
                [
                    this.MakeHandle(PictureCorner.TopLeft, HorizontalAlignment.Left, VerticalAlignment.Top, StandardCursorType.TopLeftCorner),
                    this.MakeHandle(PictureCorner.TopRight, HorizontalAlignment.Right, VerticalAlignment.Top, StandardCursorType.TopRightCorner),
                    this.MakeHandle(PictureCorner.BottomLeft, HorizontalAlignment.Left, VerticalAlignment.Bottom, StandardCursorType.BottomLeftCorner),
                    this.MakeHandle(PictureCorner.BottomRight, HorizontalAlignment.Right, VerticalAlignment.Bottom, StandardCursorType.BottomRightCorner),
                ];

                this.Element = new Grid
                {
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Margin = new Thickness(10, 6, 8, 6),
                    Children = { this.Image, bin, this.thisSizeLabel },
                };

                foreach (Border handle in this.Handles)
                    this.Element.Children.Add(handle);

                this.Element.Classes.Add("Picture");
                this.ShowDisplayWidth();
            }

            public FeedbackPicture Picture { get; }

            // How the text was split around it - the line break it took, to give back on removal.
            public PictureSplit Split { get; }

            public Grid Element { get; }

            public Image Image { get; }

            public Border[] Handles { get; }

            public void Dispose()
            {
                this.Image.Source = null;
                this.Picture.ReleaseThumbnail();
            }

            private Border MakeHandle(PictureCorner corner, HorizontalAlignment horizontal, VerticalAlignment vertical, StandardCursorType cursor)
            {
                // On the corner, half outside the picture, so the whole corner can be taken.
                double half = PictureBlock.HandleSize / 2;

                var handle = new Border
                {
                    Width = PictureBlock.HandleSize,
                    Height = PictureBlock.HandleSize,
                    HorizontalAlignment = horizontal,
                    VerticalAlignment = vertical,
                    Margin = new Thickness(
                        horizontal == HorizontalAlignment.Left ? -half : 0,
                        vertical == VerticalAlignment.Top ? -half : 0,
                        horizontal == HorizontalAlignment.Right ? -half : 0,
                        vertical == VerticalAlignment.Bottom ? -half : 0),
                    Cursor = new Cursor(cursor),
                    Tag = corner,
                };

                handle.Classes.Add("ResizeHandle");
                handle.PointerPressed += this.OnHandlePressed;
                handle.PointerMoved += this.OnHandleMoved;
                handle.PointerReleased += this.OnHandleReleased;
                handle.PointerCaptureLost += (_, _) => this.EndDrag();

                return handle;
            }

            private void OnHandlePressed(object? sender, PointerPressedEventArgs e)
            {
                if (sender is not Border handle || handle.Tag is not PictureCorner corner || !e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed)
                    return;

                // From what is SHOWN, which a narrow field can have made smaller than the width chosen.
                double shown = this.Image.Bounds.Width;
                this.thisDragStartWidth = shown > 0 ? Math.Min(this.Picture.DisplayWidth, (int)Math.Round(shown)) : this.Picture.DisplayWidth;
                this.thisDragStart = e.GetPosition(this.thisReference);
                this.thisDragCorner = corner;

                e.Pointer.Capture(handle);
                this.Element.Classes.Set("Resizing", true);
                this.thisSizeLabel.IsVisible = true;
                this.ShowDisplayWidth();
                e.Handled = true;
            }

            private void OnHandleMoved(object? sender, PointerEventArgs e)
            {
                if (this.thisDragCorner is not PictureCorner corner)
                    return;

                Point now = e.GetPosition(this.thisReference);
                PixelSize natural = this.Picture.NaturalSize;

                this.Picture.DisplayWidth = PictureResizeGeometry.DraggedWidth(
                    this.thisDragStartWidth, natural.Width, natural.Height, corner, now.X - this.thisDragStart.X, now.Y - this.thisDragStart.Y);

                this.ShowDisplayWidth();
                e.Handled = true;
            }

            private void OnHandleReleased(object? sender, PointerReleasedEventArgs e)
            {
                if (this.thisDragCorner is null)
                    return;

                this.EndDrag();
                e.Pointer.Capture(null);
                e.Handled = true;
            }

            private void EndDrag()
            {
                this.thisDragCorner = null;
                this.Element.Classes.Set("Resizing", false);
                this.thisSizeLabel.IsVisible = false;
            }

            private void ShowDisplayWidth()
            {
                this.Image.MaxWidth = this.Picture.DisplayWidth;
                this.Image.MaxHeight = this.Picture.DisplayHeight;
                this.thisSizeText.Text = $"{this.Picture.DisplayWidth} x {this.Picture.DisplayHeight}";
            }
        }
    }

    // ###########################################################################################
    // A picture in the feedback: the name it travels under in the zip (and its marker in the
    // text), the bitmap shown for it, its own size in pixels and the width it will have in the mail
    // (its own, no wider than the mail's column, until its corners are dragged), and its bytes for
    // the zip - a pasted picture's PNG, encoded in the background from the moment it was pasted; a
    // dropped file's bytes, read when sent (FilePath, which the zip streams from).
    // ###########################################################################################
    public sealed class FeedbackPicture
    {
        private readonly Func<Task<byte[]>> thisReadBytes;

        public FeedbackPicture(string name, Bitmap? thumbnail, PixelSize naturalSize, Func<Task<byte[]>> readBytes, Task? thumbnailReadBy = null, string? filePath = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            ArgumentNullException.ThrowIfNull(readBytes);

            this.Name = name;
            this.Thumbnail = thumbnail;
            this.ThumbnailReadBy = thumbnailReadBy;
            this.NaturalSize = naturalSize;
            this.DisplayWidth = FeedbackComposition.DefaultPictureWidth(naturalSize.Width);
            this.FilePath = filePath;
            this.thisReadBytes = readBytes;
        }

        public string Name { get; }

        // The file a dropped picture is, or null for a pasted one.
        public string? FilePath { get; }

        public Bitmap? Thumbnail { get; internal set; }

        // Work still reading Thumbnail on another thread - a pasted picture's PNG, encoded from the
        // very bitmap shown - or null. Cleared when the bitmap shown is one nothing else reads.
        internal Task? ThumbnailReadBy { get; set; }

        // ###########################################################################################
        // *** A BITMAP IS NEVER DISPOSED UNDER WORK STILL READING IT (code review, 2026-10-16). ***
        // A pasted picture's PNG is encoded from the bitmap shown, on the pool, for up to a second
        // (4K); removing the picture meanwhile - its bin, Backspace, Ctrl+A then Delete, the Clear
        // after a send - disposed the bitmap under the encoder, which then read freed memory, and
        // CRT could end. It is let go on the next dispatcher cycle (after the frame that stopped
        // showing it), and never before `readBy` is done with it.
        // ###########################################################################################
        internal static void DisposeWhenFree(Bitmap bitmap, Task? readBy)
        {
            ArgumentNullException.ThrowIfNull(bitmap);

            if (readBy is null || readBy.IsCompleted)
                Dispatcher.UIThread.Post(bitmap.Dispose, DispatcherPriority.Background);
            else
                readBy.ContinueWith(_ => Dispatcher.UIThread.Post(bitmap.Dispose, DispatcherPriority.Background), TaskScheduler.Default);
        }

        // The bitmap shown let go - the picture was removed or the field cleared.
        internal void ReleaseThumbnail()
        {
            Bitmap? thumbnail = this.Thumbnail;
            Task? readBy = this.ThumbnailReadBy;

            this.Thumbnail = null;
            this.ThumbnailReadBy = null;

            if (thumbnail is not null)
                FeedbackPicture.DisposeWhenFree(thumbnail, readBy);
        }

        public PixelSize NaturalSize { get; }

        public int DisplayWidth { get; internal set; }

        public int DisplayHeight => PictureResizeGeometry.HeightFor(this.DisplayWidth, this.NaturalSize.Width, this.NaturalSize.Height);

        public Task<byte[]> ReadBytesAsync() => this.thisReadBytes();
    }
}
