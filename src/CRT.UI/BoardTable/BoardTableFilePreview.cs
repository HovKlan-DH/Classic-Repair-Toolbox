using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // WHAT HOVERING A FILE CELL SHOWS (owner request, 2026-09-26): the picture, or - for a changed
    // one - the published picture and the new one side by side; for a file that cannot be drawn (a
    // PDF), a link that opens it.
    //
    // WHICH file each side is, is BoardTableFileCells' (CRT.Data); the bytes come from the host's
    // IBoardTableFileSource. This only lays them out. Built in code rather than markup, as the
    // Maintainer tab's own image comparison is, because every part depends on what loads.
    //
    // *** THE SAME PATH ON BOTH SIDES IS STILL A COMPARISON. *** A picture replaced under its own
    // name leaves the cell's text unchanged, so the cell is not orange - and it is exactly the change
    // worth seeing. Both sides are read, and only when the bytes are identical does the preview show
    // one picture, as "Unchanged" (unless the host's SaysUnchanged is false - the maintainer
    // application's new system). Not done for a file that cannot be drawn: reading a PDF only to
    // compare it would fetch megabytes on every hover, and its link opens it anyway.
    //
    // *** A DECODE FAILURE MUST NOT CRASH ANYTHING. *** The bytes are contributor-supplied - a
    // truncated upload, or not an image at all despite the name - and this runs under the pointer.
    // Every failure is written into the preview instead.
    //
    // *** A CARD THE POINTER HAS LEFT DECODES NOTHING. *** The card follows the pointer from cell to
    // cell at once, so sweeping down a column builds a card per row passed. Once released, a card
    // skips the decode its read was waiting for, and the decode itself runs off the UI thread, so a
    // large scan does not stall the pointer.
    // ###########################################################################################
    public sealed class BoardTableFilePreview : Border
    {
        // Pictures are shown at most this big; a board scan is decoded and then scaled down to
        // this width, so a 4000-pixel scan does not sit in memory at full size behind a small view.
        private const double ImageMaxWidth = 360;
        private const double ImageMaxHeight = 300;
        private const int DecodedMaxWidth = 720;

        private readonly BoardTableFileCell thisFile;
        private readonly IBoardTableFileSource thisSource;
        private readonly TextBlock thisHeadline;
        private readonly List<Bitmap> thisBitmaps = [];

        // Set by ReleaseImages: the card has closed, and a read still in flight is not decoded.
        private bool thisReleased;

        // Off for a card that says nothing about a change - the Maintainer tab's file tree, whose row
        // already says it (owner request, 2026-09-28: "only the relative path and then the other
        // functionality from the table format").
        private readonly bool thisShowsHeadline;

        public BoardTableFilePreview(BoardTableFileCell file, IBoardTableFileSource source, string? note = null, bool showHeadline = true)
        {
            ArgumentNullException.ThrowIfNull(file);
            ArgumentNullException.ThrowIfNull(source);

            this.thisFile = file;
            this.thisSource = source;
            this.thisShowsHeadline = showHeadline;

            this.MaxWidth = (ImageMaxWidth * 2) + 48;

            var content = new StackPanel { Spacing = 6 };

            if (!string.IsNullOrWhiteSpace(note))
            {
                content.Children.Add(new TextBlock
                {
                    Text = note,
                    TextWrapping = TextWrapping.Wrap,
                    FontStyle = FontStyle.Italic
                });
            }

            this.thisHeadline = new TextBlock
            {
                Text = BoardTableFileCells.Headline(file, sameContent: null) ?? string.Empty,
                FontWeight = FontWeight.SemiBold,
                TextWrapping = TextWrapping.Wrap
            };
            this.thisHeadline.IsVisible = showHeadline && this.thisHeadline.Text!.Length > 0;
            content.Children.Add(this.thisHeadline);

            var sides = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 16 };
            content.Children.Add(sides);

            this.Child = content;

            this.Loading = this.BuildSidesAsync(sides);
        }

        // Completes when every side has loaded (or failed to) - for tests, and nothing else.
        public Task Loading { get; }

        // The sides on screen, by what they are labelled - for tests.
        internal IReadOnlyList<Control> Sides => ((this.Child as StackPanel)!.Children.OfType<StackPanel>().Last()).Children;

        internal string HeadlineText => this.thisHeadline.IsVisible ? this.thisHeadline.Text ?? string.Empty : string.Empty;

        // ###########################################################################################
        // Frees the decoded pictures. Called when the hover card closes. Each Image lets go of its
        // picture FIRST: a bitmap disposed while an Image still shows it fails on the render thread,
        // which takes the whole application down (the Workbooks tab crash CLAUDE.md records).
        // ###########################################################################################
        public void ReleaseImages()
        {
            this.thisReleased = true;

            foreach (Image image in this.GetImages())
            {
                image.Source = null;
            }

            foreach (Bitmap bitmap in this.thisBitmaps)
            {
                bitmap.Dispose();
            }

            this.thisBitmaps.Clear();
        }

        private IEnumerable<Image> GetImages() =>
            this.Sides.OfType<StackPanel>().SelectMany(side => side.Children.OfType<Image>());

        private async Task BuildSidesAsync(StackPanel sides)
        {
            string? published = this.thisFile.PublishedPath;
            string? current = this.thisFile.CurrentPath;

            // A file that cannot be drawn, under the same name on both sides: one link, to the
            // current file - see the header for why it is not compared.
            if (this.thisFile.IsSamePath && !ImageFileTypes.IsDisplayable(current))
            {
                sides.Children.Add(this.BuildSide(current!, BoardTableFileSide.Current, this.thisSource.CurrentLabel, out _));
                BoardTableFilePreview.ShowSideLabels(sides);
                return;
            }

            var loads = new List<Task<byte[]?>>();
            StackPanel? publishedSide = null;

            if (published is not null)
            {
                publishedSide = this.BuildSide(published, BoardTableFileSide.Published, this.thisSource.PublishedLabel, out Task<byte[]?>? load);
                sides.Children.Add(publishedSide);

                if (load is not null)
                    loads.Add(load);
            }

            if (current is not null)
            {
                sides.Children.Add(this.BuildSide(current, BoardTableFileSide.Current, this.thisSource.CurrentLabel, out Task<byte[]?>? load));

                if (load is not null)
                    loads.Add(load);
            }

            BoardTableFilePreview.ShowSideLabels(sides);

            await Task.WhenAll(loads);

            if (!this.thisFile.IsSamePath || loads.Count != 2)
                return;

            byte[]? before = loads[0].Result;
            byte[]? after = loads[1].Result;

            if (before is null || after is null)
                return;

            bool same = before.AsSpan().SequenceEqual(after);

            this.thisHeadline.Text = same && !this.thisSource.SaysUnchanged
                ? string.Empty
                : BoardTableFileCells.Headline(this.thisFile, same) ?? string.Empty;
            this.thisHeadline.IsVisible = this.thisShowsHeadline && this.thisHeadline.Text.Length > 0;

            // Identical: one picture is the whole story.
            if (same && publishedSide is not null)
            {
                publishedSide.IsVisible = false;
                BoardTableFilePreview.ShowSideLabels(sides);
            }
        }

        // ###########################################################################################
        // *** A SIDE IS LABELLED ONLY BESIDE ANOTHER ONE (owner request, 2026-09-26). *** Next to
        // each other, "Before (published)" and "After (submitted)" say which picture is which. A
        // side on its own needs no label - the headline already says "Removed" or "New file", and
        // which one it is is obvious.
        // ###########################################################################################
        private static void ShowSideLabels(StackPanel sides)
        {
            List<StackPanel> all = sides.Children.OfType<StackPanel>().ToList();
            bool paired = all.Count(side => side.IsVisible) > 1;

            foreach (StackPanel side in all)
                side.Children[0].IsVisible = paired;
        }

        // ###########################################################################################
        // One side: its label, the path, then the picture (loaded) or a link to open the file.
        // `load` is the read of a picture's bytes, or null for a file that is not drawn.
        // ###########################################################################################
        private StackPanel BuildSide(string path, BoardTableFileSide side, string label, out Task<byte[]?>? load)
        {
            var column = new StackPanel { Spacing = 4, Width = ImageMaxWidth };

            column.Children.Add(new TextBlock { Text = label, Opacity = 0.7 });
            column.Children.Add(new TextBlock
            {
                Text = path,
                TextWrapping = TextWrapping.Wrap,
                FontSize = 11
            });

            var status = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap,
                Opacity = 0.7,
                FontStyle = FontStyle.Italic
            };

            bool isImage = ImageFileTypes.IsDisplayable(path);
            Button link = this.BuildOpenLink(path, side, isImage, status);

            if (isImage)
            {
                var image = new Image
                {
                    Stretch = Stretch.Uniform,
                    StretchDirection = StretchDirection.DownOnly,
                    MaxWidth = ImageMaxWidth,
                    MaxHeight = ImageMaxHeight,
                    HorizontalAlignment = HorizontalAlignment.Left
                };

                status.Text = "Loading...";
                column.Children.Add(image);
                column.Children.Add(status);

                load = this.LoadImageAsync(path, side, image, status, link);
            }
            else
            {
                load = null;
                column.Children.Add(status);
                status.IsVisible = false;
            }

            column.Children.Add(link);

            return column;
        }

        // ###########################################################################################
        // Reads and draws one picture, returning the bytes (null when there were none) so the two
        // sides of the same path can be compared.
        // ###########################################################################################
        // A picture with no bytes has nothing to open, so its link goes with it.
        private async Task<byte[]?> LoadImageAsync(string path, BoardTableFileSide side, Image target, TextBlock status, Button link)
        {
            byte[]? bytes;

            try
            {
                bytes = await this.thisSource.ReadAsync(path, side);
            }
            catch (Exception exception)
            {
                // The source promises not to throw; a host that does still must not crash the grid.
                status.Text = $"This file could not be read ({exception.GetType().Name}).";
                link.IsVisible = false;
                return null;
            }

            if (bytes is null)
            {
                status.Text = "There is no file at this path.";
                link.IsVisible = false;
                return null;
            }

            if (this.thisReleased)
            {
                return bytes;
            }

            try
            {
                Bitmap bitmap = await Task.Run(() => BoardTableFilePreview.DecodeForPreview(bytes));

                // Closed while it decoded: nothing will show it.
                if (this.thisReleased)
                {
                    bitmap.Dispose();
                    return bytes;
                }

                this.thisBitmaps.Add(bitmap);

                target.Source = bitmap;
                status.IsVisible = false;
            }
            catch (Exception exception)
            {
                // Broad on purpose: Avalonia's decoder reports a malformed image through more than
                // one exception type, depending on the platform's imaging backend.
                status.Text = $"This file could not be shown as a picture ({exception.GetType().Name}).";
            }

            return bytes;
        }

        private static Bitmap DecodeForPreview(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes, writable: false);
            var full = new Bitmap(stream);

            if (full.PixelSize.Width <= DecodedMaxWidth)
                return full;

            int height = Math.Max(1, (int)Math.Round(full.PixelSize.Height * (double)DecodedMaxWidth / full.PixelSize.Width));

            try
            {
                return full.CreateScaledBitmap(new PixelSize(DecodedMaxWidth, height));
            }
            finally
            {
                full.Dispose();
            }
        }

        // ###########################################################################################
        // "Open PDF file" - or, for a picture, "Open full size". Opened by the host (its viewer, the
        // system's default program); a refusal is written under it rather than lost.
        // ###########################################################################################
        private Button BuildOpenLink(string path, BoardTableFileSide side, bool isImage, TextBlock status)
        {
            string extension = Path.GetExtension(path).TrimStart('.').ToUpperInvariant();

            string text = isImage ? "Open full size"
                : extension.Length == 0 ? "Open file"
                : $"Open {extension} file";

            var link = new Button
            {
                Content = new TextBlock
                {
                    Text = text,
                    TextDecorations = TextDecorations.Underline
                },
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(0),
                Foreground = Brushes.SteelBlue,
                Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                HorizontalAlignment = HorizontalAlignment.Left,

                // Clicked, never focused: the keyboard stays with the table (Ctrl+Z, typing).
                Focusable = false
            };

            link.Click += async (_, _) =>
            {
                string? problem = await this.thisSource.OpenAsync(path, side);

                if (problem is not null)
                {
                    status.Text = problem;
                    status.IsVisible = true;
                }
            };

            return link;
        }
    }
}
