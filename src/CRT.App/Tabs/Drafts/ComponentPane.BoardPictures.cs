using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE BOARD'S PICTURES AND THEIR COLOURS (CRT 4.0 step 2, slice 4b, 2026-10-26 - the plan: "Board
    // > Pictures: a colour picker and opacity slider per picture, with a live preview on the picture,
    // plus 'Use for all pictures'"). Until now only the table, as text, reached the colours.
    //
    //   - EACH PICTURE is a card: the picture made small with the components marked on it, drawn in the
    //     colours picked as they are picked (PicturePreview), and three of them - the marked components
    //     and how strongly they show, the same on the thumbnail, and a trace on the board's other side.
    //   - A CHANGE goes into the draft as the pane's others do - after a pause, ONE step for every
    //     picture's colours (BoardPartEdits.WithPictureColours, by name AND image file). A colour the
    //     picker merely shows - CRT's own default for a blank one - writes nothing (the row's guard).
    //   - "USE FOR ALL PICTURES" gives every picture these colours, as one step.
    //   - Adding, removing, renaming and ordering pictures stays the "Schematic images" window's
    //     (ManagePicturesRequested - the host opens it); the part reads the draft again after it.
    //   - EACH CARD'S TOOLS: "Mark components" and "Align KiCad" open the Schematics tab in that tool
    //     (PictureToolRequested); "Mark all automatically" (step 3, slice 3c), on a picture aligned with
    //     KiCad, has the host draw every component's rectangle from its footprint
    //     (MarkAutomaticallyRequested - TabDrafts.MarkAutomatically.cs), said under the heading.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The width a picture is decoded to for its preview - twice the width it is drawn at.
        private const int PreviewWidth = 640;

        private readonly ObservableCollection<ComponentPaneBoardPictureRow> thisBoardPictures = new();
        private readonly List<Task> thisPreviewLoads = new();

        // Asked to open the "Schematic images" window on the board shown.
        public event EventHandler? ManagePicturesRequested;

        internal IReadOnlyList<ComponentPaneBoardPictureRow> BoardPicturesForTests => this.thisBoardPictures;

        // Every preview read and decoded - for a test.
        internal Task PreviewsLoadedForTests => Task.WhenAll(this.thisPreviewLoads.ToList());

        private void InitializeBoardPictures()
        {
            this.BoardPicturesItemsControl.ItemsSource = this.thisBoardPictures;
        }

        // ###########################################################################################
        // The pictures read again from `board` - a card each, with the highlights of the region shown -
        // their previews loaded off the UI thread.
        // ###########################################################################################
        private void ShowBoardPictures(BoardData board)
        {
            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures)
            {
                row.PropertyChanged -= this.OnPictureRowChanged;
            }

            Dictionary<string, Dictionary<string, List<Rect>>> highlights = HighlightRectBuilder.BuildHighlightRects(board, this.thisRegion);

            // "Align KiCad" needs the board's KiCad data - names under "KiCad data/" in the draft's list;
            // "Mark all automatically" a picture aligned in the draft too.
            bool boardHasKiCad = this.thisSession?.ListedFiles.Keys.Any(SubmissionKiCadFiles.IsInKiCadFolder) == true;
            var aligned = new HashSet<string>(
                this.thisSession?.Content.Calibrations.Select(entry => entry.SchematicName?.Trim() ?? string.Empty) ?? [],
                StringComparer.OrdinalIgnoreCase);

            ComponentPane.Fill(
                this.thisBoardPictures,
                board.Schematics.Select(row => ComponentPaneBoardPictureRow.From(
                    row,
                    highlights.TryGetValue(row.SchematicName?.Trim() ?? string.Empty, out var byLabel)
                        ? byLabel.Values.SelectMany(rects => rects).ToList()
                        : [],
                    boardHasKiCad,
                    aligned.Contains(row.SchematicName?.Trim() ?? string.Empty))));

            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures)
            {
                row.PropertyChanged += this.OnPictureRowChanged;
            }

            this.BoardPicturesHeading.Text = ComponentPaneWording.Heading("Pictures", this.thisBoardPictures.Count);
            this.NoBoardPicturesText.IsVisible = this.thisBoardPictures.Count == 0;

            this.LoadPreviews();
        }

        // A colour or opacity picked - into the draft once picking pauses, as typing does.
        private void OnPictureRowChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (this.thisShowing || !ComponentPaneBoardPictureRow.IsColourChange(e.PropertyName))
            {
                return;
            }

            this.thisTypingTimer.Stop();
            this.thisTypingTimer.Start();
        }

        private void CommitPictureColours()
        {
            List<(string Name, string File, PictureColours Colours)> colours = this.thisBoardPictures
                .Select(row => (row.Name, row.File, row.ToColours()))
                .ToList();

            this.ApplyBoardPart(DraftStepWording.PictureColours, board => BoardPartEdits.WithPictureColours(board, colours));
        }

        // "Use for all pictures": this card's colours on every card, as one step.
        private void OnUseColoursForAllClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneBoardPictureRow row })
            {
                this.UseColoursForAll(row);
            }
        }

        internal void UseColoursForAll(ComponentPaneBoardPictureRow source)
        {
            PictureColours colours = source.ToColours();

            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures)
            {
                row.Use(colours);
            }

            this.thisTypingTimer.Stop();
            this.CommitPictureColours();
        }

        private void OnManagePicturesClick(object? sender, RoutedEventArgs e)
        {
            this.CommitHeld();
            this.ManagePicturesRequested?.Invoke(this, EventArgs.Empty);
        }

        // ###########################################################################################
        // "Mark components" and "Align KiCad" on a picture (slice 6 - contributor mode went): what the
        // pane holds handed over first, then the host opens the Schematics tab in that tool.
        // ###########################################################################################
        internal event EventHandler<(PictureTool Tool, string Picture)>? PictureToolRequested;

        private void OnMarkComponentsClick(object? sender, RoutedEventArgs e) =>
            this.RequestPictureTool(sender, PictureTool.MarkComponents);

        private void OnAlignKiCadClick(object? sender, RoutedEventArgs e) =>
            this.RequestPictureTool(sender, PictureTool.AlignKiCad);

        private void RequestPictureTool(object? sender, PictureTool tool)
        {
            if (sender is Control { Tag: ComponentPaneBoardPictureRow row })
            {
                this.RequestPictureTool(row, tool);
            }
        }

        internal void RequestPictureTool(ComponentPaneBoardPictureRow row, PictureTool tool)
        {
            this.thisTypingTimer.Stop();
            this.CommitHeld();
            this.PictureToolRequested?.Invoke(this, (tool, row.Name));
        }

        // ###########################################################################################
        // "Mark all automatically" on a picture aligned with KiCad (step 3, slice 3c): what the pane
        // holds handed over first, then the host draws the rectangles into the draft
        // (TabDrafts.MarkAutomaticallyAsync) and says what came of it here (ShowPicturesNote).
        // ###########################################################################################
        internal event EventHandler<string>? MarkAutomaticallyRequested;

        private void OnMarkAutomaticallyClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneBoardPictureRow row })
            {
                this.RequestMarkAutomatically(row);
            }
        }

        internal void RequestMarkAutomatically(ComponentPaneBoardPictureRow row)
        {
            this.thisTypingTimer.Stop();
            this.CommitHeld();
            this.MarkAutomaticallyRequested?.Invoke(this, row.Name);
        }

        // The draft's content the note was said of - it goes once the draft moves on.
        private BoardContent? thisPicturesNoteContent;

        // ###########################################################################################
        // What the last "Mark all automatically" did - null to say nothing. Said only while the draft is
        // the content it was said of (code review, 2026-10-26: after an Undo it still said the marks
        // were made) - the part read again drops it (DropPicturesNoteIfMovedOn).
        // ###########################################################################################
        internal void ShowPicturesNote(string? note)
        {
            this.thisPicturesNoteContent = this.thisSession?.Content;
            this.BoardPicturesNoteText.Text = note;
            this.BoardPicturesNoteText.IsVisible = !string.IsNullOrWhiteSpace(note);
        }

        private void DropPicturesNoteIfMovedOn()
        {
            if (this.BoardPicturesNoteText.IsVisible && !ReferenceEquals(this.thisSession?.Content, this.thisPicturesNoteContent))
            {
                this.ShowPicturesNote(null);
            }
        }

        internal string? PicturesNoteForTests => this.BoardPicturesNoteText.IsVisible ? this.BoardPicturesNoteText.Text : null;

        // ---- previews -----------------------------------------------------------------------------

        private void LoadPreviews()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            string dataRoot = DataManager.DataRoot;
            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);
            string identity = session.ExcelDataFile;

            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures.Where(row => row.Thumbnail is null && row.File.Length > 0))
            {
                this.thisPreviewLoads.Add(this.LoadPreviewAsync(row, dataRoot, draftFolder, identity));
            }
        }

        // ###########################################################################################
        // A picture's preview, read once while the pane is open (PaneImageCache - code review,
        // 2026-10-26: every re-read of the part decoded every picture again), with the picture's own
        // pixel size read from its header for every type CRT shows (ImagePixelSize).
        // ###########################################################################################
        private async Task LoadPreviewAsync(ComponentPaneBoardPictureRow row, string dataRoot, string draftFolder, string identity)
        {
            string file = row.File;
            int generation = this.thisImages.Generation;

            (string? path, string? key) = await Task.Run(() =>
            {
                string? found = DraftFileResolver.Resolve(dataRoot, draftFolder, file, identity);
                return found is not null && ContributionPackaging.IsDisplayableImageFile(found)
                    ? (found, PaneImageCache.KeyOf(found, ComponentPane.PreviewWidth))
                    : (found, null);
            });

            if (path is null || key is null)
            {
                return;
            }

            if (!this.thisImages.TryGet(key, out PaneImage image))
            {
                image = await Task.Run(() => ComponentPane.DecodePreview(path));
                this.thisImages.Keep(key, image, generation);
            }

            if (generation == this.thisImages.Generation && string.Equals(row.File, file, StringComparison.Ordinal))
            {
                row.ImageSize = image.OriginalSize;
                row.Thumbnail = image.Bitmap;
            }
        }

        private static PaneImage DecodePreview(string path)
        {
            try
            {
                Size pixels = ImagePixelSize.TryRead(path) is { } read ? new Size(read.Width, read.Height) : default;

                using FileStream stream = File.OpenRead(path);
                return new PaneImage(Bitmap.DecodeToWidth(stream, ComponentPane.PreviewWidth, BitmapInterpolationMode.MediumQuality), pixels);
            }
            catch (Exception ex)
            {
                // A picture that does not decode is a card without a preview, never a pane that fails.
                Logger.Warning($"Could not read the picture [{path}] for its preview - [{ex.Message}]");
                return PaneImage.None;
            }
        }
    }

    // ###########################################################################################
    // One of the board's pictures as the Pictures part holds it - its colours as the board keeps them
    // (text), shown to the pickers as colours and percentages. A blank colour shows CRT's own default,
    // and a picker writing back the very colour it shows is no change (the setters' guard): a picker
    // that writes its value as it is built would otherwise turn every blank colour into a written one.
    // ###########################################################################################
    public sealed class ComponentPaneBoardPictureRow : ComponentPaneRowBase
    {
        // What CRT draws with when a colour or an opacity is blank (the Schematics tab's defaults).
        private static readonly Color DefaultHighlight = Colors.IndianRed;
        private static readonly Color DefaultOppositeTrace = Colors.DodgerBlue;
        private const double DefaultOpacity = 0.20;

        private string thisHighlightColor = string.Empty;
        private string thisHighlightOpacity = string.Empty;
        private string thisOppositeTraceColor = string.Empty;
        private string thisThumbnailColor = string.Empty;
        private string thisThumbnailOpacity = string.Empty;
        private Bitmap? thisThumbnail;
        private Size thisImageSize;

        public string Name { get; init; } = string.Empty;

        public string File { get; init; } = string.Empty;

        // The picture's highlights, in its own pixels.
        public IReadOnlyList<Rect> Highlights { get; init; } = [];

        public Bitmap? Thumbnail
        {
            get => this.thisThumbnail;
            set
            {
                this.thisThumbnail = value;
                this.Raise(nameof(this.Thumbnail));
            }
        }

        public Size ImageSize
        {
            get => this.thisImageSize;
            set
            {
                this.thisImageSize = value;
                this.Raise(nameof(this.ImageSize));
            }
        }

        public Color HighlightColor
        {
            get => RectGeometry.ParseColorOrDefault(this.thisHighlightColor, DefaultHighlight);
            set => this.SetColour(ref this.thisHighlightColor, value, DefaultHighlight, nameof(this.HighlightColor));
        }

        // 0..100, as the slider shows it.
        public double HighlightOpacity
        {
            get => RectGeometry.ParseOpacityOrDefault(this.thisHighlightOpacity, DefaultOpacity) * 100.0;
            set => this.SetOpacity(ref this.thisHighlightOpacity, value, nameof(this.HighlightOpacity), nameof(this.HighlightFraction), nameof(this.HighlightOpacityText));
        }

        public double HighlightFraction => this.HighlightOpacity / 100.0;

        public string HighlightOpacityText => PictureColours.OpacityText(this.HighlightFraction);

        public Color ThumbnailColor
        {
            get => RectGeometry.ParseColorOrDefault(this.thisThumbnailColor, DefaultHighlight);
            set => this.SetColour(ref this.thisThumbnailColor, value, DefaultHighlight, nameof(this.ThumbnailColor));
        }

        public double ThumbnailOpacity
        {
            get => RectGeometry.ParseOpacityOrDefault(this.thisThumbnailOpacity, DefaultOpacity) * 100.0;
            set => this.SetOpacity(ref this.thisThumbnailOpacity, value, nameof(this.ThumbnailOpacity), nameof(this.ThumbnailOpacityText));
        }

        public string ThumbnailOpacityText => PictureColours.OpacityText(this.ThumbnailOpacity / 100.0);

        public Color OppositeTraceColor
        {
            get => RectGeometry.ParseColorOrDefault(this.thisOppositeTraceColor, DefaultOppositeTrace);
            set => this.SetColour(ref this.thisOppositeTraceColor, value, DefaultOppositeTrace, nameof(this.OppositeTraceColor));
        }

        // Whether a property change is one of the colours - what goes into the draft.
        public static bool IsColourChange(string? property) => property is
            nameof(HighlightColor) or nameof(HighlightOpacity) or nameof(ThumbnailColor) or nameof(ThumbnailOpacity) or nameof(OppositeTraceColor);

        // "Align KiCad" works on a board with KiCad data, for a picture that names its KiCad view.
        public bool CanAlignKiCad { get; init; }

        public string AlignTip => this.CanAlignKiCad ? PictureToolWording.AlignTip : PictureToolWording.AlignNeedsKiCadTip;

        // "Mark all automatically" needs that, and KiCad aligned on the picture - whether its view is a
        // side of the board, the host finds out and says (KiCadFootprintMarks.Plan).
        public bool CanMarkAutomatically { get; init; }

        public string MarkAutomaticallyTip => this.CanMarkAutomatically
            ? PictureToolWording.MarkAutomaticallyTip
            : PictureToolWording.MarkAutomaticallyNeedsAlignTip;

        public static ComponentPaneBoardPictureRow From(BoardSchematicEntry row, IReadOnlyList<Rect> highlights, bool boardHasKiCad = false, bool aligned = false)
        {
            bool canAlign = boardHasKiCad && !string.IsNullOrWhiteSpace(row.CadName);

            var shown = new ComponentPaneBoardPictureRow
            {
                Name = row.SchematicName ?? string.Empty,
                File = row.SchematicImageFile ?? string.Empty,
                Highlights = highlights,
                CanAlignKiCad = canAlign,
                CanMarkAutomatically = canAlign && aligned,
            };

            shown.Use(row.ReadColours());
            return shown;
        }

        // The colours as the board keeps them.
        public PictureColours ToColours() => new(
            this.thisHighlightColor,
            this.thisHighlightOpacity,
            this.thisOppositeTraceColor,
            this.thisThumbnailColor,
            this.thisThumbnailOpacity);

        // Other colours, every picker told.
        public void Use(PictureColours colours)
        {
            this.thisHighlightColor = colours.HighlightColor?.Trim() ?? string.Empty;
            this.thisHighlightOpacity = colours.HighlightOpacity?.Trim() ?? string.Empty;
            this.thisOppositeTraceColor = colours.OppositeTraceColor?.Trim() ?? string.Empty;
            this.thisThumbnailColor = colours.ThumbnailHighlightColor?.Trim() ?? string.Empty;
            this.thisThumbnailOpacity = colours.ThumbnailHighlightOpacity?.Trim() ?? string.Empty;

            foreach (string property in new[]
            {
                nameof(this.HighlightColor), nameof(this.HighlightOpacity), nameof(this.HighlightFraction), nameof(this.HighlightOpacityText),
                nameof(this.ThumbnailColor), nameof(this.ThumbnailOpacity), nameof(this.ThumbnailOpacityText), nameof(this.OppositeTraceColor),
            })
            {
                this.Raise(property);
            }
        }

        // A colour picked: written as the board keeps colours - unless it is the colour already shown.
        private void SetColour(ref string field, Color value, Color fallback, string name)
        {
            if (RectGeometry.ParseColorOrDefault(field, fallback) == value)
            {
                return;
            }

            field = RectGeometry.ColorText(value);
            this.Raise(name);
        }

        // An opacity picked (0..100): written as a whole percentage - unless it reads as the one shown.
        private void SetOpacity(ref string field, double percent, params string[] names)
        {
            string wanted = PictureColours.OpacityText(percent / 100.0);

            if (string.Equals(PictureColours.OpacityText(RectGeometry.ParseOpacityOrDefault(field, DefaultOpacity)), wanted, StringComparison.Ordinal))
            {
                return;
            }

            field = wanted;

            foreach (string name in names)
            {
                this.Raise(name);
            }
        }
    }
}
