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
    // THE BOARD'S SCHEMATIC IMAGES AND THEIR COLOURS (CRT 4.0 step 2, slice 4b, 2026-10-26 - the plan:
    // "Board > Pictures: a colour picker and opacity slider per picture, with a live preview on the
    // picture, plus 'Use for all pictures'"; "Schematic images" on screen since 2026-10-29). Until now
    // only the table, as text, reached the colours.
    //
    //   - EACH IMAGE is a card, a bordered panel of its own: the image made small with the components
    //     marked on it, drawn in the colours picked as they are picked (PicturePreview), and three of
    //     them - the selected components and how strongly they show, the same on the thumbnail, and a
    //     trace on the board's other side. Since 2026-10-29 the preview zooms, a thumbnail-sized one
    //     under it shows the thumbnail's colours, and either side's traces can be drawn on it - an image
    //     with no KiCad data has its KiCad options greyed out
    //     (ComponentPane.PicturePreviews.cs).
    //   - A CHANGE goes into the draft as the pane's others do - after a pause, ONE step for every
    //     picture's colours (BoardPartEdits.WithPictureColours, by name AND image file). A colour the
    //     picker merely shows - CRT's own default for a blank one - writes nothing (the row's guard).
    //   - "USE FOR ALL PICTURES" gives every picture these colours, as one step.
    //   - ADDING, NAMING, THE KICAD VIEW, ORDER AND REMOVING are the cards' too since step 3 - the
    //     "Schematic images" window went (ComponentPane.BoardPictureList.cs).
    //   - EACH CARD'S TOOLS: "Mark components" and "Align KiCad" open the Schematics tab in that tool
    //     (PictureToolRequested); "Mark all automatically" (step 3, slice 3c), on a picture aligned with
    //     KiCad, has the host draw every component's rectangle from its footprint
    //     (MarkAutomaticallyRequested - TabDrafts.MarkAutomatically.cs), said under the heading.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        private readonly ObservableCollection<ComponentPaneBoardPictureRow> thisBoardPictures = new();
        private readonly List<Task> thisPreviewLoads = new();

        internal IReadOnlyList<ComponentPaneBoardPictureRow> BoardPicturesForTests => this.thisBoardPictures;

        // Every preview read and decoded - for a test.
        internal Task PreviewsLoadedForTests => Task.WhenAll(this.thisPreviewLoads.ToList());

        internal int PreviewLoadsHeldForTests => this.thisPreviewLoads.Count;

        // How many pictures the pane holds decoded - for a test.
        internal int PicturesHeldForTests => this.thisImages.CountForTests;

        // Told each preview decode as it starts, by width, on the pool - false makes it fail. For a test.
        internal Func<int, bool>? PreviewDecodeForTests { get; set; }

        // A read of the previews, or of a row's thumbnail, kept until it is done - the ones done well let
        // go of first (code review, 2026-10-30: every zoom step and tick added one, for as long as the pane
        // was open; 2026-11-02: and every component shown its thumbnails, for as long as the tab lived). One
        // that failed stays, so a test waiting for them all still sees it.
        private void TrackPreviewLoad(Task load)
        {
            this.thisPreviewLoads.RemoveAll(task => task.IsCompletedSuccessfully);
            this.thisPreviewLoads.Add(load);
        }

        private void InitializeBoardPictures()
        {
            this.BoardPicturesItemsControl.ItemsSource = this.thisBoardPictures;

            // The previews never take the settings' room (ContributeWidths.PreviewShown) - followed as the
            // pane is made wider or narrower.
            this.BoardPicturesItemsControl.SizeChanged += (_, e) => this.UseCardsWidth(e.NewSize.Width);
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
                row.TracesTickChanged = null;
            }

            // The cards are made again - the old one's sharp image goes with it.
            this.thisSharpRow = null;

            Dictionary<string, Dictionary<string, List<Rect>>> highlights = HighlightRectBuilder.BuildHighlightRects(board, this.thisRegion);

            // "Align KiCad" needs the board's KiCad data - names under "KiCad data/" in the draft's list;
            // "Mark all automatically" a picture aligned in the draft too.
            bool boardHasKiCad = this.thisSession?.ListedFiles.Keys.Any(SubmissionKiCadFiles.IsInKiCadFolder) == true;
            var alignments = new Dictionary<string, KiCadCalibrationEntry>(StringComparer.OrdinalIgnoreCase);

            foreach (KiCadCalibrationEntry entry in this.thisSession?.Content.Calibrations ?? [])
            {
                alignments.TryAdd(entry.SchematicName?.Trim() ?? string.Empty, entry);
            }

            ComponentPane.Fill(
                this.thisBoardPictures,
                board.Schematics.Select(row =>
                {
                    string name = row.SchematicName?.Trim() ?? string.Empty;
                    highlights.TryGetValue(name, out Dictionary<string, List<Rect>>? byLabel);
                    alignments.TryGetValue(name, out KiCadCalibrationEntry? alignment);

                    return ComponentPaneBoardPictureRow.From(
                        row,
                        byLabel?.Values.SelectMany(rects => rects).ToList() ?? [],
                        boardHasKiCad,
                        alignment is not null,
                        byLabel?.Count ?? 0,
                        byLabel,
                        alignment is null ? null : KiCadFootprintMarks.ValuesOf(alignment));
                }));

            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures)
            {
                row.PropertyChanged += this.OnPictureRowChanged;
                row.CadNameChosen = this.OnCadNameChosen;
                this.UsePreviews(row);
            }

            this.ShowPictureOrder();
            this.ShowCadNameChoices();

            // The count is in the pane's title - the part's one headline.
            this.ShowBoardPartTitle();
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
        // (TabDrafts.MarkAutomaticallyAsync) and says what came of it here (ShowPicturesNote). The
        // picture is its name AND its image file, as its traces are read (code review, 2026-10-30).
        // ###########################################################################################
        internal event EventHandler<(string Picture, string File)>? MarkAutomaticallyRequested;

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
            this.MarkAutomaticallyRequested?.Invoke(this, (row.Name, row.File));
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

        // ###########################################################################################
        // The width the cards' pictures are decoded to for their previews shown whole: enough for the
        // previews as wide as they are now, on this screen (PicturePreviewGeometry.WholeWidth's steps).
        // ###########################################################################################
        private int PreviewDecodeWidth() =>
            PicturePreviewGeometry.WholeWidth(this.thisPreviewWidth, TopLevel.GetTopLevel(this)?.RenderScaling ?? 1.0);

        // ###########################################################################################
        // Every card's preview read that wants it (ComponentPaneBoardPictureRow.WantsPreviewFor): not
        // read yet, or read narrower than the previews now are - after a drag of their width ended past
        // a step, never while it runs - and never a card already being read as wide (code review,
        // 2026-10-30: a card still loading was read twice at once).
        // ###########################################################################################
        private void LoadPreviews()
        {
            // Not while files copied in are still on their way into the draft (ComponentPane.Files.cs,
            // thisFilesArriving) - read once they are there.
            if (this.thisSession is not { } session || this.thisFilesArriving || this.thisBoardPart != PaneBoardPart.Pictures)
            {
                return;
            }

            string dataRoot = DataManager.DataRoot;
            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);
            string identity = session.ExcelDataFile;
            int step = this.PreviewDecodeWidth();

            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures.Where(row => row.WantsPreviewFor(step)))
            {
                this.TrackPreviewLoad(this.LoadPreviewAsync(row, dataRoot, draftFolder, identity, step));
            }
        }

        // ###########################################################################################
        // A picture's preview, read once while the pane is open (PaneImageCache - code review,
        // 2026-10-26: every re-read of the part decoded every picture again), with the picture's own
        // pixel size read from its header for every type CRT shows (ImagePixelSize). Read for `step`
        // (PicturePreviewGeometry.WholeWidth), never wider than the picture itself; once kept, the
        // narrower reads of the same picture are let go of - one read held per picture (code review,
        // 2026-10-30). A read that fails blanks nothing: the card keeps what it shows, and a later ask
        // may read again.
        // ###########################################################################################
        private async Task LoadPreviewAsync(ComponentPaneBoardPictureRow row, string dataRoot, string draftFolder, string identity, int step)
        {
            string file = row.File;
            int generation = this.thisImages.Generation;
            row.ThumbnailAskedFor = step;

            (string? path, string? key, int width, Size pixels, List<string> narrower, List<string> wider) = await Task.Run(() =>
            {
                string? found = DraftFileResolver.Resolve(dataRoot, draftFolder, file, identity);

                if (found is null || !ContributionPackaging.IsDisplayableImageFile(found))
                {
                    return (found, (string?)null, 0, default(Size), new List<string>(), new List<string>());
                }

                Size size = ImagePixelSize.TryRead(found) is { } read ? new Size(read.Width, read.Height) : default;
                int decodeWidth = PicturePreviewGeometry.NoWiderThanThePicture(step, size.Width);

                List<string> KeysOf(Func<int, bool> which) => PicturePreviewGeometry.WholeWidths
                    .Select(other => PicturePreviewGeometry.NoWiderThanThePicture(other, size.Width))
                    .Where(which)
                    .Distinct()
                    .Select(other => PaneImageCache.KeyOf(found, other))
                    .OfType<string>()
                    .ToList();

                return (found, PaneImageCache.KeyOf(found, decodeWidth), decodeWidth, size,
                    KeysOf(other => other < decodeWidth), KeysOf(other => other > decodeWidth));
            });

            // Not there, or no picture CRT shows: asked again later, as for one that does not decode (code
            // review, 2026-11-02: left marked as asked, the card stayed blank until it was made again).
            if (path is null || key is null)
            {
                ComponentPane.TakeBackPreviewAsk(row, step);
                return;
            }

            if (!this.thisImages.TryGet(key, out PaneImage image) && !this.TryGetWiderRead(wider, out image))
            {
                Func<int, bool>? decodeForTests = this.PreviewDecodeForTests;
                image = await Task.Run(() => decodeForTests?.Invoke(width) == false ? PaneImage.None : ComponentPane.DecodePreview(path, width, pixels));

                if (image.Bitmap is not null)
                {
                    // A wider read of the picture that ended first is the one held - one read held per
                    // picture (code review, 2026-10-30: this one was kept beside it until the pane closed).
                    // Not kept, it is still shown below while the card has nothing.
                    bool widerHeld = wider.Any(other => this.thisImages.TryGet(other, out PaneImage held) && held.Bitmap is not null);

                    if (!widerHeld && this.thisImages.Keep(key, image, generation))
                    {
                        this.thisImages.Forget(narrower);
                    }
                }
                else if (!narrower.Any(other => this.thisImages.TryGet(other, out PaneImage held) && held.Bitmap is not null))
                {
                    // A picture that does not decode at all is kept as such - not read again at every
                    // re-read. A wider read failing where a narrower one worked may be asked again.
                    this.thisImages.Keep(key, image, generation);
                }
            }

            if (generation != this.thisImages.Generation || !string.Equals(row.File, file, StringComparison.Ordinal))
            {
                return;
            }

            if (image.Bitmap is null)
            {
                // Asked again later - the card shows what it showed.
                ComponentPane.TakeBackPreviewAsk(row, step);
                return;
            }

            // A wider read asked for meanwhile is the one the card takes - this one only while the card
            // shows nothing yet.
            if (row.ThumbnailAskedFor != step && row.Thumbnail is not null)
            {
                return;
            }

            row.ImageSize = image.OriginalSize;
            row.Thumbnail = image.Bitmap;
            row.ThumbnailDecodedFor = step;
        }

        // The read asked for `step` brought nothing - unless a wider one was asked for meanwhile.
        private static void TakeBackPreviewAsk(ComponentPaneBoardPictureRow row, int step)
        {
            if (row.ThumbnailAskedFor == step)
            {
                row.ThumbnailAskedFor = row.ThumbnailDecodedFor;
            }
        }

        // ###########################################################################################
        // A wider read of the picture held already serves narrower previews (code review, 2026-11-01): it
        // stays the one read held once they are dragged narrow, and looking only for the narrow width,
        // every card made again - an undo, an add, a move - decoded its picture from disk and then
        // declined to keep it.
        // ###########################################################################################
        private bool TryGetWiderRead(IEnumerable<string> wider, out PaneImage image)
        {
            foreach (string other in wider)
            {
                if (this.thisImages.TryGet(other, out image) && image.Bitmap is not null)
                {
                    return true;
                }
            }

            image = PaneImage.None;
            return false;
        }

        private static PaneImage DecodePreview(string path, int width, Size pixels)
        {
            try
            {
                using FileStream stream = File.OpenRead(path);
                return new PaneImage(Bitmap.DecodeToWidth(stream, width, BitmapInterpolationMode.MediumQuality), pixels);
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
    // One entry in a picture's "Associate KiCad file" drop-down. An object of its own rather than the
    // CRT.Data record, so the entry a drop-down shows keeps its identity while the entries around it
    // come and go (ComponentPaneBoardPictureRow.ApplyCadNameChoices).
    // ###########################################################################################
    public sealed class CadNameChoiceItem
    {
        public CadNameChoiceItem(SchematicCadNameChoice choice)
        {
            this.CadName = choice.CadName;
            this.Label = choice.Label;
        }

        public string CadName { get; }

        public string Label { get; }

        public override string ToString() => this.Label;
    }

    // ###########################################################################################
    // One of the board's pictures as the Pictures part holds it - its colours as the board keeps them
    // (text), shown to the pickers as colours and percentages. A blank colour shows CRT's own default,
    // and a picker writing back the very colour it shows is no change (the setters' guard): a picker
    // that writes its value as it is built would otherwise turn every blank colour into a written one.
    // ###########################################################################################
    public sealed class ComponentPaneBoardPictureRow : ComponentPaneRowBase
    {
        // What CRT draws with when a colour or an opacity is blank (the Schematics tab's defaults; the
        // traces' are TraceColours').
        private static readonly Color DefaultHighlight = Colors.IndianRed;
        private const double DefaultOpacity = 0.20;

        private string thisHighlightColor = string.Empty;
        private string thisHighlightOpacity = string.Empty;
        private string thisOppositeTraceColor = string.Empty;
        private string thisThumbnailColor = string.Empty;
        private string thisThumbnailOpacity = string.Empty;
        private string thisTraceColor = string.Empty;
        private string thisTraceOpacity = string.Empty;
        private string thisOppositeTraceOpacity = string.Empty;
        private Bitmap? thisThumbnail;
        private Bitmap? thisSharpThumbnail;
        private Size thisImageSize;
        private double thisPreviewWidth = ContributeWidths.DefaultPreviewWidth;
        private Action<int>? thisWantSharper;
        private PreviewTraces? thisTraces;
        private bool thisShowThisSideTraces;
        private bool thisShowOtherSideTraces;
        private string thisTracesNote = string.Empty;
        private string thisName = string.Empty;
        private string? thisNameDraft;
        private string thisCadName = string.Empty;
        private bool thisBoardHasKiCad;
        private bool thisAligned;
        private bool thisCanMoveUp;
        private bool thisCanMoveDown;
        private CadNameChoiceItem? thisSelectedCadNameChoice;

        // The picture's name as the draft has it - set by a rename the draft took (step 3, moved inline).
        public string Name
        {
            get => this.thisName;
            set
            {
                if (this.Set(ref this.thisName, value))
                {
                    this.Raise(nameof(this.NameDraft));
                }
            }
        }

        // What the name box holds - the name until something else is typed.
        public string NameDraft
        {
            get => this.thisNameDraft ?? this.thisName;
            set
            {
                if (string.Equals(this.NameDraft, value, StringComparison.Ordinal))
                {
                    return;
                }

                this.thisNameDraft = value;
                this.Raise(nameof(this.NameDraft));
            }
        }

        // The box holds another name than the picture has - not saved yet.
        internal bool HasNameEdit => !string.Equals(this.NameDraft, this.thisName, StringComparison.Ordinal);

        // The typed name let go - the box shows the picture's name again.
        internal void ForgetNameEdit()
        {
            this.thisNameDraft = null;
            this.Raise(nameof(this.NameDraft));
        }

        public string File { get; init; } = string.Empty;

        // The KiCad view the picture shows - its CAD name.
        public string CadName
        {
            get => this.thisCadName;
            set
            {
                if (this.Set(ref this.thisCadName, value?.Trim()))
                {
                    this.RaiseTools();
                }
            }
        }

        // ###########################################################################################
        // The "Associate KiCad file" drop-down - the "Schematic images" window's, moved inline. ONE list for
        // the row's lifetime, changed in place (ApplyCadNameChoices): a ComboBox whose items are
        // replaced lets go of its selection, and the row would hear "nothing chosen".
        // ###########################################################################################
        public ObservableCollection<CadNameChoiceItem> CadNameChoices { get; } = new();

        // Told when another view is picked - the pane writes it. Never for the drop-down letting go of
        // its items (null).
        internal Action<ComponentPaneBoardPictureRow, string>? CadNameChosen { get; set; }

        public CadNameChoiceItem? SelectedCadNameChoice
        {
            get => this.thisSelectedCadNameChoice;
            set
            {
                if (value is null || ReferenceEquals(value, this.thisSelectedCadNameChoice))
                {
                    return;
                }

                this.thisSelectedCadNameChoice = value;
                this.Raise(nameof(this.SelectedCadNameChoice));

                if (!string.Equals(value.CadName, this.thisCadName, StringComparison.OrdinalIgnoreCase))
                {
                    this.CadNameChosen?.Invoke(this, value.CadName);
                }
            }
        }

        // The drop-down shows once there is a view to pick, or the picture names one already.
        public bool ShowsCadNames => this.CadNameChoices.Count > 1 || this.thisCadName.Length > 0;

        // ###########################################################################################
        // Makes CadNameChoices read these choices, in order, IN PLACE (InPlaceList.Sync) - never the
        // selected entry taken away, which the choices always hold (SchematicCadNameChoices lists the
        // picture's own name) - then selects the one CadName names.
        // ###########################################################################################
        internal void ApplyCadNameChoices(IReadOnlyList<SchematicCadNameChoice> choices)
        {
            InPlaceList.Sync(
                this.CadNameChoices,
                choices,
                (item, choice) => string.Equals(item.CadName, choice.CadName, StringComparison.OrdinalIgnoreCase),
                choice => new CadNameChoiceItem(choice),
                keep: this.thisSelectedCadNameChoice);

            this.thisSelectedCadNameChoice = choices.Count == 0
                ? null
                : this.CadNameChoices.FirstOrDefault(item => string.Equals(
                    item.CadName,
                    SchematicCadNameChoices.Selected(choices, this.thisCadName).CadName,
                    StringComparison.OrdinalIgnoreCase));

            // Always said, so a drop-down that lost its selection while its items moved shows it again.
            this.Raise(nameof(this.SelectedCadNameChoice));
            this.Raise(nameof(this.ShowsCadNames));
        }

        // Whether the picture can move up or down - not at an end.
        public bool CanMoveUp
        {
            get => this.thisCanMoveUp;
            set
            {
                this.thisCanMoveUp = value;
                this.Raise(nameof(this.CanMoveUp));
            }
        }

        public bool CanMoveDown
        {
            get => this.thisCanMoveDown;
            set
            {
                this.thisCanMoveDown = value;
                this.Raise(nameof(this.CanMoveDown));
            }
        }

        // The picture's highlights, in its own pixels.
        public IReadOnlyList<Rect> Highlights { get; init; } = [];

        public Bitmap? Thumbnail
        {
            get => this.thisThumbnail;
            set
            {
                this.thisThumbnail = value;
                this.Raise(nameof(this.Thumbnail));
                this.Raise(nameof(this.ShownThumbnail));
            }
        }

        // The step Thumbnail was read for (PicturePreviewGeometry.WholeWidth) - 0 before one is shown -
        // and the step last asked, being read or done.
        internal int ThumbnailDecodedFor { get; set; }

        internal int ThumbnailAskedFor { get; set; }

        // ###########################################################################################
        // Whether its previews want reading for `step`: never while a read as wide is asked; not read
        // yet, or read narrower than the step AND narrower than the picture itself - a picture already
        // shown at its own width gains nothing from a wider step.
        // ###########################################################################################
        internal bool WantsPreviewFor(int step) =>
            this.File.Length > 0 &&
            this.ThumbnailAskedFor < step &&
            (this.thisThumbnail is not { } shown ||
             (this.ThumbnailDecodedFor < step && shown.PixelSize.Width < PicturePreviewGeometry.NoWiderThanThePicture(step, this.thisImageSize.Width)));

        // ---- the previews (2026-10-29 - ComponentPane.PicturePreviews.cs) -----------------------------

        // The image decoded sharper for the zoomed preview - null until asked for, and again once
        // another card's is (one at a time).
        public Bitmap? SharpThumbnail
        {
            get => this.thisSharpThumbnail;
            set
            {
                this.thisSharpThumbnail = value;
                this.Raise(nameof(this.SharpThumbnail));
                this.Raise(nameof(this.ShownThumbnail));
            }
        }

        // What the zoomable preview draws - the sharper image while there is one.
        public Bitmap? ShownThumbnail => this.thisSharpThumbnail ?? this.thisThumbnail;

        // Told when the zoomed preview wants the image sharper - the pane decodes it.
        public Action<int>? WantSharper
        {
            get => this.thisWantSharper;
            set
            {
                this.thisWantSharper = value;
                this.Raise(nameof(this.WantSharper));
            }
        }

        // The previews' width - the same on every card, dragged at any card's handle (the pane sets it).
        public double PreviewWidth
        {
            get => this.thisPreviewWidth;
            set
            {
                if (this.thisPreviewWidth.Equals(value))
                {
                    return;
                }

                this.thisPreviewWidth = value;
                this.Raise(nameof(this.PreviewWidth));
                this.Raise(nameof(this.PreviewHeight));
            }
        }

        // The zoomable preview's height - 10 to 16, as it was drawn at 320 x 200.
        public double PreviewHeight => Math.Round(this.thisPreviewWidth * 0.625);

        // ###########################################################################################
        // "Traces on the preview": this side's traces and the other side's, each its own tick, both
        // drawn when both are ticked (owner request, 2026-10-29: "so we can have both in preview").
        // ###########################################################################################
        public bool ShowThisSideTraces
        {
            get => this.thisShowThisSideTraces;
            set => this.SetTicks(value, this.thisShowOtherSideTraces, tell: true);
        }

        public bool ShowOtherSideTraces
        {
            get => this.thisShowOtherSideTraces;
            set => this.SetTicks(this.thisShowThisSideTraces, value, tell: true);
        }

        internal bool ShowsAnyTraces => this.thisShowThisSideTraces || this.thisShowOtherSideTraces;

        // Told when a tick changes - the pane reads the traces when they are not read yet.
        internal Action<ComponentPaneBoardPictureRow>? TracesTickChanged { get; set; }

        // ###########################################################################################
        // The ticks as they are to be - `tell` the pane, or not when it sets them itself (a card made
        // again ticked as before, reading its traces once for both).
        // ###########################################################################################
        internal void SetTicks(bool thisSide, bool otherSide, bool tell)
        {
            if (thisSide == this.thisShowThisSideTraces && otherSide == this.thisShowOtherSideTraces)
            {
                return;
            }

            this.thisShowThisSideTraces = thisSide;
            this.thisShowOtherSideTraces = otherSide;

            foreach (string property in new[]
            {
                nameof(this.ShowThisSideTraces), nameof(this.ShowOtherSideTraces), nameof(this.ShownThisSideTraces), nameof(this.ShownOtherSideTraces),
            })
            {
                this.Raise(property);
            }

            if (tell)
            {
                this.TracesTickChanged?.Invoke(this);
            }
        }

        // Why the traces could not be read for the image - null when they were (or are not read yet).
        internal string? TracesProblem { get; set; }

        // True while its traces are being read - a second tick reads nothing more.
        internal bool TracesReading { get; set; }

        // The latest read of its traces asked for - only that one answers (ComponentPane.LoadTracesAsync).
        internal int TracesAsk { get; set; }

        // ###########################################################################################
        // The traces read let go of - another KiCad view picked: what was read, and why none could be,
        // were of the old one. A read still running answers nothing any more (TracesAsk moves on).
        // ###########################################################################################
        internal void ForgetTraces()
        {
            this.TracesAsk++;
            this.TracesReading = false;
            this.TracesProblem = null;
            this.Traces = null;
            this.TracesNote = string.Empty;
        }

        // A sharper image being decoded for the zoomed preview, and the widest asked for meanwhile -
        // decoded once that one ends (ComponentPane.LoadSharperAsync).
        internal bool SharpDecoding { get; set; }

        internal int SharpWantedNext { get; set; }

        // Both sides' traces, in the image's pixels - null until read.
        public PreviewTraces? Traces
        {
            get => this.thisTraces;
            set
            {
                this.thisTraces = value;
                this.Raise(nameof(this.Traces));
                this.Raise(nameof(this.ShownThisSideTraces));
                this.Raise(nameof(this.ShownOtherSideTraces));
            }
        }

        public IReadOnlyList<PreviewTrace>? ShownThisSideTraces => this.thisShowThisSideTraces ? this.thisTraces?.ThisSide : null;

        public IReadOnlyList<PreviewTrace>? ShownOtherSideTraces => this.thisShowOtherSideTraces ? this.thisTraces?.OtherSide : null;

        // ###########################################################################################
        // EACH SIDE'S TRACES - their colour and opacity (owner request, 2026-10-30) - drawn by the
        // Schematics tab's own rule (TraceColours.Resolve): blank, this side's colour follows the
        // image's highlight colour and both opacities its highlight opacity plus a quarter, as before.
        // A picker writing back what it shows writes nothing, as every colour of the card.
        //
        // A SCHEMATIC SHEET HAS ITS OWN BLANK, ORANGE (code review, 2026-11-04: the card said DeepSkyBlue of what the
        // tab drew Orange) - told from a side of the board once the views are read (ShowsBoardSide, RaiseTools).
        // ###########################################################################################
        private TraceColours ResolvedTraces => this.TracesOf(this.ToColours());

        // The traces' colours and opacities as they were before a change - what the pickers read back
        // as no change, the derived values of blank columns included.
        private TraceColours ResolvedTracesWithout(Func<PictureColours, PictureColours> blanked) => this.TracesOf(blanked(this.ToColours()));

        private TraceColours TracesOf(PictureColours colours) =>
            TraceColours.Resolve(new BoardSchematicEntry().WithColours(colours), this.ShowsBoardSide == false ? TraceColours.SheetBlankHighlight : null);

        public Color ThisSideTraceColor
        {
            get => this.ResolvedTraces.ThisSide;
            set => this.SetColour(ref this.thisTraceColor, value, this.ResolvedTracesWithout(colours => colours with { TraceColor = string.Empty }).ThisSide, nameof(this.ThisSideTraceColor));
        }

        // 0..100, as the slider shows it.
        public double ThisSideTraceOpacity
        {
            get => this.ResolvedTraces.ThisSideOpacity * 100.0;
            set => this.SetOpacity(
                ref this.thisTraceOpacity,
                value,
                this.ResolvedTracesWithout(colours => colours with { TraceOpacity = string.Empty }).ThisSideOpacity,
                nameof(this.ThisSideTraceOpacity), nameof(this.ThisSideTraceFraction), nameof(this.ThisSideTraceOpacityText));
        }

        public double ThisSideTraceFraction => this.ResolvedTraces.ThisSideOpacity;

        public string ThisSideTraceOpacityText => PictureColours.OpacityText(this.ThisSideTraceFraction);

        public double OppositeTraceOpacity
        {
            get => this.ResolvedTraces.OtherSideOpacity * 100.0;
            set => this.SetOpacity(
                ref this.thisOppositeTraceOpacity,
                value,
                this.ResolvedTracesWithout(colours => colours with { OppositeTraceOpacity = string.Empty }).OtherSideOpacity,
                nameof(this.OppositeTraceOpacity), nameof(this.OppositeTraceFraction), nameof(this.OppositeTraceOpacityText));
        }

        public double OppositeTraceFraction => this.ResolvedTraces.OtherSideOpacity;

        public string OppositeTraceOpacityText => PictureColours.OpacityText(this.OppositeTraceFraction);

        // What the traces' line says - being read, none there, or why they cannot be shown.
        public string TracesNote
        {
            get => this.thisTracesNote;
            set
            {
                if (this.Set(ref this.thisTracesNote, value))
                {
                    this.Raise(nameof(this.HasTracesNote));
                }
            }
        }

        public bool HasTracesNote => this.thisTracesNote.Length > 0;

        // The traces can be drawn on an image "Mark all automatically" works on - whether its view is
        // a side of the board, reading them finds out and says (KiCadFootprintMarks.Plan).
        public bool CanShowTraces => this.CanMarkAutomatically;

        public string ThisSideTracesTip => this.CanShowTraces ? ComponentPaneWording.ThisSideTracesTip : ComponentPaneWording.ShowTracesNeedsAlignTip;

        public string OtherSideTracesTip => this.CanShowTraces ? ComponentPaneWording.OtherSideTracesTip : ComponentPaneWording.ShowTracesNeedsAlignTip;

        // ###########################################################################################
        // The image has KiCad data - the board's KiCad project, and a KiCad view picked for it - so
        // its KiCad options mean something (owner request, 2026-10-29: "if the schematic image has no
        // KiCad data, then the KiCad (traces) options should be disabled"): the trace colour, and the
        // ticks, which need it aligned too (CanShowTraces).
        // ###########################################################################################
        public bool HasKiCadData => this.CanAlignKiCad;

        public string TraceColourTip => this.HasKiCadData ? ComponentPaneWording.TraceColourTip : ComponentPaneWording.TraceColourNeedsKiCadTip;

        public string ThisSideTraceColourTip => this.HasKiCadData ? ComponentPaneWording.ThisSideTraceColourTip : ComponentPaneWording.TraceColourNeedsKiCadTip;

        public string ThisSideTraceOpacityTip => this.HasKiCadData ? ComponentPaneWording.ThisSideTraceOpacityTip : ComponentPaneWording.TraceColourNeedsKiCadTip;

        public string OppositeTraceOpacityTip => this.HasKiCadData ? ComponentPaneWording.OtherSideTraceOpacityTip : ComponentPaneWording.TraceColourNeedsKiCadTip;

        // The thumbnail's opacity, 0..1, as its preview draws it.
        public double ThumbnailFraction => this.ThumbnailOpacity / 100.0;

        public Size ImageSize
        {
            get => this.thisImageSize;
            set
            {
                this.thisImageSize = value;
                this.Raise(nameof(this.ImageSize));
                this.RaiseCounts();
            }
        }

        public Color HighlightColor
        {
            get => RectGeometry.ParseColorOrDefault(this.thisHighlightColor, DefaultHighlight);
            set => this.SetColour(ref this.thisHighlightColor, value, DefaultHighlight, nameof(this.HighlightColor), nameof(this.ThisSideTraceColor));
        }

        // 0..100, as the slider shows it. The traces' opacities follow it while they are blank.
        public double HighlightOpacity
        {
            get => RectGeometry.ParseOpacityOrDefault(this.thisHighlightOpacity, DefaultOpacity) * 100.0;
            set => this.SetOpacity(
                ref this.thisHighlightOpacity,
                value,
                DefaultOpacity,
                nameof(this.HighlightOpacity), nameof(this.HighlightFraction), nameof(this.HighlightOpacityText),
                nameof(this.ThisSideTraceOpacity), nameof(this.ThisSideTraceFraction), nameof(this.ThisSideTraceOpacityText),
                nameof(this.OppositeTraceOpacity), nameof(this.OppositeTraceFraction), nameof(this.OppositeTraceOpacityText));
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
            set => this.SetOpacity(ref this.thisThumbnailOpacity, value, DefaultOpacity, nameof(this.ThumbnailOpacity), nameof(this.ThumbnailOpacityText), nameof(this.ThumbnailFraction));
        }

        public string ThumbnailOpacityText => PictureColours.OpacityText(this.ThumbnailOpacity / 100.0);

        public Color OppositeTraceColor
        {
            get => this.ResolvedTraces.OtherSide;
            set => this.SetColour(ref this.thisOppositeTraceColor, value, TraceColours.DefaultOtherSide, nameof(this.OppositeTraceColor));
        }

        // Whether a property change is one of the colours - what goes into the draft.
        public static bool IsColourChange(string? property) => property is
            nameof(HighlightColor) or nameof(HighlightOpacity) or nameof(ThumbnailColor) or nameof(ThumbnailOpacity) or nameof(OppositeTraceColor) or
            nameof(ThisSideTraceColor) or nameof(ThisSideTraceOpacity) or nameof(OppositeTraceOpacity);

        // "Align KiCad" works on a board with KiCad data, for a picture that names its KiCad view - which
        // the drop-down can change while the card is shown.
        public bool CanAlignKiCad => this.thisBoardHasKiCad && this.thisCadName.Length > 0;

        public string AlignTip => this.CanAlignKiCad ? PictureToolWording.AlignTip : PictureToolWording.AlignNeedsKiCadTip;

        // "Mark all automatically" needs that, and KiCad aligned on the picture - from a side of the board's
        // footprints or a schematic sheet's symbols (KiCadFootprintMarks.Plan).
        public bool CanMarkAutomatically => this.CanAlignKiCad && this.thisAligned;

        public string MarkAutomaticallyTip => this.CanMarkAutomatically
            ? PictureToolWording.MarkAutomaticallyTip
            : PictureToolWording.MarkAutomaticallyNeedsAlignTip;

        // ###########################################################################################
        // The board's KiCad views that are a side of the board (pcb_top, pcb_bottom) - null while not
        // known; set by the pane with the views (ComponentPane.ShowCadNameChoices). Whether this picture
        // shows one of them - null while that is not known.
        // ###########################################################################################
        internal IReadOnlyCollection<string>? BoardSides
        {
            get => this.thisBoardSides;
            set
            {
                if (!ReferenceEquals(this.thisBoardSides, value))
                {
                    this.thisBoardSides = value;
                    this.RaiseTools();
                }
            }
        }

        private IReadOnlyCollection<string>? thisBoardSides;

        // ###########################################################################################
        // HOW FAR THE PICTURE IS, UNDER ITS THUMBNAIL (owner requests, 2026-11-01: "if it has KiCad data and
        // the image has been mapped to a KiCad file, then per image I want to see status of how many
        // components is available from KiCad and how many has been marked", then in a list, "If any
        // components have been manually marked, then I think it should show like this"): how many
        // components its KiCad view shows (KiCadComponentTable.CountsByView - null while not known, set
        // by the pane with the views) and how many are marked on it, of the region shown; the marks count
        // is the part's as it was read, so a new card follows a change of them.
        //
        // *** A MARK IS AUTOMATIC WHEN IT IS WHERE KiCad PUTS IT (KiCadFootprintMarks.CountAutomatic) - NOTHING
        // IS STORED. *** So one moved by hand counts as manual, and so does every mark on a picture not
        // aligned: nothing says where KiCad would put it.
        // ###########################################################################################
        internal IReadOnlyDictionary<string, int>? ViewComponents
        {
            get => this.thisViewComponents;
            set
            {
                if (!ReferenceEquals(this.thisViewComponents, value))
                {
                    this.thisViewComponents = value;
                    this.RaiseCounts();
                }
            }
        }

        private IReadOnlyDictionary<string, int>? thisViewComponents;

        // Where each view has its components, a side of the board's footprints and a sheet's symbols - null
        // while not known; set by the pane with the views.
        internal IReadOnlyDictionary<string, KiCadViewFootprints>? ViewFootprints
        {
            get => this.thisViewFootprints;
            set
            {
                if (!ReferenceEquals(this.thisViewFootprints, value))
                {
                    this.thisViewFootprints = value;
                    this.RaiseCounts();
                }
            }
        }

        private IReadOnlyDictionary<string, KiCadViewFootprints>? thisViewFootprints;

        public int MarkedComponents { get; init; }

        // The picture's marks by label, in its own pixels, and its alignment in the draft - null for none.
        internal IReadOnlyDictionary<string, List<Rect>>? MarksByLabel { get; init; }

        internal KiCadCalibrationValues? Alignment { get; init; }

        private int? ComponentsFromKiCad =>
            this.HasKiCadData && this.thisViewComponents is { } counts && counts.TryGetValue(this.thisCadName, out int count) ? count : null;

        public bool ShowsKiCadCounts => this.ComponentsFromKiCad is not null;

        private int AutomaticallyMarked =>
            this.Alignment is { } alignment &&
            this.MarksByLabel is { } marks &&
            this.thisViewFootprints is { } views &&
            views.TryGetValue(this.thisCadName, out KiCadViewFootprints? footprints)
                ? KiCadFootprintMarks.CountAutomatic(marks, footprints, alignment, (int)this.thisImageSize.Width, (int)this.thisImageSize.Height)
                : 0;

        private int ManuallyMarked => Math.Max(0, this.MarkedComponents - this.AutomaticallyMarked);

        public bool ShowsManuallyMarked => this.ManuallyMarked > 0;

        public bool ShowsAutomaticallyMarked => this.ShowsManuallyMarked;

        public string ComponentsAvailableText =>
            this.ComponentsFromKiCad is { } count ? ComponentPaneWording.PictureComponentsAvailable(count) : string.Empty;

        public string ComponentsAutomaticallyMarkedText => ComponentPaneWording.PictureComponentsAutomaticallyMarked(this.AutomaticallyMarked);

        public string ComponentsManuallyMarkedText => ComponentPaneWording.PictureComponentsManuallyMarked(this.ManuallyMarked);

        public string MarkedComponentsNumber => this.MarkedComponents.ToString(System.Globalization.CultureInfo.InvariantCulture);

        public string MarkedComponentsWords => ComponentPaneWording.PictureComponentsMarkedWords(this.MarkedComponents);

        private void RaiseCounts()
        {
            foreach (string property in new[]
            {
                nameof(this.ShowsKiCadCounts), nameof(this.ComponentsAvailableText), nameof(this.ComponentsAutomaticallyMarkedText),
                nameof(this.ComponentsManuallyMarkedText), nameof(this.ShowsAutomaticallyMarked), nameof(this.ShowsManuallyMarked),
            })
            {
                this.Raise(property);
            }
        }

        private bool? ShowsBoardSide => this.thisBoardSides is null || this.thisCadName.Length == 0
            ? null
            : this.thisBoardSides.Contains(this.thisCadName, StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // WHAT COMES NEXT ON A PICTURE WITH A KiCad VIEW NOT ALIGNED YET (owner request, 2026-11-01: "it is
        // not clear that I must first do the Align KiCad, before I can do the Mark all automatically from
        // KiCad") - said on the card, not only in a greyed-out button's tooltip: align it, and its Done
        // marks every component (PictureToolWording.AlignFirst).
        // ###########################################################################################
        public bool AlignsFirst => this.CanAlignKiCad && !this.thisAligned;

        public string AlignFirstText => PictureToolWording.AlignFirst;

        // "Mark all automatically" is offered on every picture with a KiCad view (hidden otherwise, not
        // greyed out - owner request, 2026-10-30: "less disabled elements to relate to").
        public bool ShowsMarkAutomatically => this.HasKiCadData;

        private void RaiseTools()
        {
            foreach (string property in new[]
            {
                nameof(this.AlignsFirst), nameof(this.ShowsMarkAutomatically),
                nameof(this.CanAlignKiCad), nameof(this.AlignTip), nameof(this.CanMarkAutomatically), nameof(this.MarkAutomaticallyTip),
                nameof(this.ShowsCadNames), nameof(this.CanShowTraces), nameof(this.ThisSideTracesTip), nameof(this.OtherSideTracesTip),
                nameof(this.HasKiCadData), nameof(this.TraceColourTip), nameof(this.ThisSideTraceColourTip),
                nameof(this.ThisSideTraceOpacityTip), nameof(this.OppositeTraceOpacityTip), nameof(this.ThisSideTraceColor),
            })
            {
                this.Raise(property);
            }

            this.RaiseCounts();
        }

        public static ComponentPaneBoardPictureRow From(
            BoardSchematicEntry row,
            IReadOnlyList<Rect> highlights,
            bool boardHasKiCad = false,
            bool aligned = false,
            int markedComponents = 0,
            IReadOnlyDictionary<string, List<Rect>>? marksByLabel = null,
            KiCadCalibrationValues? alignment = null)
        {
            var shown = new ComponentPaneBoardPictureRow
            {
                File = row.SchematicImageFile ?? string.Empty,
                Highlights = highlights,
                MarkedComponents = markedComponents,
                MarksByLabel = marksByLabel,
                Alignment = alignment,
                thisName = row.SchematicName ?? string.Empty,
                thisCadName = row.CadName?.Trim() ?? string.Empty,
                thisBoardHasKiCad = boardHasKiCad,
                thisAligned = aligned,
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
            this.thisThumbnailOpacity,
            this.thisTraceColor,
            this.thisTraceOpacity,
            this.thisOppositeTraceOpacity);

        // Other colours, every picker told.
        public void Use(PictureColours colours)
        {
            this.thisHighlightColor = colours.HighlightColor?.Trim() ?? string.Empty;
            this.thisHighlightOpacity = colours.HighlightOpacity?.Trim() ?? string.Empty;
            this.thisOppositeTraceColor = colours.OppositeTraceColor?.Trim() ?? string.Empty;
            this.thisThumbnailColor = colours.ThumbnailHighlightColor?.Trim() ?? string.Empty;
            this.thisThumbnailOpacity = colours.ThumbnailHighlightOpacity?.Trim() ?? string.Empty;
            this.thisTraceColor = colours.TraceColor?.Trim() ?? string.Empty;
            this.thisTraceOpacity = colours.TraceOpacity?.Trim() ?? string.Empty;
            this.thisOppositeTraceOpacity = colours.OppositeTraceOpacity?.Trim() ?? string.Empty;

            foreach (string property in new[]
            {
                nameof(this.HighlightColor), nameof(this.HighlightOpacity), nameof(this.HighlightFraction), nameof(this.HighlightOpacityText),
                nameof(this.ThumbnailColor), nameof(this.ThumbnailOpacity), nameof(this.ThumbnailOpacityText), nameof(this.OppositeTraceColor),
                nameof(this.ThumbnailFraction), nameof(this.ThisSideTraceColor),
                nameof(this.ThisSideTraceOpacity), nameof(this.ThisSideTraceFraction), nameof(this.ThisSideTraceOpacityText),
                nameof(this.OppositeTraceOpacity), nameof(this.OppositeTraceFraction), nameof(this.OppositeTraceOpacityText),
            })
            {
                this.Raise(property);
            }
        }

        // A colour picked: written as the board keeps colours - unless it is the colour already shown.
        private void SetColour(ref string field, Color value, Color fallback, params string[] names)
        {
            if (RectGeometry.ParseColorOrDefault(field, fallback) == value)
            {
                return;
            }

            field = RectGeometry.ColorText(value);

            foreach (string name in names)
            {
                this.Raise(name);
            }
        }

        // An opacity picked (0..100): written as a whole percentage - unless it reads as the one shown,
        // `fallback` being what a blank field shows.
        private void SetOpacity(ref string field, double percent, double fallback, params string[] names)
        {
            string wanted = PictureColours.OpacityText(percent / 100.0);

            if (string.Equals(PictureColours.OpacityText(RectGeometry.ParseOpacityOrDefault(field, fallback)), wanted, StringComparison.Ordinal))
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
