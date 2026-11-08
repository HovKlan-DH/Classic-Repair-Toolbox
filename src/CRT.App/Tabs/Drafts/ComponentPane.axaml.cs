using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT PANE (CRT 4.0 step 2, 2026-10-25 - the plan:
    //
    //     | IC1  Ferranti ULA 6C001E-7              [Changed] |
    //     | Basics                                            |
    //     |   Short description [one line ...........] 42/100 |
    //     | Pinout     DIP-40 [drawn top view]                |
    //     | Replacements (0)                 [Add replacement]|
    //
    // "A component opens in a pane, never in a maximized window ... Every change is saved within a
    // second ... An edit in one shows up in all the others at once").
    //
    // On the right of the Contribute tab, on the draft held in memory (DraftSession) - the same one the
    // table, the header's Undo and every other editor use:
    //
    //   - EVERY CHANGE GOES INTO THE DRAFT AS IT IS MADE: a box when it is left, Enter is pressed, or
    //     typing pauses for a second (TypingPause); a pin, a replacement part or a link when one of
    //     its boxes is left or its drop-down list picked. Each is one step of the draft's history
    //     (ComponentPaneEdits works it out - a change that changes nothing is no step).
    //   - THE DRAFT SAVES ITSELF a second later, off the UI thread (the save timer, as the table's);
    //     the draft's header says Saved, Saving... or "Could not save" with Try again.
    //   - WHAT ANOTHER EDITOR CHANGES SHOWS AT ONCE: before any change not the pane's own, what the
    //     pane holds goes in first (ChangeComing), and after it the pane reads the component again.
    //   - ITS OSCILLOSCOPE BASELINES, PICTURES AND FILES TOO (slice 3b, 2026-10-26): a file added goes into the
    //     draft under a name of the board's own and is saved at once (ComponentPane.Files.cs).
    //   - AND THE COMPONENT AS A WHOLE (slice 3c, 2026-10-26): a new one added, one deleted, "Show on
    //     board" and "Preview popup" (ComponentPane.Whole.cs) - the component editor's window is gone.
    //   - AND, IN PLACE OF A COMPONENT, ONE OF THE BOARD'S OWN LISTS (slice 4a, 2026-10-26): its files,
    //     its links, its credits or its important signals (ComponentPane.Board.cs) - and its schematic
    //     images, added, named, matched to a KiCad view, ordered and deleted on their cards since step 3
    //     (ComponentPane.BoardPictureList.cs - the "Schematic images" window went).
    //
    // The file map:
    //   ComponentPane.axaml.cs  - showing a component, the fields, pins, replacement parts and links,
    //                             committing and saving, Undo changes
    //   ComponentPane.Pages.cs  - a component's pages (the left column picks them), the region shown and
    //                             "Applies to", its label, what Submit would refuse marked where it is, and
    //                             the temporary tool making pinouts of oscilloscope baselines
    //   ComponentPane.RegionsAndCategories.cs - "Board: Regions" and "Board: Categories"
    //   ComponentPane.Files.cs  - oscilloscope baselines (with the new-baseline panel and the capture),
    //                             pictures and files: adding, changing, moving, thumbnails
    //   ComponentPane.Whole.cs  - the new-component form, Delete component and "Bring back", the
    //                             buttons that ask the host to show the component on the board or
    //                             preview its popup
    //   ComponentPane.Board.cs  - the board's files, links, credits and important signals
    //   ComponentPane.BoardPictures.cs - the board's schematic images and the colours a selected
    //                             component is drawn in on each, with their previews
    //   ComponentPane.BoardPictureList.cs - the images themselves: Add schematic image(s), each one's
    //                             name, its KiCad view, Move up/down and Delete
    //   ComponentPane.PicturePreviews.cs - the cards' previews: their width dragged, the image decoded
    //                             sharper while zoomed, and the other side's traces drawn on them
    //   ComponentPane.SignalSuggestions.cs - important signals suggested from the KiCad net names, and
    //                             "Show on board" lighting a net up on the Schematics tab
    //   ComponentPane.KiCadPins.cs - the component's pins from the board's KiCad data: "Fill from
    //                             KiCad" and each pin's net on this board
    //   ComponentPane.Copy.cs   - "Clone component from another board": everything of one picked there;
    //                             "Also on other boards": this draft's change of the part, and the
    //                             files it replaced, onto the other boards' own drafts
    // ###########################################################################################
    public partial class ComponentPane : UserControl
    {
        // How long typing may pause before what was typed goes into the draft.
        internal static readonly TimeSpan TypingPause = TimeSpan.FromSeconds(1);

        // The widest the sections and the board's lists are laid out (the markup's MaxWidth="980").
        internal const double ReadableWidth = 980;

        // The mark the pane's own changes carry in the draft's history - its own are never read back.
        private readonly object thisOrigin = new DraftEditOrigin("Component pane");

        private readonly ObservableCollection<ComponentPanePinRow> thisPins = new();
        private readonly ObservableCollection<ComponentPaneReplacementRow> thisReplacements = new();
        private readonly ObservableCollection<ComponentPaneLinkRow> thisLinks = new();
        private ListRowDrag<ComponentPaneLinkRow>? thisLinksDrag;

        // Each field's value as the draft holds it - a box reading otherwise holds a change.
        private readonly Dictionary<ComponentField, string> thisShownValues = new();

        private readonly DispatcherTimer thisSaveTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
        private readonly DispatcherTimer thisTypingTimer = new() { Interval = ComponentPane.TypingPause };

        // Whether a typing pause is counting - what it commits waits for it. For a test.
        internal bool IsTypingPauseRunningForTests => this.thisTypingTimer.IsEnabled;

        private DraftSession? thisSession;
        private string thisLabel = string.Empty;
        private string thisRegion = string.Empty;
        private string thisRowRegion = string.Empty;
        private BoardData? thisPublished;

        // The published board's file, for its list of files - "" for a board only the draft has.
        private string thisPublishedBoardFile = string.Empty;

        // True while the boxes are filled from the draft - nothing they say then is an edit.
        private bool thisShowing;

        private Task<DraftWorkbookEditOutcome>? thisSaveInFlight;

        // The last save the pane started, Saved raised and all - what opening the table waits for
        // before it reads the draft's file, and what a test waits for, so no save of its pane is still
        // finishing while the next test runs.
        private Task thisLastSave = Task.CompletedTask;

        internal Task LastSave => this.thisLastSave;

        // A save still running, for a test of what waits for it - a real one cannot be held.
        internal void UseLastSaveForTests(Task save) => this.thisLastSave = save;

        public ComponentPane()
        {
            this.InitializeComponent();

            this.PinsItemsControl.ItemsSource = this.thisPins;
            this.ReplacementsItemsControl.ItemsSource = this.thisReplacements;
            this.LinksItemsControl.ItemsSource = this.thisLinks;
            this.thisLinksDrag = new ListRowDrag<ComponentPaneLinkRow>(
                this, this.LinksItemsControl, this.SectionsScrollViewer, this.thisLinks, (_, _, _) => this.CommitLinks());
            this.PackageBox.ItemsSource = ComponentVocabulary.Packages;
            this.PackageBox.ItemFilter = this.PackageShowsAll;
            this.PackageBox.AddHandler(InputElement.GotFocusEvent, this.OnPackageBoxFocused, handledEventsToo: true);
            this.InitializeImagesAndFiles();
            this.InitializeWhole();
            this.InitializeBoardParts();

            this.thisSaveTimer.Tick += (_, _) => this.SaveIfDue();
            this.thisTypingTimer.Tick += (_, _) =>
            {
                this.thisTypingTimer.Stop();
                this.CommitHeld(fromTyping: true);
            };

            this.AddHandler(TextBox.TextChangedEvent, this.OnAnyTextChanged, handledEventsToo: true);
            this.AddHandler(InputElement.LostFocusEvent, this.OnAnyLostFocus, handledEventsToo: true);
            this.AddHandler(SelectingItemsControl.SelectionChangedEvent, this.OnAnySelectionChanged, handledEventsToo: true);
            this.AddHandler(InputElement.KeyDownEvent, this.OnPaneKeyDown, RoutingStrategies.Tunnel);
            this.AddHandler(InputElement.PointerWheelChangedEvent, this.OnPaneWheel, RoutingStrategies.Tunnel);
        }

        // ---- what the host sees -----------------------------------------------------------------

        // Asked to close - the new component's form's Cancel.
        public event EventHandler? CloseRequested;

        // One of the pane's own changes went into the draft - the host's list follows.
        public event EventHandler? ContentChanged;

        // The draft was saved by the pane, with changes in it - its board file named. Raised also for
        // a save started as the pane closed.
        public event EventHandler<string>? Saved;

        internal bool IsShowing => this.thisSession is not null;

        internal string Label => this.thisLabel;

        internal DraftSession? Session => this.thisSession;

        // Whether the draft changes the component shown - against the board as published, a file of
        // it replaced under its own name included (the draft's list as the session last saw it). On a new
        // board nothing is published to differ from, so nothing is said (owner request, 2026-11-02).
        internal bool IsComponentChanged => this.thisPublished is not null && this.HoldsWorkOfItsOwn;

        // The same, whether or not anything is published: a new board's component is all work of its own -
        // no mark and no undo say so, while "Also on other boards" offers it (code review, 2026-11-02).
        //
        // *** REMEMBERED WHILE ALL IT DEPENDS ON STANDS (code review, 2026-11-02). *** A commit asks this up
        // to three times on the UI thread, and once more after the save, and each read the whole board:
        // about half a second of frozen window per typing pause on the C128's larger chips. The content
        // and the lists are replaced, never changed, so the very same ones are the same answer.
        private bool HoldsWorkOfItsOwn
        {
            get
            {
                if (this.thisSession is not { } session)
                {
                    return false;
                }

                BoardContent content = session.Content;
                IReadOnlyDictionary<string, BoardDocument.StoredFileRow> listed = session.ListedFiles;
                IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? publishedFiles = this.PublishedFiles;

                if (this.thisWorkOfItsOwn is { } known &&
                    ReferenceEquals(known.Content, content) &&
                    ReferenceEquals(known.Listed, listed) &&
                    ReferenceEquals(known.PublishedFiles, publishedFiles) &&
                    ReferenceEquals(known.Published, this.thisPublished) &&
                    string.Equals(known.Label, this.thisLabel, StringComparison.Ordinal))
                {
                    return known.Answer;
                }

                bool answer = ComponentRevert.HasChanges(content.Board, this.thisPublished, this.thisLabel, publishedFiles, listed);

                this.thisWorkOfItsOwn = (content, listed, publishedFiles, this.thisPublished, this.thisLabel, answer);

                return answer;
            }
        }

        private (BoardContent Content,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow> Listed,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? PublishedFiles,
            BoardData? Published,
            string Label,
            bool Answer)? thisWorkOfItsOwn;

        internal bool HasUnsavedChanges => this.thisSession?.HasUnsavedChanges == true;

        // What the boxes hold handed to the draft first - then whether the draft waits to be saved.
        internal bool HasHeldOrUnsavedChanges()
        {
            this.CommitHeld();
            return this.HasUnsavedChanges;
        }

        internal DraftWorkbookEditOutcome? SaveFailure => this.thisSession?.SaveFailure;

        // False keeps the timer from saving - for a test that saves when it says so.
        internal bool AutoSave { get; set; } = true;

        // ###########################################################################################
        // Shows `label` of the draft held in `session`, as it reads in `region`. `published` is the
        // board as published - what "Draft" and "Undo changes to" compare with; null for a board
        // that exists only as a draft - and `publishedBoardFile` its file, whose list of files the
        // draft's is compared with. What the pane held for another component goes in first.
        // ###########################################################################################
        internal void Show(DraftSession session, string label, string? region, BoardData? published, string publishedBoardFile = "")
        {
            ArgumentNullException.ThrowIfNull(session);

            if (this.thisSession is not null && !ReferenceEquals(session, this.thisSession))
            {
                this.Close();
            }
            else
            {
                this.CommitHeld();
            }

            this.UseSession(session);

            string shown = label?.Trim() ?? string.Empty;

            // Another component, not the same one read again (code review, 2026-11-02) - asked after what
            // the pane held went in, which may have renamed the one it was on.
            if (!string.Equals(shown, this.thisLabel, StringComparison.OrdinalIgnoreCase))
            {
                this.LetGoOfWhatIsOfTheComponentShown();
            }

            this.thisLabel = shown;
            this.thisRegion = region?.Trim() ?? string.Empty;
            this.thisPublished = published;
            this.thisPublishedBoardFile = publishedBoardFile ?? string.Empty;
            this.thisIsNew = false;
            this.thisBoardPart = null;
            this.CloseClonePanel();
            this.ForgetKiCadPins();

            this.ShowComponent();
            this.RaiseComponentShown();
        }

        // ###########################################################################################
        // Closes the pane - what it holds handed to the draft first, and a save started when one is
        // waiting (finished after the pane has let go: Saved still says so).
        // ###########################################################################################
        internal void Close()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            this.CommitHeld();
            this.thisTypingTimer.Stop();

            if (session.HasUnsavedChanges)
            {
                this.thisLastSave = this.SaveSessionAsync(session);
            }

            this.Forget();
        }

        // ###########################################################################################
        // Lets go of the draft without handing anything over or saving - for a draft that is going
        // (discarded, or retired once published): nothing may be written into it any more.
        // ###########################################################################################
        internal void Forget()
        {
            this.thisTypingTimer.Stop();
            this.UseSession(null);
            this.thisLabel = string.Empty;
            this.thisIsNew = false;
            this.thisBoardPart = null;
            this.thisNetChoices = null;
            this.thisPublished = null;
            this.thisPublishedBoardFile = string.Empty;
            this.thisPins.Clear();
            this.thisReplacements.Clear();
            this.thisLinks.Clear();
            this.ForgetImagesAndFiles();
            this.ResetRowDrags();
            this.thisCategories.Clear();
            this.thisBoardFiles.Clear();
            this.thisBoardLinks.Clear();
            this.thisCredits.Clear();
            this.thisSignals.Clear();
            this.thisSignalSuggestions.Clear();
            this.thisBoardPictures.Clear();
            this.ForgetPreviews();
            this.CloseClonePanel();
            this.ForgetKiCadPins();
        }

        // ###########################################################################################
        // Everything the pane holds into the draft, and saved now. True when nothing is left unsaved.
        // ###########################################################################################
        internal async Task<bool> SaveNowAsync()
        {
            if (this.thisSession is not { } session)
            {
                return true;
            }

            this.CommitHeld();

            if (!session.HasUnsavedChanges)
            {
                return true;
            }

            Task<DraftWorkbookEditOutcome> save = this.SaveSessionAsync(session);
            this.thisLastSave = save;

            return await save == DraftWorkbookEditOutcome.Saved;
        }

        // ###########################################################################################
        // Whether the save something waited for was made, said under the title: what it was for has
        // not gone ahead (code review, 2026-11-02). Saved, a line from an earlier try goes.
        // ###########################################################################################
        internal void SayIfNotSaved(bool saved) =>
            this.ShowMessage(saved ? null : DraftSaveWording.CouldNotSave(this.SaveFailure));

        // The draft's changes that could not be saved thrown away - the file read again, one
        // undoable step. False when the file could not be read either.
        internal bool DiscardUnsavedChanges() =>
            this.thisSession?.ReloadFromFile(DraftSaveWording.ReadAgainStep) == true;

        private void UseSession(DraftSession? session)
        {
            if (ReferenceEquals(session, this.thisSession))
            {
                return;
            }

            if (this.thisSession is { } previous)
            {
                previous.ChangeComing -= this.OnSessionChangeComing;
                previous.Changed -= this.OnSessionChanged;
                previous.SaveStateChanged -= this.OnSessionSaveStateChanged;
            }

            this.thisSession = session;
            this.thisRowsFollowed = session?.Content.Board.Components;

            if (session is not null)
            {
                session.ChangeComing += this.OnSessionChangeComing;
                session.Changed += this.OnSessionChanged;
                session.SaveStateChanged += this.OnSessionSaveStateChanged;
            }

            this.thisSaveTimer.IsEnabled = session is not null;
        }

        // ---- reading the component --------------------------------------------------------------

        // ###########################################################################################
        // Everything read again from the draft as it is now: the fields, the pins, the replacement
        // parts, the links - or, when the draft no longer has the component (an undo took away the
        // change that added it), a line saying so.
        // ###########################################################################################
        private void ShowComponent()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            this.thisTypingTimer.Stop();

            // The new-component form reads the draft only to judge the label typed.
            if (this.thisIsNew)
            {
                this.ShowNewComponentState();
                return;
            }

            // A board part in place of a component (ComponentPane.Board.cs).
            if (this.thisBoardPart is not null)
            {
                this.ShowBoardPartRows();
                return;
            }

            BoardData board = session.Content.Board;

            ComponentEntry? row = ComponentPaneEdits.RowShown(board, this.thisLabel, this.thisRegion);

            this.NewComponentPanel.IsVisible = false;
            this.BoardPartPanel.IsVisible = false;
            this.GonePanel.IsVisible = row is null;
            this.SectionsScrollViewer.IsVisible = row is not null;

            if (row is null)
            {
                this.TitleText.Text = ComponentPaneWording.Title(this.thisLabel);
                this.RegionText.Text = string.Empty;
                this.RegionText.IsVisible = false;
                this.GoneText.Text = ComponentPaneWording.NotOnBoard(this.thisLabel);
                this.ShowChangedState();
                this.ShowGoneActions();
                this.TellPagesChanged();
                return;
            }

            // ###########################################################################################
            // A component with a row of ONE region only, open in the other (RowShown fell back to its
            // first row - an NTSC-only component opened while CRT shows PAL): the pane works in the row's
            // own region. In the other, every section read back empty, so each commit filed the rows again
            // beside the ones already there (slice 3c, ComponentPaneTests).
            // ###########################################################################################
            string ownRegion = row.Region?.Trim() ?? string.Empty;

            if (ownRegion.Length > 0 && !string.Equals(ownRegion, this.thisRegion, StringComparison.OrdinalIgnoreCase))
            {
                this.thisRegion = ownRegion;
            }

            this.ShowActions(board);
            this.ResetRowDrags();
            this.thisShowing = true;

            try
            {
                this.thisRowRegion = row.Region?.Trim() ?? string.Empty;

                // Before the fields: a drop-down list can select only what it lists.
                this.CategoryBox.ItemsSource = this.CategoryChoicesWith(board, row.Category);

                foreach (ComponentField field in Enum.GetValues<ComponentField>())
                {
                    string value = ComponentPaneEdits.ValueOf(row, field);
                    this.thisShownValues[field] = value;
                    this.SetBoxText(field, value);
                }

                this.thisLabelShown = row.BoardLabel?.Trim() ?? string.Empty;
                this.BoardLabelBox.Text = this.thisLabelShown;
                this.ShowLabelProblem(null);

                ComponentPane.Fill(this.thisPins, ComponentPaneEdits.PinsOf(board, this.thisLabel, this.thisRegion).Select(ComponentPanePinRow.From));
                ComponentPane.Fill(this.thisReplacements, ComponentPaneEdits.ReplacementsOf(board, this.thisLabel, this.thisRegion).Select(ComponentPaneReplacementRow.From));
                ComponentPane.Fill(this.thisLinks, ComponentPaneEdits.AllLinksOf(board, this.thisLabel).Select(ComponentPaneLinkRow.From));
                this.ShowImagesAndFiles(board);

                this.TitleText.Text = ComponentPaneWording.Title(row);
                this.RegionText.Text = ComponentPaneWording.RegionLine(this.thisRowRegion);
                this.RegionText.IsVisible = this.RegionText.Text.Length > 0;
                this.ShowRegionChoices(board);
            }
            finally
            {
                this.thisShowing = false;
            }

            this.ShowCounter();
            this.ShowCategoryProblem();
            this.ShowPinout();
            this.ShowEmptyLines();
            this.ShowCurrentPage();
            this.ShowChangedState();
        }

        private static void Fill<T>(ObservableCollection<T> rows, IEnumerable<T> items)
        {
            rows.Clear();

            foreach (T item in items)
            {
                rows.Add(item);
            }
        }

        private TextBox? BoxOf(ComponentField field) => field switch
        {
            ComponentField.FriendlyName => this.FriendlyNameBox,
            ComponentField.TechnicalNameOrValue => this.TechnicalNameBox,
            ComponentField.PartNumber => this.PartNumberBox,
            ComponentField.Description => this.DescriptionBox,
            _ => null,
        };

        private string BoxText(ComponentField field) => field switch
        {
            ComponentField.Category => this.CategoryBox.SelectedItem as string ?? string.Empty,
            ComponentField.Package => this.PackageBox.Text ?? string.Empty,
            _ => this.BoxOf(field)?.Text ?? string.Empty,
        };

        private void SetBoxText(ComponentField field, string value)
        {
            switch (field)
            {
                case ComponentField.Category:
                    this.CategoryBox.SelectedItem = value.Length > 0 ? value : null;
                    break;
                case ComponentField.Package:
                    this.PackageBox.Text = value;
                    break;
                default:
                    if (this.BoxOf(field) is { } box)
                    {
                        box.Text = value;
                    }

                    break;
            }
        }

        // The short description's length against the popup's one line.
        private void ShowCounter()
        {
            int length = ComponentPaneEdits.Normalised(ComponentField.Description, this.DescriptionBox.Text).Length;

            this.DescriptionCounter.Text = string.Create(CultureInfo.InvariantCulture, $"{length}/{ComponentVocabulary.ShortDescriptionLimit}");
            this.DescriptionCounter.Classes.Set("over", length > ComponentVocabulary.ShortDescriptionLimit);

            this.DescriptionTooLongText.Text = ComponentPaneWording.DescriptionTooLong(length);
            this.DescriptionTooLongText.IsVisible = length > ComponentVocabulary.ShortDescriptionLimit;
        }

        // The case drawn with the pins as the pane holds them, and the offer to add the case's pins.
        private void ShowPinout()
        {
            string package = this.PackageBox.Text ?? string.Empty;
            List<ComponentPinEntry> pins = this.thisPins.Select(row => row.ToEntry()).ToList();

            this.PinoutPicture.Show(package, pins);

            int missing = ComponentPaneWording.MissingCasePins(package, pins);

            this.AddCasePinsButton.IsVisible = missing > 0;
            this.AddCasePinsButton.Content = ComponentPaneWording.AddCasePins(package, missing, pins.Count == 0);
            this.ShowKiCadPins();
            this.ShowPinProblems();
            this.TellPagesChanged();
        }

        private void ShowEmptyLines()
        {
            this.NoPinsText.IsVisible = this.thisPins.Count == 0;
            this.PinsHeading.IsVisible = this.thisPins.Count > 0;
            this.NoReplacementsText.IsVisible = this.thisReplacements.Count == 0;
            this.NoLinksText.IsVisible = this.thisLinks.Count == 0;
            this.LinksHeading.IsVisible = this.thisLinks.Count > 0;
            this.ShowImageAndFileHeadings();
            this.FollowRegionChoices(this.thisBoardHasRegions);
            this.TellPagesChanged();
        }

        // "Draft", and "Undo changes to <label>", while the draft changes the component.
        private void ShowChangedState()
        {
            bool changed = this.IsComponentChanged;

            this.ChangedMark.IsVisible = changed;
            ToolTip.SetTip(this.ChangedMark, ComponentPaneWording.ComponentChangedTip);
            this.UndoChangesButton.IsVisible = changed;
            this.UndoChangesButton.Content = ReviewChangesWindow.UndoLabel(this.thisLabel);
            this.ShowAlsoOffer();
        }

        // ---- the pane's changes into the draft --------------------------------------------------

        // ###########################################################################################
        // Everything the boxes hold that the draft does not have yet, into the draft - each field one
        // step, each section one step; nothing at all when nothing changed.
        //
        // `fromTyping`: the typing pause asked. A picture's NAME is never taken then - a rename happens
        // on Enter or when its box is left (code review, 2026-10-28: a pause mid-word saved "Top" as a
        // rename, or refused it as taken and put the old name back under the caret).
        //
        // `leftABox`: a box lost the focus. A new row of another region the focus is still in waits, as it
        // does on the typing pause (ComponentPane.Pages.cs); every other commit takes all there is, as
        // nothing may wait out CRT closing or the pane being handed over.
        // ###########################################################################################
        internal void CommitHeld(bool fromTyping = false, bool leftABox = false)
        {
            if (this.thisSession is null || this.thisShowing)
            {
                return;
            }

            // A board part holds its own lists (ComponentPane.Board.cs).
            if (this.thisBoardPart is not null)
            {
                this.CommitBoardPart(fromTyping);
                return;
            }

            if (!this.SectionsScrollViewer.IsVisible)
            {
                return;
            }

            // Its label first: what follows is filed under it. Never on a typing pause (the method's
            // own note: a pause mid-word would rename the component to half a label).
            if (!fromTyping)
            {
                this.CommitLabel();
            }

            foreach (ComponentField field in Enum.GetValues<ComponentField>())
            {
                this.CommitField(field);
            }

            this.ShowCategoryProblem();

            // Put back as it was, not to None: a focused row taken off the page commits inside this one.
            RowsWaiting waitingBefore = this.thisRowsWaiting;
            this.thisRowsWaiting = fromTyping ? RowsWaiting.AllOfAnotherRegion : leftABox ? RowsWaiting.TheOneInFocus : RowsWaiting.None;

            try
            {
                this.CommitPins();
                this.CommitBaselines();
                this.CommitPictures();
                this.CommitReplacements();
                this.CommitFiles();
                this.CommitLinks();
                this.LetGoOfRowsOfOtherRegions();
            }
            finally
            {
                this.thisRowsWaiting = waitingBefore;
            }

            // The case drawn as the package now reads.
            this.ShowPinout();
        }

        private void CommitField(ComponentField field)
        {
            string text = this.BoxText(field);
            string wanted = ComponentPaneEdits.Normalised(field, text);

            if (this.thisShownValues.TryGetValue(field, out string? shown) && string.Equals(shown, wanted, StringComparison.Ordinal))
            {
                return;
            }

            // Before the change is told: what follows it reads this - a part number typed alone is what
            // "Also on other boards" waits for, and it stayed hidden until the next change.
            this.thisShownValues[field] = wanted;

            this.Apply(
                DraftStepWording.ComponentField(this.thisLabel, ComponentPaneEdits.NameOf(field)),
                board => ComponentPaneEdits.WithField(board, this.thisLabel, this.thisRegion, field, text));

            // The headline says the names, as the popup's title does - and the pane reads nothing back after
            // its own edit (code review, 2026-11-04).
            if (this.thisSession is { } session &&
                ComponentPaneEdits.RowShown(session.Content.Board, this.thisLabel, this.thisRegion) is { } row)
            {
                this.TitleText.Text = ComponentPaneWording.Title(row);
            }
        }

        private void CommitPins()
        {
            List<ComponentPinEntry> pins = this.thisPins.Select(row => row.ToEntry()).ToList();

            this.Apply(
                DraftStepWording.ComponentPins(this.thisLabel),
                board => ComponentPaneEdits.WithPins(board, this.thisLabel, this.thisRegion, pins));
        }

        private void CommitReplacements()
        {
            List<ComponentReplacementEntry> replacements = this.thisReplacements.Where(row => !this.IsWaitingToBeTyped(row)).Select(row => row.ToEntry()).ToList();

            this.Apply(
                DraftStepWording.ComponentReplacements(this.thisLabel),
                board => ComponentPaneEdits.WithReplacements(board, this.thisLabel, this.thisRegion, replacements));
        }

        private void CommitLinks()
        {
            List<ComponentLinkEntry> links = this.thisLinks.Select(row => row.ToEntry()).ToList();

            this.Apply(
                DraftStepWording.ComponentLinks(this.thisLabel),
                board => ComponentPaneEdits.WithAllLinks(board, this.thisLabel, links));
        }

        // One change applied to the draft held in memory, as the pane's own step.
        private void Apply(string description, Func<BoardData, BoardData> transform)
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            bool changed = session.Apply(
                description,
                content => transform(content.Board) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null,
                this.thisOrigin);

            if (changed)
            {
                this.ShowChangedState();
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        // ---- the draft's other changes ----------------------------------------------------------

        // ###########################################################################################
        // Another change is coming - what the pane holds goes in first, so that change lands on top.
        // An undo or a redo carries the mark of the step it moves, which may be the pane's own: only
        // the pane's own EDIT is skipped (found by its test, 2026-10-25: an undo of a field typed in
        // the pane left the box showing the undone text).
        // ###########################################################################################
        private void OnSessionChangeComing(object? sender, DraftSessionChange change)
        {
            if (!this.IsOwnEdit(change))
            {
                this.CommitHeld();
            }
        }

        // Another editor's change, an undo, a redo, the file read again: the component read again.
        private void OnSessionChanged(object? sender, DraftSessionChange change)
        {
            // A write run off the UI thread (DraftSession.WriteFileDirectly - the KiCad data window's
            // Remove, "Bring my draft up to date") raises it there, and the pane drew from there and
            // threw (2026-10-26) - as the table does, it goes to the UI thread.
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.Post(() => this.OnSessionChanged(sender, change));
                return;
            }

            if (!ReferenceEquals(sender, this.thisSession) || this.thisSession is not { } session)
            {
                return;
            }

            // Every change, the pane's own too: what the next is compared with is the rows as this one left them.
            BoardData board = session.Content.Board;
            IReadOnlyList<ComponentEntry>? before = this.thisRowsFollowed;
            this.thisRowsFollowed = board.Components;

            if (!this.IsOwnEdit(change))
            {
                // An undo or a redo of a rename left the component under another label.
                this.FollowRelabelInDraft(before, board);

                this.ShowComponent();
                this.FollowKiCadFiles();
            }
        }

        private bool IsOwnEdit(DraftSessionChange change) =>
            change.Kind == DraftSessionChangeKind.Edit && ReferenceEquals(change.Origin, this.thisOrigin);

        // ---- typing, leaving a box, picking ------------------------------------------------------

        private void OnAnyTextChanged(object? sender, TextChangedEventArgs e)
        {
            if (this.thisShowing || this.thisSession is null)
            {
                return;
            }

            // The new-component form is judged as it is typed in, and goes into the draft only when added.
            if (this.thisIsNew)
            {
                this.OnNewComponentTextChanged();
                return;
            }

            if (ReferenceEquals(e.Source, this.DescriptionBox))
            {
                this.ShowCounter();
            }

            // A picture's name waits for Enter or its box being left - typing it starts no pause. So do
            // a component's label and a category's name (2026-11-01). Nor does a net typed into the new
            // signal's picker: it is no part of the draft until "Add signal", and the commit a pause
            // made swapped the picker's list under it (code review, 2026-10-30).
            if (e.Source is TextBox { DataContext: ComponentPaneBoardPictureRow or ComponentPaneCategoryRow } ||
                ReferenceEquals(e.Source, this.BoardLabelBox) ||
                ReferenceEquals(e.Source, this.NewCategoryNameBox) ||
                (e.Source is Visual typedIn && ReferenceEquals(typedIn.FindAncestorOfType<AutoCompleteBox>(), this.NewSignalNetBox)))
            {
                return;
            }

            this.thisTypingTimer.Stop();
            this.thisTypingTimer.Start();
        }

        private void OnAnyLostFocus(object? sender, RoutedEventArgs e)
        {
            this.thisTypingTimer.Stop();
            this.CommitHeld(leftABox: true);
            this.ShowPinout();
        }

        // A drop-down list picked (a pin's direction, a part's kind) - into the draft once the pick
        // has reached the row.
        private void OnAnySelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisShowing || this.thisSession is null || e.Source is not ComboBox)
            {
                return;
            }

            // The new-component form holds what it holds until it is added - judged as it is picked in.
            if (this.thisIsNew)
            {
                this.OnNewComponentTextChanged();
                return;
            }

            Dispatcher.UIThread.Post(() => this.CommitHeld(), DispatcherPriority.Background);
        }

        // ###########################################################################################
        // THE WHEEL NEVER CHANGES A CLOSED DROP-DOWN LIST HERE (code review, 2026-11-04). A ComboBox with the
        // focus steps its selection under the wheel, so scrolling the page with the pointer over the category
        // just picked gave the component another one, a step of the draft nobody meant - and over the clone
        // panel's board, each notch read another board's file. A pick is made from the open list, as on Main's
        // device picker. The page scrolls on, by the wheel's own step (Avalonia's ScrollContentPresenter), as it
        // does over any other box; an open list keeps its wheel, and one without the focus never took it.
        // ###########################################################################################
        private const double WheelStep = 50;

        private void OnPaneWheel(object? sender, PointerWheelEventArgs e)
        {
            if (e.Source is not Visual source ||
                source.FindAncestorOfType<ComboBox>(includeSelf: true) is not { IsDropDownOpen: false, IsFocused: true } box)
            {
                return;
            }

            e.Handled = true;

            if (box.FindAncestorOfType<ScrollViewer>() is { } page)
            {
                page.Offset = page.Offset.WithY(page.Offset.Y - e.Delta.Y * ComponentPane.WheelStep);
            }
        }

        // Enter in a one-line box takes what it says into the draft at once.
        private void OnPaneKeyDown(object? sender, KeyEventArgs e)
        {
            if (this.AddOnEnter(e))
            {
                return;
            }

            if (e.Key == Key.Enter && e.Source is TextBox { AcceptsReturn: false })
            {
                this.thisTypingTimer.Stop();
                this.CommitHeld();
                this.ShowPinout();
            }
        }

        // ---- buttons ----------------------------------------------------------------------------

        // (Its Close button went on 2026-10-30 - the left column's "Board status" closes the pane, and
        // stays put over a picture's name that cannot be taken, as Close did: TabDrafts.ShowBoardStatus.)

        // The new component's form: Cancel adds nothing and closes the pane - "Board status" shows.
        private void OnCancelNewComponentClick(object? sender, RoutedEventArgs e) => this.CloseRequested?.Invoke(this, EventArgs.Empty);

        // ###########################################################################################
        // "Undo changes to <label>": the component back as published - every row of it, every region
        // - as ONE step (ComponentRevert), so the draft's Undo brings the changes back. A file of it
        // replaced under its own name goes back too, by the draft's list of files, once the rows are
        // saved (DraftEdits.PutBackPublishedFilesAsync - code review, 2026-10-26); the returned task
        // finishes with it.
        // ###########################################################################################
        private async void OnUndoChangesClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.UndoComponentChanges();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Undoing the changes to a component in the pane failed - [{ex}]");
            }
        }

        internal Task UndoComponentChanges()
        {
            if (this.thisSession is not { } session || this.thisLabel.Length == 0)
            {
                return Task.CompletedTask;
            }

            this.CommitHeld();

            BoardData? published = this.thisPublished;
            string label = this.thisLabel;
            string description = DraftStepWording.ComponentReverted(label);

            // Asked before the rows go back: a relabelled component's files are its old label's.
            List<string> replaced = ComponentRevert.ReplacedFiles(
                session.Content.Board, published, label, this.PublishedFiles, session.ListedFiles);

            // Not the pane's own step: the pane reads the component back afterwards.
            session.Apply(
                description,
                content => ComponentRevert.RowsToPublished(content.Board, published, label) is BoardData board
                    ? content with { Board = board }
                    : null);

            this.ContentChanged?.Invoke(this, EventArgs.Empty);

            return replaced.Count == 0 ? Task.CompletedTask : this.PutBackReplacedFilesAsync(session, replaced, description);
        }

        // The files put back - the board file written, which the host follows as a save of the pane's.
        private async Task PutBackReplacedFilesAsync(DraftSession session, IReadOnlyList<string> names, string description)
        {
            DraftWorkbookEditOutcome outcome = await DraftEdits.PutBackPublishedFilesAsync(session, names, this.thisPublishedBoardFile, description);

            if (outcome == DraftWorkbookEditOutcome.Saved)
            {
                this.Saved?.Invoke(this, session.ExcelDataFile);
            }
            else
            {
                Logger.Warning($"Could not put the published files [{string.Join(", ", names)}] back in the draft of [{session.ExcelDataFile}] - [{outcome}]");
            }
        }

        private void OnAddPinClick(object? sender, RoutedEventArgs e)
        {
            this.thisPins.Add(new ComponentPanePinRow { Region = this.RegionForNewPin() });
            this.ShowEmptyLines();
        }

        private void OnRemovePinClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPanePinRow row } && this.thisPins.Remove(row))
            {
                this.CommitPins();
                this.ShowPinout();
                this.ShowEmptyLines();
            }
        }

        // Every pin number of the case not listed yet, added in number order - one step.
        private void OnAddCasePinsClick(object? sender, RoutedEventArgs e) => this.AddCasePins();

        internal void AddCasePins()
        {
            if (ComponentVocabulary.PinCountOf(this.PackageBox.Text) is not int count)
            {
                return;
            }

            // The package typed goes in first, so the step that adds its pins follows it.
            this.CommitHeld();

            List<ComponentPinEntry> pins = ComponentPaneEdits.WithEveryPinOf(
                this.thisPins.Select(row => row.ToEntry()).ToList(),
                count,
                this.RegionForNewPin());

            ComponentPane.Fill(this.thisPins, pins.Select(ComponentPanePinRow.From));

            this.CommitPins();
            this.ShowPinout();
            this.ShowEmptyLines();
        }

        private void OnAddReplacementClick(object? sender, RoutedEventArgs e)
        {
            this.thisReplacements.Add(new ComponentPaneReplacementRow { Region = this.thisRowRegion });
            this.ShowEmptyLines();
        }

        private void OnRemoveReplacementClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneReplacementRow row } && this.thisReplacements.Remove(row))
            {
                this.CommitReplacements();
                this.ShowEmptyLines();
            }
        }

        private void OnAddLinkClick(object? sender, RoutedEventArgs e)
        {
            this.thisLinks.Add(new ComponentPaneLinkRow());
            this.ShowEmptyLines();
        }

        private void OnRemoveLinkClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneLinkRow row } && this.thisLinks.Remove(row))
            {
                this.CommitLinks();
                this.ShowEmptyLines();
            }
        }

        // A link dragged by its handle, or moved by Alt+Up and Alt+Down, as the board's are (owner request,
        // 2026-11-02: "the reordering is missing in the component part").
        private void OnLinkHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisLinksDrag?.Press(sender, e);

        internal bool MoveLink(ComponentPaneLinkRow row, int by)
        {
            if (!this.MoveRow(this.thisLinks, row, by))
            {
                return false;
            }

            this.CommitLinks();
            return true;
        }

        // ---- saving -----------------------------------------------------------------------------

        private void SaveIfDue()
        {
            if (this.thisSession is not { } session || this.thisSaveInFlight is not null)
            {
                return;
            }

            if (this.AutoSave && session.IsSaveDueNow)
            {
                this.thisLastSave = this.SaveSessionAsync(session);
            }
        }

        // Runs the timer's check now - for a test, whose session clock it moves itself.
        internal void SaveIfDueForTests() => this.SaveIfDue();

        // ###########################################################################################
        // Saves the draft off the UI thread, after any save already running. Never throws: a save that
        // fails is Failed, which the draft's header says with its reason.
        // ###########################################################################################
        private async Task<DraftWorkbookEditOutcome> SaveSessionAsync(DraftSession session)
        {
            // *** A FINISHED SAVE IS FINISHED, ITS OWN CONTINUATION RUN OR NOT (owner report, 2026-11-06: CRT
            // froze for good, a note typed and the "Resources" tab clicked). *** That continuation clears the
            // field - posted at the priority the save began under, the timer's low one; a save asked for from
            // a click comes first, and awaiting a finished task returns at once: it looped for ever here.
            while (this.thisSaveInFlight is { IsCompleted: false } running)
            {
                await running;
            }

            bool hadChanges = session.HasUnsavedChanges;
            Task<DraftWorkbookEditOutcome> save = session.SaveAsync();
            this.thisSaveInFlight = save;

            DraftWorkbookEditOutcome outcome;

            try
            {
                outcome = await save;
            }
            finally
            {
                if (ReferenceEquals(this.thisSaveInFlight, save))
                {
                    this.thisSaveInFlight = null;
                }
            }

            if (outcome == DraftWorkbookEditOutcome.Saved && hadChanges)
            {
                this.Saved?.Invoke(this, session.ExcelDataFile);
            }

            // A file added is listed - with its bytes, what "Also on other boards" goes by - once written.
            if (ReferenceEquals(session, this.thisSession))
            {
                this.ShowAlsoOffers();
            }

            return outcome;
        }
    }

    // ###########################################################################################
    // One pin as the pane holds it. A drop-down list's empty write (a list built before its value
    // arrived) is not a change: a direction is picked, never cleared, in the pane.
    // ###########################################################################################
    public sealed class ComponentPanePinRow : INotifyPropertyChanged
    {
        private string thisPin = string.Empty;
        private string thisName = string.Empty;
        private string thisDirection = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Region { get; init; } = string.Empty;

        public string Pin { get => this.thisPin; set => this.Set(ref this.thisPin, value ?? string.Empty); }

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string? Direction
        {
            get => this.thisDirection.Length == 0 ? null : this.thisDirection;
            set
            {
                if (value is not null)
                {
                    this.Set(ref this.thisDirection, value);
                }
            }
        }

        // The vocabulary, and a word the board already holds that is not in it.
        public IReadOnlyList<string> DirectionChoices =>
            ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.PinDirections, this.thisDirection)
                ? ComponentVocabulary.PinDirections
                : [.. ComponentVocabulary.PinDirections, this.thisDirection];

        // The net the pin is on on this board, from its KiCad data - shown, never part of the entry.
        public string Net { get; private set; } = string.Empty;

        public bool HasNet => this.Net.Length > 0;

        public void UseNet(string? net)
        {
            string wanted = net ?? string.Empty;

            if (string.Equals(this.Net, wanted, StringComparison.Ordinal))
            {
                return;
            }

            this.Net = wanted;
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.Net)));
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.HasNet)));
        }

        // A pin the component's case does not have - an error, marked on its box (ShowPinProblems).
        public bool PinOutside => this.PinProblem.Length > 0;

        public string PinProblem { get; private set; } = string.Empty;

        // The case that has no such pin - null when it can be on it.
        public void UseCaseThatRefuses(string? package)
        {
            string wanted = package is null ? string.Empty : BoardDataChecks.PinNotOnPackage(this.thisPin, package);

            if (!string.Equals(this.PinProblem, wanted, StringComparison.Ordinal))
            {
                this.PinProblem = wanted;
                this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.PinProblem)));
                this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.PinOutside)));
            }
        }

        public static ComponentPanePinRow From(ComponentPinEntry pin) => new()
        {
            Region = pin.Region ?? string.Empty,
            Pin = pin.Pin ?? string.Empty,
            Name = pin.Name ?? string.Empty,
            Direction = string.IsNullOrWhiteSpace(pin.Direction) ? null : pin.Direction.Trim(),
        };

        public ComponentPinEntry ToEntry() => new()
        {
            Region = this.Region,
            Pin = this.thisPin,
            Name = this.thisName,
            Direction = this.thisDirection,
        };

        private void Set(ref string field, string value, [CallerMemberName] string? name = null)
        {
            if (string.Equals(field, value, StringComparison.Ordinal))
            {
                return;
            }

            field = value;
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    // One replacement part as the pane holds it - the kind picked as a pin's direction is, its note of
    // several lines.
    public sealed class ComponentPaneReplacementRow : ComponentPaneRegionalRow
    {
        private string thisPartNumber = string.Empty;
        private string thisKind = string.Empty;
        private string thisNote = string.Empty;

        public string PartNumber { get => this.thisPartNumber; set => this.Set(ref this.thisPartNumber, value ?? string.Empty); }

        public string? Kind
        {
            get => this.thisKind.Length == 0 ? null : this.thisKind;
            set
            {
                if (value is not null)
                {
                    this.Set(ref this.thisKind, value);
                }
            }
        }

        public string Note { get => this.thisNote; set => this.Set(ref this.thisNote, value ?? string.Empty); }

        // ComponentPaneEdits.WithReplacements' rule for a row it leaves out.
        public override bool HoldsNothingYet =>
            string.IsNullOrWhiteSpace(this.thisPartNumber) && string.IsNullOrWhiteSpace(this.thisKind) && string.IsNullOrWhiteSpace(this.thisNote);

        public IReadOnlyList<string> KindChoices =>
            ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.ReplacementKinds, this.thisKind)
                ? ComponentVocabulary.ReplacementKinds
                : [.. ComponentVocabulary.ReplacementKinds, this.thisKind];

        public static ComponentPaneReplacementRow From(ComponentReplacementEntry row) => new()
        {
            Region = row.Region ?? string.Empty,
            PartNumber = row.PartNumber ?? string.Empty,
            Kind = string.IsNullOrWhiteSpace(row.Kind) ? null : row.Kind.Trim(),
            Note = row.Note ?? string.Empty,
        };

        public ComponentReplacementEntry ToEntry() => new()
        {
            Region = this.Region,
            PartNumber = this.thisPartNumber,
            Kind = this.thisKind,
            Note = this.thisNote,
        };
    }

    // One link as the pane holds it.
    public sealed class ComponentPaneLinkRow : ComponentPaneRegionalRow
    {
        private string thisName = string.Empty;
        private string thisUrl = string.Empty;

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string Url
        {
            get => this.thisUrl;
            set
            {
                this.Set(ref this.thisUrl, value ?? string.Empty);
                this.Raise(nameof(this.UrlNotWeb));
                this.Raise(nameof(this.CanOpen));
            }
        }

        // "Open link" is on while the address is one a browser can be sent to (LinkTarget) - a board link's rule.
        public bool CanOpen => LinkTarget.For(this.thisUrl) is not null;

        // An address the checks refuse, marked on its box (ComponentPaneBoardLinkRow says why).
        public bool UrlNotWeb => BoardDataChecks.IsRefusedLink(this.thisUrl);

        public string NotWebText => ComponentPaneWording.LinkNotWeb;

        public static ComponentPaneLinkRow From(ComponentLinkEntry row) => new()
        {
            Region = row.Region ?? string.Empty,
            Name = row.Name ?? string.Empty,
            Url = row.Url ?? string.Empty,
        };

        public ComponentLinkEntry ToEntry() => new() { Region = this.Region, Name = this.thisName, Url = this.thisUrl };
    }
}
