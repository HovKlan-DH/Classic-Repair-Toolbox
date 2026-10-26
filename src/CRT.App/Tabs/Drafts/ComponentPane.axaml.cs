using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
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
    //     |   Repair notes [                              ]   |
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
    //   - ITS SCOPE BASELINES, PICTURES AND FILES TOO (slice 3b, 2026-10-26): a file added goes into the
    //     draft under a name of the board's own and is saved at once (ComponentPane.Files.cs).
    //   - AND THE COMPONENT AS A WHOLE (slice 3c, 2026-10-26): a new one added, one deleted, "Show on
    //     board" and "Preview popup" (ComponentPane.Whole.cs) - the component editor's window is gone.
    //   - AND, IN PLACE OF A COMPONENT, ONE OF THE BOARD'S OWN LISTS (slice 4a, 2026-10-26): its files
    //     and links, its credits or its important signals (ComponentPane.Board.cs).
    //
    // The file map:
    //   ComponentPane.axaml.cs  - showing a component, the fields, pins, replacement parts and links,
    //                             committing and saving, Undo changes
    //   ComponentPane.Files.cs  - scope baselines (with the new-baseline panel and the capture),
    //                             pictures and files: adding, changing, thumbnails
    //   ComponentPane.Whole.cs  - the new-component form, Delete component and "Bring back", the
    //                             buttons that ask the host to show the component on the board or
    //                             preview its popup
    //   ComponentPane.Board.cs  - the board's files and links, credits and important signals
    //   ComponentPane.BoardPictures.cs - the board's pictures and the colours a marked component is
    //                             drawn in on each, with their previews
    //   ComponentPane.SignalSuggestions.cs - important signals suggested from the KiCad net names, and
    //                             "Show on board" lighting a net up on the Schematics tab
    //   ComponentPane.KiCadPins.cs - the component's pins from the board's KiCad data: "Fill from
    //                             KiCad" and each pin's net on this board
    //   ComponentPane.Copy.cs   - "Copy from another board": the same part's data from another board;
    //                             "Also on other boards": this draft's change of the part, and the
    //                             files it replaced, onto the other boards' own drafts
    // ###########################################################################################
    public partial class ComponentPane : UserControl
    {
        // How long typing may pause before what was typed goes into the draft.
        internal static readonly TimeSpan TypingPause = TimeSpan.FromSeconds(1);

        // The mark the pane's own changes carry in the draft's history - its own are never read back.
        private readonly object thisOrigin = new DraftEditOrigin("Component pane");

        private readonly ObservableCollection<ComponentPanePinRow> thisPins = new();
        private readonly ObservableCollection<ComponentPaneReplacementRow> thisReplacements = new();
        private readonly ObservableCollection<ComponentPaneLinkRow> thisLinks = new();

        // Each field's value as the draft holds it - a box reading otherwise holds a change.
        private readonly Dictionary<ComponentField, string> thisShownValues = new();

        private readonly DispatcherTimer thisSaveTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
        private readonly DispatcherTimer thisTypingTimer = new() { Interval = ComponentPane.TypingPause };

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
            this.PackageBox.ItemsSource = ComponentVocabulary.Packages;
            this.InitializeImagesAndFiles();
            this.InitializeWhole();
            this.InitializeBoardParts();

            this.thisSaveTimer.Tick += (_, _) => this.SaveIfDue();
            this.thisTypingTimer.Tick += (_, _) =>
            {
                this.thisTypingTimer.Stop();
                this.CommitHeld();
            };

            this.AddHandler(TextBox.TextChangedEvent, this.OnAnyTextChanged, handledEventsToo: true);
            this.AddHandler(InputElement.LostFocusEvent, this.OnAnyLostFocus, handledEventsToo: true);
            this.AddHandler(SelectingItemsControl.SelectionChangedEvent, this.OnAnySelectionChanged, handledEventsToo: true);
            this.AddHandler(InputElement.KeyDownEvent, this.OnPaneKeyDown, RoutingStrategies.Tunnel);
        }

        // ---- what the host sees -----------------------------------------------------------------

        // Asked to close (its Close button).
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
        // it replaced under its own name included (the draft's list as the session last saw it).
        internal bool IsComponentChanged =>
            this.thisSession is { } session &&
            ComponentRevert.HasChanges(session.Content.Board, this.thisPublished, this.thisLabel, DraftFiles.ListedIn(this.thisPublishedBoardFile), session.ListedFiles);

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
        // board as published - what "Changed" and "Undo changes to" compare with; null for a board
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

            this.thisLabel = label?.Trim() ?? string.Empty;
            this.thisRegion = region?.Trim() ?? string.Empty;
            this.thisPublished = published;
            this.thisPublishedBoardFile = publishedBoardFile ?? string.Empty;
            this.thisIsNew = false;
            this.thisBoardPart = null;
            this.CloseCopyPanel();
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
            this.thisBoardFiles.Clear();
            this.thisBoardLinks.Clear();
            this.thisCredits.Clear();
            this.thisSignals.Clear();
            this.thisSignalSuggestions.Clear();
            this.thisBoardPictures.Clear();
            this.CloseCopyPanel();
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
                this.TitleText.Text = this.thisLabel;
                this.RegionText.Text = string.Empty;
                this.RegionText.IsVisible = false;
                this.GoneText.Text = ComponentPaneWording.NotOnBoard(this.thisLabel);
                this.ShowChangedState();
                this.ShowGoneActions();
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
            this.thisShowing = true;

            try
            {
                this.thisRowRegion = row.Region?.Trim() ?? string.Empty;

                foreach (ComponentField field in Enum.GetValues<ComponentField>())
                {
                    string value = ComponentPaneEdits.ValueOf(row, field);
                    this.thisShownValues[field] = value;
                    this.SetBoxText(field, value);
                }

                this.CategoryBox.ItemsSource = ComponentListBuilder.BuildDistinctCategories(board)
                    .OrderBy(category => category, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                ComponentPane.Fill(this.thisPins, ComponentPaneEdits.PinsOf(board, this.thisLabel, this.thisRegion).Select(ComponentPanePinRow.From));
                ComponentPane.Fill(this.thisReplacements, ComponentPaneEdits.ReplacementsOf(board, this.thisLabel, this.thisRegion).Select(ComponentPaneReplacementRow.From));
                ComponentPane.Fill(this.thisLinks, ComponentPaneEdits.LinksOf(board, this.thisLabel).Select(ComponentPaneLinkRow.From));
                this.ShowImagesAndFiles(board);

                this.TitleText.Text = ComponentPaneWording.Title(row);
                this.RegionText.Text = ComponentPaneWording.RegionLine(this.thisRowRegion);
                this.RegionText.IsVisible = this.RegionText.Text.Length > 0;
            }
            finally
            {
                this.thisShowing = false;
            }

            this.ShowCounter();
            this.ShowPinout();
            this.ShowEmptyLines();
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
            ComponentField.RepairNotes => this.RepairNotesBox,
            _ => null,
        };

        private string BoxText(ComponentField field) => field switch
        {
            ComponentField.Category => this.CategoryBox.Text ?? string.Empty,
            ComponentField.Package => this.PackageBox.Text ?? string.Empty,
            _ => this.BoxOf(field)?.Text ?? string.Empty,
        };

        private void SetBoxText(ComponentField field, string value)
        {
            switch (field)
            {
                case ComponentField.Category:
                    this.CategoryBox.Text = value;
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
        }

        private void ShowEmptyLines()
        {
            this.NoPinsText.IsVisible = this.thisPins.Count == 0;
            this.PinsHeading.IsVisible = this.thisPins.Count > 0;
            this.NoReplacementsText.IsVisible = this.thisReplacements.Count == 0;
            this.NoLinksText.IsVisible = this.thisLinks.Count == 0;
            this.LinksHeading.IsVisible = this.thisLinks.Count > 0;
            this.ShowImageAndFileHeadings();
        }

        // "Changed", and "Undo changes to <label>", while the draft changes the component.
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
        // ###########################################################################################
        internal void CommitHeld()
        {
            if (this.thisSession is null || this.thisShowing)
            {
                return;
            }

            // A board part holds its own lists (ComponentPane.Board.cs).
            if (this.thisBoardPart is not null)
            {
                this.CommitBoardPart();
                return;
            }

            if (!this.SectionsScrollViewer.IsVisible)
            {
                return;
            }

            foreach (ComponentField field in Enum.GetValues<ComponentField>())
            {
                this.CommitField(field);
            }

            this.CommitPins();
            this.CommitBaselines();
            this.CommitPictures();
            this.CommitReplacements();
            this.CommitFiles();
            this.CommitLinks();

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

            this.Apply(
                DraftStepWording.ComponentField(this.thisLabel, ComponentPaneEdits.NameOf(field)),
                board => ComponentPaneEdits.WithField(board, this.thisLabel, this.thisRegion, field, text));

            this.thisShownValues[field] = wanted;
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
            List<ComponentReplacementEntry> replacements = this.thisReplacements.Select(row => row.ToEntry()).ToList();

            this.Apply(
                DraftStepWording.ComponentReplacements(this.thisLabel),
                board => ComponentPaneEdits.WithReplacements(board, this.thisLabel, this.thisRegion, replacements));
        }

        private void CommitLinks()
        {
            List<ComponentLinkEntry> links = this.thisLinks.Select(row => row.ToEntry()).ToList();

            this.Apply(
                DraftStepWording.ComponentLinks(this.thisLabel),
                board => ComponentPaneEdits.WithLinks(board, this.thisLabel, links));
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

            if (ReferenceEquals(sender, this.thisSession) && !this.IsOwnEdit(change))
            {
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

            this.thisTypingTimer.Stop();
            this.thisTypingTimer.Start();
        }

        private void OnAnyLostFocus(object? sender, RoutedEventArgs e)
        {
            this.thisTypingTimer.Stop();
            this.CommitHeld();
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

            Dispatcher.UIThread.Post(this.CommitHeld, DispatcherPriority.Background);
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

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.CloseRequested?.Invoke(this, EventArgs.Empty);

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
                session.Content.Board, published, label, DraftFiles.ListedIn(this.thisPublishedBoardFile), session.ListedFiles);

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
            this.thisPins.Add(new ComponentPanePinRow { Region = this.thisRowRegion });
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
                this.thisRowRegion);

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
            while (this.thisSaveInFlight is { } running)
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

    // One replacement part as the pane holds it - the kind picked as a pin's direction is.
    public sealed class ComponentPaneReplacementRow : INotifyPropertyChanged
    {
        private string thisPartNumber = string.Empty;
        private string thisKind = string.Empty;
        private string thisChanges = string.Empty;
        private string thisSource = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Region { get; init; } = string.Empty;

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

        public string Changes { get => this.thisChanges; set => this.Set(ref this.thisChanges, value ?? string.Empty); }

        public string Source { get => this.thisSource; set => this.Set(ref this.thisSource, value ?? string.Empty); }

        public IReadOnlyList<string> KindChoices =>
            ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.ReplacementKinds, this.thisKind)
                ? ComponentVocabulary.ReplacementKinds
                : [.. ComponentVocabulary.ReplacementKinds, this.thisKind];

        public static ComponentPaneReplacementRow From(ComponentReplacementEntry row) => new()
        {
            Region = row.Region ?? string.Empty,
            PartNumber = row.PartNumber ?? string.Empty,
            Kind = string.IsNullOrWhiteSpace(row.Kind) ? null : row.Kind.Trim(),
            Changes = row.Changes ?? string.Empty,
            Source = row.Source ?? string.Empty,
        };

        public ComponentReplacementEntry ToEntry() => new()
        {
            Region = this.Region,
            PartNumber = this.thisPartNumber,
            Kind = this.thisKind,
            Changes = this.thisChanges,
            Source = this.thisSource,
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

    // One link as the pane holds it.
    public sealed class ComponentPaneLinkRow : INotifyPropertyChanged
    {
        private string thisName = string.Empty;
        private string thisUrl = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string Url { get => this.thisUrl; set => this.Set(ref this.thisUrl, value ?? string.Empty); }

        public static ComponentPaneLinkRow From(ComponentLinkEntry row) => new()
        {
            Name = row.Name ?? string.Empty,
            Url = row.Url ?? string.Empty,
        };

        public ComponentLinkEntry ToEntry() => new() { Name = this.thisName, Url = this.thisUrl };

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
}
