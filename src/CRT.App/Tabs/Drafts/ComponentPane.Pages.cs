using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A COMPONENT'S PAGES (owner decision, 2026-11-01: "the functionality in this is too massive, if all
    // data is filled, so I think we need to split this into multiple menu items" - as tabs over the pane,
    // so the left column keeps its room for the component list).
    //
    // Owns:
    //   - THE TABS: Basics, Replacement parts, Pinout, Pictures, Oscilloscope baseline, Files and Links,
    //     each saying how many it holds. One page shows; every section stays filled behind it, so what the
    //     pane commits is what it always was. The page stays as another component is picked: walking a
    //     board's chips stays on one page.
    //   - THE REGION SHOWN, on a board that has regions (BoardRegions). What a page lists is what that
    //     region shows: its own rows and those for every region. A row's "Applies to" moves it (one
    //     leaving the region shown is let go of here, and said), and Basics and Pinout have a "This data
    //     is different per region" tick instead - a set of rows that belong together. The switch under
    //     the component list is on every page again (PaneRegionSwitch; RegionPicked - Main's, so all of
    //     CRT follows): under Basics' and Pinout's tick and only while it is ticked (owner request,
    //     2026-11-01), and always on the five others (owner request, 2026-11-02).
    //   - ITS LABEL (owner report, 2026-11-01: "unable to rename the component label"): on Enter or when
    //     its box is left, never on a pause in typing - every row of the component moves with it.
    //   - WHAT SUBMIT WOULD REFUSE, MARKED WHERE IT IS AND ON ITS TAB: no category; a pin the case does not
    //     have; a baseline lacking its signal name or image; a link's address that is no web address.
    //   - THE TEMPORARY TOOL making pinouts of oscilloscope baselines (BaselinePinMigration - gone before
    //     CRT 4.0 goes live, with this section).
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        private ComponentPage thisPage = ComponentPage.Basics;

        // The draft's component rows as the last change left them - what the next is compared with, for a
        // rename an undo or a redo took back (FollowRelabelInDraft).
        private IReadOnlyList<ComponentEntry>? thisRowsFollowed;

        // The label as the box last read it - its capitals are a change of their own.
        private string thisLabelShown = string.Empty;

        // Asked to show another region - Main's, so all of CRT follows.
        public event EventHandler<string>? RegionPicked;

        // Whether the board open has regions - what offers "Applies to" and the ticks.
        private bool thisBoardHasRegions;

        internal ComponentPage Page => this.thisPage;

        private void OnPageClick(object? sender, RoutedEventArgs e)
        {
            foreach (ComponentPage page in Enum.GetValues<ComponentPage>())
            {
                if (ReferenceEquals(sender, this.PageButton(page)))
                {
                    this.ShowPage(page);
                    return;
                }
            }
        }

        // The page asked for, what the boxes hold handed to the draft first, at its top.
        internal void ShowPage(ComponentPage page)
        {
            this.CommitHeld();
            this.thisPage = page;
            this.ShowCurrentPage();
            this.SectionsScrollViewer.Offset = default;
        }

        private Button PageButton(ComponentPage page) => page switch
        {
            ComponentPage.Basics => this.BasicsPageButton,
            ComponentPage.Replacements => this.ReplacementsPageButton,
            ComponentPage.Pinout => this.PinoutPageButton,
            ComponentPage.Pictures => this.PicturesPageButton,
            ComponentPage.Baselines => this.BaselinesPageButton,
            ComponentPage.Files => this.FilesPageButton,
            _ => this.LinksPageButton,
        };

        private TextBlock PageLabel(ComponentPage page) => page switch
        {
            ComponentPage.Basics => this.BasicsPageLabel,
            ComponentPage.Replacements => this.ReplacementsPageLabel,
            ComponentPage.Pinout => this.PinoutPageLabel,
            ComponentPage.Pictures => this.PicturesPageLabel,
            ComponentPage.Baselines => this.BaselinesPageLabel,
            ComponentPage.Files => this.FilesPageLabel,
            _ => this.LinksPageLabel,
        };

        private void ShowCurrentPage()
        {
            this.BasicsSection.IsVisible = this.thisPage == ComponentPage.Basics;
            this.ReplacementsSection.IsVisible = this.thisPage == ComponentPage.Replacements;
            this.PinoutSection.IsVisible = this.thisPage == ComponentPage.Pinout;
            this.PicturesSection.IsVisible = this.thisPage == ComponentPage.Pictures;
            this.BaselinesSection.IsVisible = this.thisPage == ComponentPage.Baselines;
            this.FilesSection.IsVisible = this.thisPage == ComponentPage.Files;
            this.LinksSection.IsVisible = this.thisPage == ComponentPage.Links;

            // Which row of its Basics this is - the other pages' rows say their own region.
            this.RegionText.IsVisible = this.thisPage == ComponentPage.Basics && !string.IsNullOrEmpty(this.RegionText.Text);

            foreach (ComponentPage page in Enum.GetValues<ComponentPage>())
            {
                this.PageButton(page).Classes.Set("Selected", page == this.thisPage);
            }

            this.ShowPageTabs();
        }

        // ###########################################################################################
        // The tabs, while a component's pages are what shows - each with how many it holds, and marked
        // when something on it stops the draft being submitted. The line under the header gives way to
        // the strip's own.
        // ###########################################################################################
        private void ShowPageTabs()
        {
            bool shown = this.SectionsScrollViewer.IsVisible && this.thisBoardPart is null && !this.thisIsNew;

            this.PageTabs.IsVisible = shown;
            this.PaneHeader.BorderThickness = shown ? default : new Thickness(0, 0, 0, 1);

            if (!shown)
            {
                return;
            }

            this.ShowPageTab(ComponentPage.Basics, 0, string.IsNullOrWhiteSpace(this.CategoryBox.Text));
            this.ShowPageTab(ComponentPage.Replacements, this.thisReplacements.Count, hasError: false);
            this.ShowPageTab(ComponentPage.Pinout, this.thisPins.Count, this.thisPins.Any(row => row.PinOutside));
            this.ShowPageTab(ComponentPage.Pictures, this.thisPictures.Count, hasError: false);
            this.ShowPageTab(ComponentPage.Baselines, this.thisBaselines.Count, this.thisBaselines.Any(row => row.PinOutside || row.LacksWhatItNeeds));
            this.ShowPageTab(ComponentPage.Files, this.thisFiles.Count, hasError: false);
            this.ShowPageTab(ComponentPage.Links, this.thisLinks.Count, this.thisLinks.Any(row => row.UrlNotWeb));
        }

        private void ShowPageTab(ComponentPage page, int count, bool hasError)
        {
            TextBlock label = this.PageLabel(page);

            label.Text = ComponentPaneWording.PageTitle(page, count);
            label.Classes.Set("error", hasError);
            ToolTip.SetTip(this.PageButton(page), hasError ? ComponentPaneWording.PageHasErrorTip : null);
        }

        // ---- the region shown ----------------------------------------------------------------------

        private void InitializeRegionSwitches()
        {
            foreach (PaneRegionSwitch regionSwitch in this.RegionSwitches())
            {
                regionSwitch.RegionPicked += (_, region) => this.OnPaneRegionPicked(region);
            }
        }

        private PaneRegionSwitch[] RegionSwitches() =>
        [
            this.BasicsRegionSwitch,
            this.PinsRegionSwitch,
            this.ReplacementsRegionSwitch,
            this.PicturesRegionSwitch,
            this.BaselinesRegionSwitch,
            this.FilesRegionSwitch,
            this.LinksRegionSwitch,
        ];

        private void OnPaneRegionPicked(string region)
        {
            if (string.Equals(region, this.thisRegion, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            this.UseRegion(region);
            this.RegionPicked?.Invoke(this, region);
        }

        // ###########################################################################################
        // The region CRT shows, from either switch: what the boxes hold goes in under the region it was
        // typed in, then the pages read that region's rows.
        // ###########################################################################################
        internal void UseRegion(string? region)
        {
            string wanted = region?.Trim() ?? string.Empty;

            if (string.Equals(wanted, this.thisRegion, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (this.thisSession is null)
            {
                this.thisRegion = wanted;
                return;
            }

            this.CommitHeld();
            this.thisRegion = wanted;
            this.ShowComponent();
        }

        private void UseBoardRegions(BoardData board) => this.thisBoardHasRegions = BoardRegions.Has(board);

        // The ticks with their switches, and the "Applies to" lists, as the board has regions or not.
        private void ShowRegionChoices(BoardData board)
        {
            bool has = BoardRegions.Has(board);
            bool basics = ComponentPaneEdits.IsRegional(board, this.thisLabel);
            bool pins = ComponentPaneEdits.PinsAreRegional(board, this.thisLabel);

            this.UseBoardRegions(board);

            this.BasicsRegionalCheck.IsVisible = has;
            this.BasicsRegionalCheck.IsChecked = basics;
            this.BasicsRegionSwitch.Show(has && basics, this.thisRegion);

            this.PinsRegionalCheck.IsVisible = has;
            this.PinsRegionalCheck.IsChecked = pins;
            this.PinsRegionSwitch.Show(has && pins, this.thisRegion);

            // The pages whose rows each say what they apply to need no tick - a row is moved by its own.
            this.ReplacementsRegionSwitch.Show(has, this.thisRegion);
            this.PicturesRegionSwitch.Show(has, this.thisRegion);
            this.BaselinesRegionSwitch.Show(has, this.thisRegion);
            this.FilesRegionSwitch.Show(has, this.thisRegion);
            this.LinksRegionSwitch.Show(has, this.thisRegion);

            this.FollowRegionChoices(has);
        }

        // Every row of the page offers "Applies to" while the board has regions - new rows too - and the
        // lists with a heading over their columns say it there.
        private void FollowRegionChoices(bool has)
        {
            foreach (ComponentPaneRegionalRow row in this.thisBaselines.Cast<ComponentPaneRegionalRow>()
                .Concat(this.thisPictures)
                .Concat(this.thisReplacements)
                .Concat(this.thisFiles)
                .Concat(this.thisLinks))
            {
                row.ShowsAppliesTo = has;
            }

            this.FilesAppliesToHeading.IsVisible = has;
            this.LinksAppliesToHeading.IsVisible = has;
        }

        // ###########################################################################################
        // A row's "Applies to" moved it out of the region shown: it is in the draft now (the commit
        // just made), and no row of this region's page any more - left here it would be filed again
        // beside itself at the next commit, as the rows of the region shown are what a commit replaces.
        //
        // *** BY WHERE THE DRAFT HAS IT (code review, 2026-11-01). *** A baseline held back stayed in the
        // draft in its old region: let go of here for what its list says, the next commit - replacing that
        // region's rows with the page's - deleted it from the draft, under a "Moved to NTSC".
        // ###########################################################################################
        private void LetGoOfRowsOfOtherRegions()
        {
            string? movedTo = null;

            movedTo = this.LetGoOfRowsOfOtherRegions(this.thisBaselines) ?? movedTo;
            movedTo = this.LetGoOfRowsOfOtherRegions(this.thisPictures) ?? movedTo;
            movedTo = this.LetGoOfRowsOfOtherRegions(this.thisReplacements) ?? movedTo;
            movedTo = this.LetGoOfRowsOfOtherRegions(this.thisFiles) ?? movedTo;
            movedTo = this.LetGoOfRowsOfOtherRegions(this.thisLinks) ?? movedTo;

            if (movedTo is not null)
            {
                this.ShowMessage(ComponentPaneWording.MovedToRegion(movedTo));
                this.ShowEmptyLines();
            }
        }

        private string? LetGoOfRowsOfOtherRegions<T>(IList<T> rows)
            where T : ComponentPaneRegionalRow
        {
            string? movedTo = null;

            for (int index = rows.Count - 1; index >= 0; index--)
            {
                // A row nothing is typed in yet is in no draft to move: it stays until it holds something
                // (2026-11-01: "Applies to" picked first took a new row off the page and out of existence).
                if (!rows[index].HoldsNothingYet &&
                    !ComponentPaneEdits.InRegion(rows[index].RegionInDraft, this.thisRegion) &&
                    !this.IsWaitingToBeTyped(rows[index]))
                {
                    movedTo = rows[index].Region;
                    rows.RemoveAt(index);
                }
            }

            return movedTo;
        }

        // ###########################################################################################
        // *** A NEW ROW GIVEN ANOTHER REGION WAITS UNTIL IT HAS BEEN TYPED IN (code review, 2026-11-02). ***
        // "Add link", "NTSC" picked, a name typed and a pause: into the draft half typed and off this
        // region's page, the box gone from under the caret, and no way there to type the address. So it
        // waits - out of the draft and on the page - through the typing pause and the Tab to its next box,
        // and moves as the focus leaves the row. Whatever else commits takes it as it is (RowsWaiting.None).
        //
        // Only such a row can be here: an existing row moves the moment its region is picked, and a new
        // part, file or link is the only row that can stay behind for holding nothing (HoldsNothingYet). A
        // baseline or a picture cannot wait - CommitBaselines has already filed it.
        // ###########################################################################################
        private enum RowsWaiting
        {
            None,
            TheOneInFocus,
            AllOfAnotherRegion,
        }

        private RowsWaiting thisRowsWaiting;

        private bool IsWaitingToBeTyped(ComponentPaneRegionalRow row)
        {
            if (this.thisRowsWaiting == RowsWaiting.None ||
                row is not (ComponentPaneReplacementRow or ComponentPaneFileRow or ComponentPaneLinkRow) ||
                ComponentPaneEdits.InRegion(row.Region, this.thisRegion))
            {
                return false;
            }

            return this.thisRowsWaiting == RowsWaiting.AllOfAnotherRegion ||
                   (TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() is StyledElement focused && ReferenceEquals(focused.DataContext, row));
        }

        // The region a new row of the kind that mostly is per region starts in: the one shown.
        private string RegionForNewBaseline() =>
            this.thisSession is { } session && BoardRegions.Has(session.Content.Board) ? this.thisRegion : string.Empty;

        // ###########################################################################################
        // A new pin joins the set it stands with: the region's own while the pins are told apart, else
        // the one for every region - but a component that is in one region only: its pins are that
        // region's, the region it is worked on in (ShowComponent).
        // ###########################################################################################
        private string RegionForNewPin()
        {
            if (this.thisSession is not { } session)
            {
                return string.Empty;
            }

            BoardData board = session.Content.Board;

            if (ComponentPaneEdits.PinsAreRegional(board, this.thisLabel))
            {
                return this.thisRegion;
            }

            List<ComponentEntry> rows = board.Components
                .Where(row => string.Equals(row.BoardLabel?.Trim(), this.thisLabel, StringComparison.OrdinalIgnoreCase))
                .ToList();

            return rows.Count == 1 ? rows[0].Region?.Trim() ?? string.Empty : string.Empty;
        }

        private void OnBasicsRegionalClick(object? sender, RoutedEventArgs e) =>
            this.SetRegional(
                this.BasicsRegionalCheck.IsChecked == true,
                DraftStepWording.ComponentRegional,
                ComponentPaneEdits.WithRegionalRows,
                "Its fields are");

        private void OnPinsRegionalClick(object? sender, RoutedEventArgs e) =>
            this.SetRegional(
                this.PinsRegionalCheck.IsChecked == true,
                DraftStepWording.PinsRegional,
                ComponentPaneEdits.WithRegionalPins,
                "Its pins are",
                (regional, region) => regional ? ComponentPaneWording.NoPinsToTellApart : ComponentPaneWording.NoPinsInRegion(region));

        // ###########################################################################################
        // One tick's change, as one step. Not the pane's own edit: the pane reads the component back.
        // Nothing changed - no pins to tell apart, or none in the region shown to make every region's - and
        // the tick goes back to what the draft says, saying why (`whyNothingChanged`): no step means
        // nothing reads it back.
        // ###########################################################################################
        private void SetRegional(bool regional, Func<string, bool, string> step, Func<BoardData, string?, string?, bool, BoardData> edit, string what, Func<bool, string, string>? whyNothingChanged = null)
        {
            if (this.thisSession is not { } session || this.thisLabel.Length == 0)
            {
                return;
            }

            this.CommitHeld();
            this.ShowMessage(null);

            string label = this.thisLabel;
            string region = this.thisRegion;

            bool changed = session.Apply(
                step(label, regional),
                content => edit(content.Board, label, region, regional) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null);

            if (changed)
            {
                if (!regional)
                {
                    this.ShowMessage(ComponentPaneWording.MadeTheSameInEveryRegion(what, region));
                }

                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                this.ShowRegionChoices(session.Content.Board);

                if (whyNothingChanged is not null)
                {
                    this.ShowMessage(whyNothingChanged(regional, region));
                }
            }
        }

        // ---- its label -----------------------------------------------------------------------------

        // Escape puts the label back; Enter is every box's (OnPaneKeyDown).
        private void OnBoardLabelKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.BoardLabelBox.Text = this.thisLabelShown;
                this.ShowLabelProblem(null);
            }
        }

        private void ShowLabelProblem(string? problem)
        {
            this.BoardLabelProblemText.Text = problem ?? string.Empty;
            this.BoardLabelProblemText.IsVisible = problem is not null;
        }

        // ###########################################################################################
        // The label typed, into the draft as one step - every row of the component with it. One that
        // cannot be taken stays in its box, the reason under it, and nothing changes.
        // ###########################################################################################
        private void CommitLabel()
        {
            if (this.thisSession is not { } session || this.thisLabel.Length == 0)
            {
                return;
            }

            string typed = ComponentPaneEdits.Normalised(ComponentField.FriendlyName, this.BoardLabelBox.Text);

            if (string.Equals(typed, this.thisLabelShown, StringComparison.Ordinal))
            {
                this.ShowLabelProblem(null);
                return;
            }

            string from = this.thisLabel;

            string? problem = ComponentPaneWording.LabelProblem(
                typed,
                session.Content.Board.Components
                    .Select(row => row.BoardLabel)
                    .Where(label => !string.Equals(label?.Trim(), from, StringComparison.OrdinalIgnoreCase)));

            this.ShowLabelProblem(problem);

            if (problem is not null)
            {
                return;
            }

            // Named before the step: what follows it - the left column, "Draft" - goes by the new label.
            this.thisLabel = typed;

            this.Apply(DraftStepWording.ComponentRelabelled(from, typed), board => ComponentPaneEdits.WithLabel(board, from, typed));

            this.ShowRelabelled();
            this.RaiseComponentShown();
        }

        // ###########################################################################################
        // What goes by the label, read back after a rename - NOT the rows (code review, 2026-11-01): read
        // again, a pin just added and still blank and a baseline's edit held back, its reason under it,
        // went without a word. The rows carry no label of their own; a commit files them under this one.
        // ###########################################################################################
        private void ShowRelabelled()
        {
            if (this.thisSession is not { } session ||
                ComponentPaneEdits.RowShown(session.Content.Board, this.thisLabel, this.thisRegion) is not { } row)
            {
                this.ShowComponent();
                return;
            }

            this.thisShowing = true;

            try
            {
                this.thisLabelShown = row.BoardLabel?.Trim() ?? string.Empty;
                this.BoardLabelBox.Text = this.thisLabelShown;
                this.TitleText.Text = ComponentPaneWording.Title(row);
                this.ShowActions(session.Content.Board);
                this.ShowMigrationOffer(session.Content.Board);
            }
            finally
            {
                this.thisShowing = false;
            }

            this.ShowChangedState();
        }

        // ###########################################################################################
        // The component an undo or a redo of a rename left under another label: the step ITSELF says where
        // it went, by the one rule of what a rename is (ComponentRelabel.Between - a row whose label alone
        // changed). Nothing renamed, it really is gone - and the pane says so.
        //
        // *** NOT A RECORD OF RENAMES (code review, 2026-11-02). *** One was kept here, and no record could
        // tell the cases apart: U1 renamed U2, a new U1 added, U2 deleted, and it walked back to the new U1,
        // a different component; told to forget a label given out, it lost the way back of U1 to U2 to U3 to
        // U2, and of a new U1 undone again. Which component a label means depends on where in its history the
        // draft stands, and each step says exactly that.
        // ###########################################################################################
        private void FollowRelabelInDraft(IReadOnlyList<ComponentEntry>? before, BoardData board)
        {
            if (this.thisLabel.Length > 0 &&
                ComponentPaneEdits.RowShown(board, this.thisLabel, this.thisRegion) is null &&
                ComponentRelabel.Between(before, board.Components).TryGetValue(this.thisLabel, out string? now))
            {
                this.thisLabel = now;
            }
        }

        // ---- what Submit would refuse, marked where it is ----------------------------------------------

        // No component may have an empty category (owner decision, 2026-11-01) - an error, like a field a
        // row of the board's lists needs.
        private void ShowCategoryProblem()
        {
            bool empty = string.IsNullOrWhiteSpace(this.CategoryBox.Text);

            this.CategoryBox.Classes.Set("Missing", empty);
            this.CategoryRequiredText.IsVisible = empty;
        }

        // ###########################################################################################
        // A pin the component's case does not have, on each row that names one (owner request,
        // 2026-11-01: "if I e.g. put pin 48 on a 24 pin package, this is an error") - by the checks' own
        // rule, against the draft as it reads now: the package typed is committed before this runs.
        // ###########################################################################################
        private void ShowPinProblems()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            List<ComponentEntry> components = session.Content.Board.Components;

            foreach (ComponentPanePinRow row in this.thisPins)
            {
                row.UseCaseThatRefuses(BoardDataChecks.PackageRefusing(components, this.thisLabel, row.Region, row.Pin));
            }

            foreach (ComponentPaneBaselineRow row in this.thisBaselines)
            {
                row.UseCaseThatRefuses(BoardDataChecks.PackageRefusing(components, this.thisLabel, row.Region, row.Pin));
            }
        }

        // ---- TEMPORARY: pinouts made of scope baselines --------------------------------------------------
        // *** REMOVED BEFORE CRT 4.0 GOES LIVE - see BaselinePinMigration, whose test fails a release. ***

        private void ShowMigrationOffer(BoardData board)
        {
            this.MigrationNoteText.Text = ComponentPaneWording.MigrationNote;
            this.MigrationPanel.IsVisible = board.ComponentImages.Any(ComponentImageKinds.IsScopeBaseline);
            this.MigrateComponentButton.IsEnabled = BaselinePinMigration.CanRun(board, this.thisLabel);
        }

        private void OnMigrateComponentClick(object? sender, RoutedEventArgs e) => this.MigrateComponent();

        private void OnMigrateBoardClick(object? sender, RoutedEventArgs e) => this.MigrateBoard();

        internal void MigrateComponent()
        {
            if (this.thisSession is not { } session || this.thisLabel.Length == 0)
            {
                return;
            }

            this.CommitHeld();
            this.ShowMessage(null);

            string label = this.thisLabel;

            bool changed = session.Apply(
                DraftStepWording.PinsFromBaselines(label),
                content => BaselinePinMigration.Apply(content.Board, label) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null);

            if (changed)
            {
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                this.ShowMessage(ComponentPaneWording.MigrationNothing);
            }
        }

        internal void MigrateBoard()
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            this.CommitHeld();

            BaselinePinMigration.ApplyToBoard(session.Content.Board, out int components);

            bool changed = components > 0 && session.Apply(
                DraftStepWording.PinsFromBaselinesOfBoard(components),
                content => BaselinePinMigration.ApplyToBoard(content.Board, out _) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null);

            this.ShowMessage(changed ? ComponentPaneWording.MigratedBoard(components) : ComponentPaneWording.MigrationNothing);

            if (changed)
            {
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }

    // ###########################################################################################
    // A row that can apply to one region or to all (owner decisions, 2026-11-01 and 2026-11-02): a
    // component's baseline, picture, replacement part, file or link, and the board's files and links (the
    // list rows' base - a credit and an important signal never say one). The list's empty write is no
    // change, as a pin's direction.
    //
    // *** THE BOX'S LIST STAYS WHILE IT HOLDS THE WORD (code review, 2026-11-02). *** It was made anew
    // whenever the region changed - which a pick is - and a list replaced in the middle of a pick makes the
    // box put the item before it back and write that: no row's region could be changed in the pane. So a
    // new list only for a word it does not hold, which no pick can be. A row's word that is no region stays
    // listed until the row is shown again.
    // ###########################################################################################
    public abstract class ComponentPaneRegionalRow : ComponentPaneDraggableRow
    {
        private string thisRegion = string.Empty;
        private IReadOnlyList<string> thisAppliesToChoices = ComponentPaneWording.AppliesToChoices(null);
        private bool thisShowsAppliesTo;

        public string Region
        {
            get => this.thisRegion;
            set
            {
                if (this.Set(ref this.thisRegion, value))
                {
                    if (!this.thisAppliesToChoices.Contains(ComponentPaneWording.AppliesToOf(value), StringComparer.Ordinal))
                    {
                        this.thisAppliesToChoices = ComponentPaneWording.AppliesToChoices(value);
                        this.Raise(nameof(this.AppliesToChoices));
                    }

                    this.Raise(nameof(this.AppliesTo));
                }
            }
        }

        // The region the draft has this row in, once the pane has committed - what decides whether it is
        // still a row of the page.
        public virtual string RegionInDraft => this.Region;

        // A row just added with nothing in it - no row of the draft.
        public virtual bool HoldsNothingYet => false;

        public bool ShowsAppliesTo
        {
            get => this.thisShowsAppliesTo;
            set
            {
                if (this.thisShowsAppliesTo != value)
                {
                    this.thisShowsAppliesTo = value;
                    this.Raise(nameof(this.ShowsAppliesTo));
                }
            }
        }

        public IReadOnlyList<string> AppliesToChoices => this.thisAppliesToChoices;

        public string? AppliesTo
        {
            get => ComponentPaneWording.AppliesToOf(this.thisRegion);
            set
            {
                if (value is not null)
                {
                    this.Region = ComponentPaneWording.RegionOfAppliesTo(value);
                }
            }
        }
    }
}
