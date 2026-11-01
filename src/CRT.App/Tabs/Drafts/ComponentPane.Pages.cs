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
    //   - THE TABS: Basics, Replacement parts, Pinout, Pictures, Scope baseline, Files and Links, each
    //     saying how many it holds. One page shows; every section stays filled behind it, so what the
    //     pane commits is what it always was. The page stays as another component is picked: walking a
    //     board's chips stays on one page.
    //   - THE REGION SHOWN, on a board that has regions (BoardRegions): the switch under the component
    //     list, again (RegionPicked - Main's, so all of CRT follows). What a page lists is what that
    //     region shows: its own rows and those for every region. A row's "Applies to" moves it (one
    //     leaving the region shown is let go of here, and said), and Basics and Pinout have a "Different
    //     per region" tick instead - a set of rows that belong together.
    //   - ITS LABEL (owner report, 2026-11-01: "unable to rename the component label"): on Enter or when
    //     its box is left, never on a pause in typing - every row of the component moves with it.
    //   - WHAT SUBMIT WOULD REFUSE, MARKED WHERE IT IS: no category; a pin the case does not have.
    //   - THE TEMPORARY TOOL making pinouts of scope baselines (BaselinePinMigration - gone before CRT 4.0
    //     goes live, with this section).
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        private ComponentPage thisPage = ComponentPage.Basics;

        // Every rename made in the pane: an undo takes one back, and the component is found again by it.
        private readonly List<(string From, string To)> thisRelabels = new();

        // The label as the box last read it - its capitals are a change of their own.
        private string thisLabelShown = string.Empty;

        // Asked to show another region - Main's, so all of CRT follows.
        public event EventHandler<string>? RegionPicked;

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
            this.ShowPageTab(ComponentPage.Baselines, this.thisBaselines.Count, this.thisBaselines.Any(row => row.PinOutside));
            this.ShowPageTab(ComponentPage.Files, this.thisFiles.Count, hasError: false);
            this.ShowPageTab(ComponentPage.Links, this.thisLinks.Count, hasError: false);
        }

        private void ShowPageTab(ComponentPage page, int count, bool hasError)
        {
            TextBlock label = this.PageLabel(page);

            label.Text = ComponentPaneWording.PageTitle(page, count);
            label.Classes.Set("error", hasError);
            ToolTip.SetTip(this.PageButton(page), hasError ? ComponentPaneWording.PageHasErrorTip : null);
        }

        // ---- the region shown ----------------------------------------------------------------------

        private void OnPaneRegionClick(object? sender, RoutedEventArgs e)
        {
            string region = ReferenceEquals(sender, this.PaneRegionNtscButton) ? BoardRegions.Ntsc : BoardRegions.Pal;

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

        // The switch, the "Applies to" lists and the ticks, as the board has regions or not.
        private void ShowRegionChoices(BoardData board)
        {
            bool has = BoardRegions.Has(board);

            this.RegionSwitchPanel.IsVisible = has;
            this.PaneRegionPalButton.Classes.Set("active", string.Equals(this.thisRegion, BoardRegions.Pal, StringComparison.OrdinalIgnoreCase));
            this.PaneRegionNtscButton.Classes.Set("active", string.Equals(this.thisRegion, BoardRegions.Ntsc, StringComparison.OrdinalIgnoreCase));

            this.BasicsRegionalCheck.IsVisible = has;
            this.BasicsRegionalCheck.IsChecked = ComponentPaneEdits.IsRegional(board, this.thisLabel);
            this.PinsRegionalCheck.IsVisible = has;
            this.PinsRegionalCheck.IsChecked = ComponentPaneEdits.PinsAreRegional(board, this.thisLabel);

            this.FollowRegionChoices(has);
        }

        // Every row of the page offers "Applies to" while the board has regions - new rows too.
        private void FollowRegionChoices(bool has)
        {
            foreach (ComponentPaneRegionalRow row in this.thisBaselines.Cast<ComponentPaneRegionalRow>().Concat(this.thisPictures).Concat(this.thisReplacements))
            {
                row.ShowsAppliesTo = has;
            }
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

            movedTo = ComponentPane.LetGoOfRowsOfOtherRegions(this.thisBaselines, this.thisRegion) ?? movedTo;
            movedTo = ComponentPane.LetGoOfRowsOfOtherRegions(this.thisPictures, this.thisRegion) ?? movedTo;
            movedTo = ComponentPane.LetGoOfRowsOfOtherRegions(this.thisReplacements, this.thisRegion) ?? movedTo;

            if (movedTo is not null)
            {
                this.ShowMessage(ComponentPaneWording.MovedToRegion(movedTo));
                this.ShowEmptyLines();
            }
        }

        private static string? LetGoOfRowsOfOtherRegions<T>(IList<T> rows, string region)
            where T : ComponentPaneRegionalRow
        {
            string? movedTo = null;

            for (int index = rows.Count - 1; index >= 0; index--)
            {
                if (!ComponentPaneEdits.InRegion(rows[index].RegionInDraft, region))
                {
                    movedTo = rows[index].Region;
                    rows.RemoveAt(index);
                }
            }

            return movedTo;
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

            // Named before the step: what follows it - the left column, "Changed" - goes by the new label.
            this.thisLabel = typed;
            this.thisRelabels.Add((from, typed));

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
        // The component an undo or a redo of a rename left under another label: the next one it has
        // been known by that the draft has. Nothing found, it really is gone - and the pane says so.
        // ###########################################################################################
        private void FollowRelabelInDraft(BoardData board)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { this.thisLabel };
            var waiting = new Queue<string>(seen);

            while (waiting.Count > 0)
            {
                string label = waiting.Dequeue();

                if (ComponentPaneEdits.RowShown(board, label, this.thisRegion) is not null)
                {
                    this.thisLabel = label;
                    return;
                }

                foreach ((string from, string to) in this.thisRelabels)
                {
                    if (string.Equals(from, label, StringComparison.OrdinalIgnoreCase) && seen.Add(to))
                    {
                        waiting.Enqueue(to);
                    }

                    if (string.Equals(to, label, StringComparison.OrdinalIgnoreCase) && seen.Add(from))
                    {
                        waiting.Enqueue(from);
                    }
                }
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
    // A row of a component that can apply to one region or to all (owner decision, 2026-11-01): a
    // baseline, a picture, a replacement part. The list's empty write is no change, as a pin's direction.
    // ###########################################################################################
    public abstract class ComponentPaneRegionalRow : ComponentPaneRowBase
    {
        private string thisRegion = string.Empty;
        private bool thisShowsAppliesTo;

        public string Region
        {
            get => this.thisRegion;
            set
            {
                if (this.Set(ref this.thisRegion, value))
                {
                    this.Raise(nameof(this.AppliesTo));
                    this.Raise(nameof(this.AppliesToChoices));
                }
            }
        }

        // The region the draft has this row in, once the pane has committed - what decides whether it is
        // still a row of the page.
        public virtual string RegionInDraft => this.Region;

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

        public IReadOnlyList<string> AppliesToChoices => ComponentPaneWording.AppliesToChoices(this.thisRegion);

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
