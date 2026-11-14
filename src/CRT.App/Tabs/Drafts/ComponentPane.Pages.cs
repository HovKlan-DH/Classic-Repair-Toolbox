using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // A COMPONENT'S PAGES (owner decision, 2026-11-01: "the functionality in this is too massive, if all
    // data is filled, so I think we need to split this into multiple menu items"; tabs over the pane until
    // 2026-11-03 - owner request: the left column has the room for them now, "and no tabs in the 'component
    // view'" - TabDrafts.Entries.cs).
    //
    // Owns:
    //   - THE PAGES: Basics, Replacement parts, Pinout, Pictures, Oscilloscope baseline, Files and Links,
    //     what each holds told to the host (PagesShown). One page shows; every section stays filled behind
    //     it, so what the pane commits is what it always was. The page stays as another component is
    //     picked: walking a board's chips stays on one page.
    //   - THE REGION SHOWN, on a board that has regions (BoardRegions). What a page lists is what that
    //     region shows: its own rows and those for every region. A row's "Applies to" moves it (one
    //     leaving the region shown is let go of here, and said), and Basics and Pinout have a "This data
    //     is different per region" tick instead - a set of rows that belong together. THE BUTTONS SWITCHING
    //     THE REGION ARE THE ONES UNDER CRT'S COMPONENT LIST (owner request, 2026-11-03: "use the regions
    //     from the left-side instead of adding new region buttons in the component pages"); a line says
    //     which region a page is showing (ShowRegionShown).
    //   - FILES AND LINKS ARE THE COMPONENT'S WHOLE LISTS (owner request, 2026-11-03: "Each file
    //     or link will have set a region (if any) and then the complete list can be viewed and sorted. So,
    //     it should work in same was as 'Board: Files' and 'Board: Links'"): every row whatever region CRT
    //     shows, each saying what it applies to, so one changing region stays where it is.
    //   - ITS LABEL (owner report, 2026-11-01: "unable to rename the component label"): on Enter or when
    //     its box is left, never on a pause in typing - every row of the component moves with it.
    //   - WHAT SUBMIT WOULD REFUSE, MARKED WHERE IT IS AND ON ITS ENTRY: no category; a pin the case does not
    //     have; a baseline lacking its signal name or image; a link's address that is no web address.
    //
    // The temporary buttons making pinouts of scope baselines and KiCad data went on 2026-11-06 (owner
    // request): the data tree's conversion makes every board's pinouts now (ConversionPinouts).
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

        // The regions of the board open (BoardRegions.Of) - what offers "Applies to" and the ticks, and what they list.
        private IReadOnlyList<string> thisBoardRegions = [];

        private bool thisBoardHasRegions => this.thisBoardRegions.Count > 0;

        internal ComponentPage Page => this.thisPage;

        // The page asked for, what the boxes hold handed to the draft first, at its top.
        internal void ShowPage(ComponentPage page)
        {
            this.CommitHeld();
            this.thisPage = page;
            this.ShowCurrentPage();
            this.SectionsScrollViewer.Offset = default;
        }

        private void ShowCurrentPage()
        {
            this.BasicsSection.IsVisible = this.thisPage == ComponentPage.Basics;
            this.ReplacementsSection.IsVisible = this.thisPage == ComponentPage.Replacements;
            this.PinoutSection.IsVisible = this.thisPage == ComponentPage.Pinout;
            this.PinsArea.IsVisible = true;
            this.PicturesSection.IsVisible = this.thisPage == ComponentPage.Pictures;
            this.BaselinesSection.IsVisible = this.thisPage == ComponentPage.Baselines;
            this.FilesSection.IsVisible = this.thisPage == ComponentPage.Files;
            this.LinksSection.IsVisible = this.thisPage == ComponentPage.Links;

            // Which row of its Basics this is - the other pages' rows say their own region.
            this.RegionText.IsVisible = this.thisPage == ComponentPage.Basics && BoldRuns.TextOf(this.RegionText).Length > 0;
            this.ShowRegionShown();
            this.HoldBackPagesOfOne();

            this.TellPagesChanged();
        }

        // What a page holds may have changed - the left column reads PagesShown.
        internal event EventHandler? PagesChanged;

        private void TellPagesChanged() => this.PagesChanged?.Invoke(this, EventArgs.Empty);

        // ###########################################################################################
        // What each page holds and whether something on it stops Submit, while a component's pages are
        // what shows - from the rows on screen, which know what is held back and what is still blank. Null
        // otherwise. Basics holds the component's rows: one per region it differs in.
        // ###########################################################################################
        internal IReadOnlyDictionary<ComponentPage, ComponentPageFact>? PagesShown
        {
            get
            {
                if (!this.SectionsScrollViewer.IsVisible || this.thisBoardPart is not null || this.thisIsNew || this.thisSession is not { } session)
                {
                    return null;
                }

                int rows = this.IsSeveral
                    ? this.thisSeveral.Count
                    : session.Content.Board.Components.Count(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, this.thisLabel));

                return new Dictionary<ComponentPage, ComponentPageFact>
                {
                    [ComponentPage.Basics] = new(rows, this.IsCategoryMissing),
                    [ComponentPage.Replacements] = new(this.thisReplacements.Count, false),
                    [ComponentPage.Pinout] = new(this.thisPins.Count, this.thisPins.Any(row => row.PinOutside)),
                    [ComponentPage.Pictures] = new(this.thisPictures.Count, false),
                    [ComponentPage.Baselines] = new(this.thisBaselines.Count, this.thisBaselines.Any(row => row.PinOutside || row.LacksWhatItNeeds)),
                    [ComponentPage.Files] = new(this.thisFiles.Count, false),
                    [ComponentPage.Links] = new(this.thisLinks.Count, this.thisLinks.Any(row => row.UrlNotWeb)),
                };
            }
        }

        // ---- the region shown ----------------------------------------------------------------------

        // ###########################################################################################
        // The region CRT shows, from the buttons under its component list: what the boxes hold goes in
        // under the region it was typed in, then the pages read that region's rows.
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
            this.FollowRegionInClonePanel();
        }

        internal string RegionShownForTests => this.thisRegion;

        private void UseBoardRegions(BoardData board) => this.thisBoardRegions = BoardRegions.Of(board);

        // What the two ticks said as last shown - all of them, none, or (null) some of several.
        private bool? thisBasicsRegionalShown;
        private bool? thisPinsRegionalShown;

        // ###########################################################################################
        // The ticks and the "Applies to" lists, as the board has regions or not. WITH SEVERAL PICKED A TICK SAYS ALL OF THEM,
        // NONE, OR SOME - half-ticked (owner request, 2026-11-09: "Can't it still have this, even if selecting multiple
        // components?"); hidden until then.
        // ###########################################################################################
        private void ShowRegionChoices(BoardData board)
        {
            bool has = BoardRegions.Has(board);

            this.UseBoardRegions(board);

            this.thisBasicsRegionalShown = this.AllOrNone(board, ComponentPaneEdits.IsRegional);
            this.BasicsRegionalCheck.IsVisible = has;
            this.BasicsRegionalCheck.IsChecked = this.thisBasicsRegionalShown;

            // The one box of Basics that is never per region: it is what every row of the component is
            // filed under, so it is said (owner report, 2026-11-04).
            this.BoardLabelOneText.Text = ComponentPaneWording.LabelInEveryRegion;
            this.BoardLabelOneText.IsVisible = has && !this.IsSeveral && this.BasicsRegionalCheck.IsChecked == true;

            this.thisPinsRegionalShown = this.AllOrNone(board, ComponentPaneEdits.PinsAreRegional);

            // Not over pins held back - several whose packages differ (HoldBackPagesOfOne).
            this.PinsRegionalCheck.IsVisible = has && !(this.IsSeveral && this.PackagesDiffer(board));
            this.PinsRegionalCheck.IsChecked = this.thisPinsRegionalShown;

            this.ShowRegionShown();
            this.FollowRegionChoices(has);
        }

        // `test` of every component shown: true or false when all say so, null when they differ.
        private bool? AllOrNone(BoardData board, Func<BoardData, string?, bool> test)
        {
            List<bool> each = this.Labels.Select(label => test(board, label)).Distinct().ToList();

            return each.Count switch
            {
                0 => false,
                1 => each[0],
                _ => null,
            };
        }

        // ###########################################################################################
        // Which region the page shows - on a page that lists what the region shows: Basics and Pinout while
        // ticked (unticked they are one set, for every region), and the three whose rows each say what they
        // apply to. Files and Links list every row, so nothing there depends on it.
        // ###########################################################################################
        private void ShowRegionShown()
        {
            // Half-ticked - some of several told apart - the page is the region's for those.
            bool depends = this.thisBoardHasRegions && this.thisPage switch
            {
                ComponentPage.Basics => this.BasicsRegionalCheck.IsChecked != false,
                ComponentPage.Pinout => this.PinsRegionalCheck.IsChecked != false,
                ComponentPage.Replacements or ComponentPage.Pictures or ComponentPage.Baselines => true,
                _ => false,
            };

            BoldRuns.ShowMarked(this.RegionShownText, ComponentPaneWording.RegionShownLine(this.thisRegion, this.thisBoardRegions));
            this.RegionShownText.IsVisible = depends;
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
                row.UseRegions(this.thisBoardRegions);
            }

            this.FilesAppliesToHeading.IsVisible = has;
            this.LinksAppliesToHeading.IsVisible = has;
            ComponentPane.SizeAppliesToColumn(this.FilesAppliesToSizer, has, this.thisFiles);
            ComponentPane.SizeAppliesToColumn(this.LinksAppliesToSizer, has, this.thisLinks);
        }

        // ###########################################################################################
        // A list's "Applies to" column as wide as the widest choice any of its rows offers, made or not (code review,
        // 2026-11-14): a row's box is as wide as what it shows, and a column the rows share counts only the rows made - in a
        // list that makes only the rows in view, it widened and narrowed as the page scrolled. The heading's ColumnSizer
        // holds a box per choice.
        // ###########################################################################################
        internal static void SizeAppliesToColumn(ItemsControl sizer, bool shown, IEnumerable<ComponentPaneRegionalRow> rows)
        {
            List<string> choices = shown
                ? rows.SelectMany(row => row.AppliesToChoices).Distinct(StringComparer.Ordinal).ToList()
                : [];

            if (sizer.ItemsSource is not IEnumerable<string> held || !held.SequenceEqual(choices, StringComparer.Ordinal))
            {
                sizer.ItemsSource = choices;
            }

            sizer.IsVisible = choices.Count > 0;
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
        // "Add replacement part", "NTSC" picked, a part number typed and a pause: into the draft half typed
        // and off this region's page, the box gone from under the caret, and no way there to type the rest.
        // So it waits - out of the draft and on the page - through the typing pause and the Tab to its next
        // box, and moves as the focus leaves the row. Whatever else commits takes it as it is
        // (RowsWaiting.None).
        //
        // Only such a row can be here: an existing row moves the moment its region is picked, and a new
        // part is the only row that can stay behind for holding nothing (HoldsNothingYet). A baseline or a
        // picture cannot wait - CommitBaselines has already filed it - and a file or a link never leaves its
        // page, the component's whole list.
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
                row is not ComponentPaneReplacementRow ||
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

        // A new pin of the component shown joins the set it stands with (ComponentPaneEdits.RegionForNewPin) - with several
        // shown, each one's own, worked out as it is given to each (ComponentBulk.Pins).
        private string RegionForNewPin() =>
            this.thisSession is { } session ? ComponentPaneEdits.RegionForNewPin(session.Content.Board, this.thisLabel, this.thisRegion) : string.Empty;

        // A half-ticked box clicked ticks them all - its own toggle would untick them all.
        private void OnBasicsRegionalClick(object? sender, RoutedEventArgs e) =>
            this.SetRegional(
                this.thisBasicsRegionalShown is null || this.BasicsRegionalCheck.IsChecked == true,
                DraftStepWording.ComponentRegional,
                ComponentPaneEdits.WithRegionalRows,
                this.IsSeveral ? "Their fields are" : "Its fields are");

        private void OnPinsRegionalClick(object? sender, RoutedEventArgs e) =>
            this.SetRegional(
                this.thisPinsRegionalShown is null || this.PinsRegionalCheck.IsChecked == true,
                DraftStepWording.PinsRegional,
                ComponentPaneEdits.WithRegionalPins,
                this.IsSeveral ? "Their pins are" : "Its pins are",
                (regional, region) => regional ? ComponentPaneWording.NoPinsToTellApart : ComponentPaneWording.NoPinsInRegion(region));

        // ###########################################################################################
        // One tick's change, as one step - on every component shown, each as it would be on its own. Not the pane's own
        // edit: the pane reads the component back. Nothing changed - no pins to tell apart, or none in the region shown to
        // make every region's - and the tick goes back to what the draft says, saying why (`whyNothingChanged`): no step
        // means nothing reads it back.
        // ###########################################################################################
        private void SetRegional(bool regional, Func<string, bool, string> step, Func<BoardData, string?, string?, bool, BoardData> edit, string what, Func<bool, string, string>? whyNothingChanged = null)
        {
            if (this.thisSession is not { } session || this.Labels.Count == 0)
            {
                return;
            }

            this.CommitHeld();
            this.ShowMessage(null);

            IReadOnlyList<string> labels = this.Labels.ToList();
            string region = this.thisRegion;

            bool changed = session.Apply(
                step(this.StepLabel, regional),
                content => labels.Aggregate(content.Board, (board, label) => edit(board, label, region, regional)) is BoardData board &&
                           !ReferenceEquals(board, content.Board)
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
            // A label names one component - never several (ComponentPane.Several.cs).
            if (this.thisSession is not { } session || this.thisLabel.Length == 0 || this.IsSeveral)
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

        // ---- the package's list --------------------------------------------------------------------------

        // ###########################################################################################
        // The list opens with the focus, so what can be picked is seen at once (owner request, 2026-11-03) -
        // every case, as the box filled in already would narrow it to itself, and narrowing once typed in.
        // ###########################################################################################
        private string thisPackageTextAtFocus = string.Empty;

        // ###########################################################################################
        // Only the box ENTERED opens the list, never the focus moving inside it (2026-11-13, Avalonia 12.1.4 -> 12.1.6):
        // a press on a row of the list takes the focus to the row, and 12.1.6 keeps the list open for it, so the row's
        // GotFocus reaches the box, whose OnGotFocus takes the focus straight back to the text - closing the list
        // while the press is still picking. Opened again here, the list was made anew and the pick was lost: a package
        // could not be picked with the mouse. (In 12.1.4 the list closed as the row took the focus, and none of this
        // reached the box.)
        // ###########################################################################################
        private void OnPackageBoxFocused(object? sender, RoutedEventArgs e)
        {
            if (e is FocusChangedEventArgs { OldFocusedElement: Visual from } && this.IsInPackageBox(from))
                return;

            this.thisPackageTextAtFocus = this.PackageBox.Text ?? string.Empty;
            this.PackageBox.IsDropDownOpen = true;
        }

        // The box's text, or a row of its list - a popup of its own, outside the box's visual tree.
        private bool IsInPackageBox(Visual element)
        {
            if (this.PackageBox.IsVisualAncestorOf(element))
                return true;

            Control? list = this.PackageBox.GetVisualDescendants().OfType<Avalonia.Controls.Primitives.Popup>().FirstOrDefault()?.Child;

            return list is not null && (ReferenceEquals(list, element) || list.IsVisualAncestorOf(element));
        }

        private bool PackageShowsAll(string? search, object? item) =>
            string.Equals(search ?? string.Empty, this.thisPackageTextAtFocus, StringComparison.Ordinal) ||
            (item as string)?.Contains(search ?? string.Empty, StringComparison.OrdinalIgnoreCase) == true;

        // ---- what Submit would refuse, marked where it is ----------------------------------------------

        // No component may have an empty category (owner decision, 2026-11-01) - an error, like a field a
        // row of the board's lists needs.
        private void ShowCategoryProblem()
        {
            bool empty = this.IsCategoryMissing;

            this.CategoryBox.Classes.Set("Missing", empty);
            this.CategoryRequiredText.IsVisible = empty;
        }

        // ###########################################################################################
        // Whether the category box says the component has none - the one rule for the box's red and "Basics"' own (code review,
        // 2026-11-09: "Basics" went red for several whose categories differ, the box left empty on purpose). Several whose
        // categories differ: each has its own, so none is missing (ComponentPane.Several.cs).
        // ###########################################################################################
        private bool IsCategoryMissing =>
            string.IsNullOrWhiteSpace(this.BoxText(ComponentField.Category)) &&
            !(this.IsSeveral && this.thisDifferingValues.ContainsKey(ComponentField.Category));

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
        private IReadOnlyList<string> thisRegions = BoardRegions.PalAndNtsc;
        private IReadOnlyList<string> thisAppliesToChoices = ComponentPaneWording.AppliesToChoices(BoardRegions.PalAndNtsc, null);
        private bool thisShowsAppliesTo;

        public string Region
        {
            get => this.thisRegion;
            set
            {
                if (this.Set(ref this.thisRegion, value))
                {
                    if (!this.thisAppliesToChoices.Contains(ComponentPaneWording.AppliesToOf(this.thisRegions, value), StringComparer.Ordinal))
                    {
                        this.thisAppliesToChoices = ComponentPaneWording.AppliesToChoices(this.thisRegions, value);
                        this.Raise(nameof(this.AppliesToChoices));
                    }

                    this.Raise(nameof(this.AppliesTo));
                }
            }
        }

        // ###########################################################################################
        // The board's regions (BoardRegions.Of; 2026-11-09: names of its own since): "Applies to" shows while it has any,
        // and lists them. The list is made again only for regions that are not those it was made for.
        // ###########################################################################################
        public void UseRegions(IReadOnlyList<string> regions)
        {
            ArgumentNullException.ThrowIfNull(regions);

            if (!regions.SequenceEqual(this.thisRegions, StringComparer.Ordinal))
            {
                this.thisRegions = regions;
                this.thisAppliesToChoices = ComponentPaneWording.AppliesToChoices(regions, this.thisRegion);
                this.Raise(nameof(this.AppliesToChoices));
                this.Raise(nameof(this.AppliesTo));
            }

            this.ShowsAppliesTo = regions.Count > 0;
        }

        // The region the draft has this row in, once the pane has committed - what decides whether it is
        // still a row of the page.
        public virtual string RegionInDraft => this.Region;

        // A row just added with nothing in it - no row of the draft.
        public virtual bool HoldsNothingYet => false;

        public bool ShowsAppliesTo
        {
            get => this.thisShowsAppliesTo;
            private set
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
            get => ComponentPaneWording.AppliesToOf(this.thisRegions, this.thisRegion);
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
