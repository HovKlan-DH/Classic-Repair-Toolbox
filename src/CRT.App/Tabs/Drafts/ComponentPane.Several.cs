using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // SEVERAL COMPONENTS, EDITED AS ONE (owner request, 2026-11-08: "I want to be able to mark multiple components and when
    // doing that I am bulk editing whatever data I changes. E.g. selecting components "R21" and "R22" and I go to "Basics"
    // in UI view, then it should disable the "Board label" ... if I e.g. change "Category" or "Friendly name" fields, then
    // both components will get this new value"; the proposal agreed the same day). Every rule is CRT.Data's ComponentBulk;
    // this part shows them and routes the pane's commits through it.
    //
    //   - THE KHAKI BAR names them and says that every change reaches all of them; its second line says what only some have
    //     on a list page, or why a page is held back.
    //   - BASICS: the label is off (a label names ONE component). A field they share shows its value; one they differ in is
    //     empty, "Differs - N values" in it and their values listed under it while it has the focus - so nothing is set
    //     until something is typed or picked, and then it is set on all of them.
    //   - THE LIST PAGES (replacement parts, pinout, pictures, files, links) list the rows ALL of them have; a row changed,
    //     removed or added there is changed, removed or added on every one, and what only some have stays as it is. Their
    //     rows do not move here: an order is each component's own.
    //   - "MAKE ALL LIKE": the page shown made the same on every one as the one picked has it - ONE step, which Undo takes
    //     back. A button, not a mode that silently changes what later edits do.
    //   - PINOUT only while they share a package; OSCILLOSCOPE BASELINES not at all - a baseline is measured on one component.
    //     Neither offers "Make all like" then.
    //
    // Each commit is one step of the draft's history, named after them all. With one left - an undo took the others away -
    // the pane is that one's again.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The components shown together, two or more - empty while one, or none, is.
        private List<string> thisSeveral = [];

        internal bool IsSeveral => this.thisSeveral.Count > 1;

        // ###########################################################################################
        // *** WHETHER A ROW OF A COMPONENT'S LIST MAY MOVE - THE ONE RULE (code review, 2026-11-09: an IsSeveral check beside
        // every drag, move and button). *** An order is each component's own, so none moves while several are shown. Every
        // place that moves a row asks this - a drag's press, Alt+Up and Alt+Down (MoveRow), the pictures' Move up and Move
        // down - and a new one asks it too.
        // ###########################################################################################
        private bool RowsMove => !this.IsSeveral;

        // Every component shown - the one, or the several.
        internal IReadOnlyList<string> Labels =>
            this.IsSeveral ? this.thisSeveral : this.thisLabel.Length > 0 ? [this.thisLabel] : [];

        // What a step is named after: the component, or all of them - the first few and how many more (code review, 2026-11-09:
        // three hundred picked named a step with every one, in "Undo"'s and "Redo"'s tips and the draft's history).
        private string StepLabel => this.IsSeveral ? SeveralComponentsWording.Labels(this.thisSeveral) : this.thisLabel;

        // The rows each list showed, as they were shown - what an edit of them is worked out against - and the place among
        // them each row on a page stands for (a row added there has none).
        private List<ComponentPinEntry> thisCommonPins = [];
        private List<ComponentReplacementEntry> thisCommonReplacements = [];
        private List<ComponentImageEntry> thisCommonPictures = [];
        private List<ComponentLocalFileEntry> thisCommonFiles = [];
        private List<ComponentLinkEntry> thisCommonLinks = [];
        private readonly Dictionary<object, int> thisCommonPlace = new(ReferenceEqualityComparer.Instance);

        // The boxes' own hints, put back once a field is alike again.
        private readonly Dictionary<Control, string?> thisOwnHints = new(ReferenceEqualityComparer.Instance);

        // The lists' own empty lines, put back once one component is shown again.
        private readonly Dictionary<TextBlock, string?> thisOwnEmptyLines = new(ReferenceEqualityComparer.Instance);

        // An empty list of several is none THEY ALL have - not none at all.
        private void ShowEmptyLinesOfSeveral(bool several)
        {
            foreach (TextBlock line in new[] { this.NoPinsText, this.NoReplacementsText, this.NoPicturesText, this.NoFilesText, this.NoLinksText })
            {
                if (!this.thisOwnEmptyLines.ContainsKey(line))
                {
                    this.thisOwnEmptyLines[line] = line.Text;
                }

                line.Text = several ? SeveralComponentsWording.NoneInCommon : this.thisOwnEmptyLines[line];
            }
        }

        // ###########################################################################################
        // Several components of the draft held in `session`, as they read in `region` - one is shown as Show shows it.
        // ###########################################################################################
        internal void ShowSeveral(DraftSession session, IReadOnlyList<string> labels, string? region, BoardData? published, string publishedBoardFile = "")
        {
            ArgumentNullException.ThrowIfNull(labels);

            this.ShowLabels(session, labels, region, published, publishedBoardFile);
        }

        // ---- showing ------------------------------------------------------------------------------

        private void ShowSeveralComponents(DraftSession session)
        {
            BoardData board = session.Content.Board;

            // Those the draft no longer has drop out (an undo took away the change that added one); one left is shown alone.
            List<string> held = this.thisSeveral.Where(label => ComponentPaneEdits.RowShown(board, label, this.thisRegion) is not null).ToList();

            if (held.Count != this.thisSeveral.Count)
            {
                this.thisSeveral = held.Count > 1 ? held : [];
                this.thisLabel = held.FirstOrDefault() ?? this.thisLabel;

                if (!this.IsSeveral)
                {
                    this.ShowComponent();
                    return;
                }
            }

            IReadOnlyList<string> labels = this.thisSeveral;

            this.NewComponentPanel.IsVisible = false;
            this.BoardPartPanel.IsVisible = false;
            this.GonePanel.IsVisible = false;
            this.SectionsScrollViewer.IsVisible = true;

            // Nothing that acts on one component as a whole.
            this.ActionsPanel.IsVisible = false;
            this.ChangedMark.IsVisible = false;
            this.UndoChangesButton.IsVisible = false;
            this.DeleteComponentButton.IsVisible = false;
            this.ResetRowDrags();
            this.thisShowing = true;

            try
            {
                this.thisRowRegion = string.Empty;

                string? category = ComponentBulk.CommonValue(board, labels, this.thisRegion, ComponentField.Category);
                this.CategoryBox.ItemsSource = this.CategoryChoicesWith(board, category ?? string.Empty);
                this.thisDifferingValues.Clear();

                foreach (ComponentField field in Enum.GetValues<ComponentField>())
                {
                    string? common = ComponentBulk.CommonValue(board, labels, this.thisRegion, field);
                    this.thisShownValues[field] = common ?? string.Empty;
                    this.SetBoxText(field, common ?? string.Empty);
                    // Each as an edit keeps it, so values that read the same are listed once (code review, 2026-11-09).
                    this.ShowDiffers(
                        field,
                        common is null
                            ? ComponentBulk.ValuesOf(board, labels, this.thisRegion, field).Select(value => ComponentPaneEdits.Normalised(field, value)).ToList()
                            : null);
                }

                this.thisLabelShown = string.Join(", ", labels);
                this.BoardLabelBox.Text = this.thisLabelShown;
                this.BoardLabelBox.IsEnabled = false;
                this.ShowLabelProblem(null);

                this.thisCommonPlace.Clear();
                this.thisCommonPins = this.FillCommon(this.thisPins, ComponentBulk.Pins, ComponentPanePinRow.From);
                this.thisCommonReplacements = this.FillCommon(this.thisReplacements, ComponentBulk.Replacements, ComponentPaneReplacementRow.From);
                this.thisCommonLinks = this.FillCommon(this.thisLinks, ComponentBulk.Links, ComponentPaneLinkRow.From);
                this.thisCommonPictures = this.FillCommon(this.thisPictures, ComponentBulk.Pictures, ComponentPanePictureRow.From);
                this.thisCommonFiles = this.FillCommon(this.thisFiles, ComponentBulk.Files, ComponentPaneFileRow.From);
                this.thisBaselines.Clear();
                this.LoadThumbnails();
                this.ShowEmptyLinesOfSeveral(several: true);

                this.TitleText.Text = SeveralComponentsWording.Title(labels);
                BoldRuns.ShowMarked(this.RegionText, string.Empty);
                this.RegionText.IsVisible = false;

                // The per-region ticks say all of them, none, or some (ShowRegionChoices); the label is off.
                this.ShowRegionChoices(board);
                this.BoardLabelOneText.IsVisible = false;
            }
            finally
            {
                this.thisShowing = false;
            }

            this.ShowCounter();
            this.ShowCategoryProblem();
            this.ShowPinout();
            this.FillFromKiCadButton.IsVisible = false;
            this.ShowEmptyLines();
            this.ShowMakeAllLikeChoices();
            this.ShowCurrentPage();
        }

        private List<T> FillCommon<T, TRow>(ObservableCollection<TRow> rows, ComponentBulk.Section<T> section, Func<T, TRow> rowOf)
            where TRow : class
        {
            BoardData board = this.thisSession!.Content.Board;
            List<T> common = ComponentBulk.CommonRows(board, this.thisSeveral, this.thisRegion, section);

            rows.Clear();

            for (int at = 0; at < common.Count; at++)
            {
                TRow row = rowOf(common[at]);
                this.thisCommonPlace[row] = at;
                rows.Add(row);
            }

            return common;
        }

        // The box of a field of Basics or Pinout.
        private Control? FieldBox(ComponentField field) => field switch
        {
            ComponentField.Category => this.CategoryBox,
            ComponentField.Package => this.PackageBox,
            _ => this.BoxOf(field),
        };

        // Each one's value of a field they differ in, in the order of the labels shown - what its list under it offers.
        private readonly Dictionary<ComponentField, IReadOnlyList<string>> thisDifferingValues = new();

        // A field they differ in: empty, saying so - its values listed under it while it has the focus (DiffersPopup).
        // Alike: the box as it always is.
        private void ShowDiffers(ComponentField field, IReadOnlyList<string>? values)
        {
            if (this.FieldBox(field) is not { } box)
            {
                return;
            }

            if (!this.thisOwnHints.ContainsKey(box))
            {
                this.thisOwnHints[box] = ComponentPane.HintOf(box);
                box.AddHandler(InputElement.GotFocusEvent, (_, e) => this.OnDifferingFieldFocused(box, e), handledEventsToo: true);
                box.AddHandler(InputElement.LostFocusEvent, (_, _) => this.LetGoOfDifferingValuesSoon(box), handledEventsToo: true);

                if (box is ComboBox list)
                {
                    // Its own list open, the values' list makes way.
                    list.DropDownOpened += (_, _) => this.CloseDifferingValues();
                }
            }

            if (values is null)
            {
                this.thisDifferingValues.Remove(field);
                ComponentPane.SetHint(box, this.thisOwnHints[box]);

                if (this.thisDiffersField == field)
                {
                    this.CloseDifferingValues();
                }

                return;
            }

            this.thisDifferingValues[field] = values;
            ComponentPane.SetHint(box, SeveralComponentsWording.Differs(values.Distinct(StringComparer.Ordinal).Count()));
        }

        // ---- the values of a field they differ in, listed under it ----------------------------------

        // The field whose values the list shows, while it is open, and the line the arrow keys have lit (-1 none).
        private ComponentField? thisDiffersField;
        private int thisDiffersLit = -1;

        internal bool AreDifferingValuesShownForTests => this.DiffersPopup.IsOpen;

        internal IReadOnlyList<ComponentPaneDifferingValue> DifferingValuesShownForTests =>
            this.DiffersItems.ItemsSource as IReadOnlyList<ComponentPaneDifferingValue> ?? [];

        // The field of several whose box `source` is or stands in, while they differ in it - null for any other.
        private ComponentField? DifferingFieldOf(object? source)
        {
            if (!this.IsSeveral || source is not Visual visual)
            {
                return null;
            }

            foreach ((ComponentField field, _) in this.thisDifferingValues)
            {
                if (this.FieldBox(field) is { } box && (ReferenceEquals(box, visual) || box.IsVisualAncestorOf(visual)))
                {
                    return field;
                }
            }

            return null;
        }

        // ###########################################################################################
        // A FIELD THEY DIFFER IN LISTS THEIR VALUES UNDER IT WHILE IT HAS THE FOCUS, NEVER UNDER THE POINTER (owner requests,
        // 2026-11-09: "show the values in a drop-down when I hover that field - just in the same way as in "Board: Links" for
        // the "Category" field", then "I do not want that mouse hover will trigger that list of entries, as it makes it hard to
        // navigate in the other fields. Only when cursor is active in a field it should show the list ... where I can start
        // typing also and then it will filter the list"): each value once, with whose it is; what is typed narrows it
        // (ComponentPaneDifferingValue.Matching); one clicked, or lit by the arrow keys and Enter, is set on all of them, as
        // typing it would. A note with formatting is listed as its words (RichNote.PlainOf). A drop-down list's own list
        // opens on a click, so the values' list comes to it only by the keyboard.
        // ###########################################################################################
        private void OnDifferingFieldFocused(Control box, RoutedEventArgs e)
        {
            if (box is ComboBox && e is FocusChangedEventArgs { NavigationMethod: NavigationMethod.Pointer })
            {
                return;
            }

            this.OfferDifferingValues(box);
        }

        private void OfferDifferingValues(Control box)
        {
            if (this.DifferingFieldOf(box) is not { } field || !this.thisDifferingValues.TryGetValue(field, out IReadOnlyList<string>? values))
            {
                return;
            }

            this.thisDiffersField = field;
            this.DiffersPopupHeading.Text = SeveralComponentsWording.DiffersHeading;
            this.DiffersPopup.PlacementTarget = box;
            this.DiffersPopup.MinWidth = box.Bounds.Width;
            this.NarrowDifferingValues(values);
        }

        // The list as the box now reads: what is typed narrows it; nothing left to show, it is gone until something is.
        private void NarrowDifferingValues(IReadOnlyList<string> values)
        {
            if (this.thisDiffersField is not { } field)
            {
                return;
            }

            IReadOnlyList<ComponentPaneDifferingValue> all = ComponentPaneDifferingValue.Of(this.thisSeveral, values, field == ComponentField.Purpose);
            string typed = field == ComponentField.Purpose ? RichNote.PlainOf(this.BoxText(field)) : this.BoxText(field);
            IReadOnlyList<ComponentPaneDifferingValue> shown = ComponentPaneDifferingValue.Matching(all, typed);

            this.thisDiffersLit = -1;
            this.DiffersItems.ItemsSource = shown;
            this.DiffersPopup.IsOpen = shown.Count > 0;
        }

        // Whether `field` is one several differ in and its box has the keyboard - what is typed there waits (CommitHeld).
        private bool IsDifferingFieldInFocus(ComponentField field) =>
            this.IsSeveral &&
            this.thisDifferingValues.ContainsKey(field) &&
            this.FieldBox(field) is { IsKeyboardFocusWithin: true };

        // Typed into a field they differ in: the list follows. True when the text was such a field's.
        private bool FollowTypingInDifferingField(object? source)
        {
            if (this.DifferingFieldOf(source) is not { } field || !this.thisDifferingValues.TryGetValue(field, out IReadOnlyList<string>? values))
            {
                return false;
            }

            if (this.thisDiffersField == field)
            {
                this.NarrowDifferingValues(values);
            }

            return true;
        }

        // The focus gone from the box - not into the list, whose lines take no focus - the list goes.
        private void LetGoOfDifferingValuesSoon(Control box) =>
            Dispatcher.UIThread.Post(
                () =>
                {
                    if (this.DiffersPopup.IsOpen && ReferenceEquals(this.DiffersPopup.PlacementTarget, box) &&
                        !box.IsKeyboardFocusWithin && !this.DiffersPopupContent.IsPointerOver)
                    {
                        this.CloseDifferingValues();
                    }
                },
                DispatcherPriority.Background);

        // ###########################################################################################
        // Down and Up light a line, Enter picks it, Escape closes the list - as a drop-down's list - while the list is open
        // under the box the key is typed in. Every other key is the box's. Asked first by the pane's own tunnelling key
        // handler, so its Enter never takes what was typed to narrow the list into the draft. True when the key was taken.
        // ###########################################################################################
        private bool TakeDifferingListKey(KeyEventArgs e)
        {
            if (!this.DiffersPopup.IsOpen ||
                this.thisDiffersField is not { } field ||
                this.DifferingFieldOf(e.Source) != field ||
                this.DiffersItems.ItemsSource is not IReadOnlyList<ComponentPaneDifferingValue> shown ||
                shown.Count == 0)
            {
                return false;
            }

            // "Purpose or note" is a note of several lines: Down moves its caret down them until it stands at the note's end,
            // and only then reaches the list (code review, 2026-11-09: the caret could not leave the first line while the list
            // was open). Up is the list's only while a line of it is lit.
            if (field == ComponentField.Purpose && e.Key == Key.Down && this.thisDiffersLit < 0 &&
                this.PurposeBox.CaretIndex < (this.PurposeBox.Text?.Length ?? 0))
            {
                return false;
            }

            switch (e.Key)
            {
                case Key.Down:
                    this.LightDifferingValue(Math.Min(this.thisDiffersLit + 1, shown.Count - 1));
                    break;

                case Key.Up when this.thisDiffersLit >= 0:
                    this.LightDifferingValue(this.thisDiffersLit - 1);
                    break;

                case Key.Enter when this.thisDiffersLit >= 0:
                    this.PickDifferingValue(shown[this.thisDiffersLit]);
                    break;

                case Key.Escape:
                    this.CloseDifferingValues();
                    break;

                default:
                    return false;
            }

            e.Handled = true;
            return true;
        }

        private void LightDifferingValue(int index)
        {
            this.thisDiffersLit = index;

            for (int i = 0; i < this.DiffersItems.ItemCount; i++)
            {
                if (this.DiffersItems.ContainerFromIndex(i) is { } container &&
                    container.GetVisualDescendants().Prepend(container).OfType<Button>().FirstOrDefault() is { } line)
                {
                    line.Classes.Set("Lit", i == index);

                    // In sight, the list being scrolled (DiffersScroller).
                    if (i == index)
                    {
                        line.BringIntoView();
                    }
                }
            }
        }

        internal int DifferingValueLitForTests => this.thisDiffersLit;

        private void CloseDifferingValues()
        {
            this.DiffersPopup.IsOpen = false;
            this.thisDiffersField = null;
            this.thisDiffersLit = -1;
        }

        private void OnDifferingValueClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneDifferingValue picked })
            {
                this.PickDifferingValue(picked);
            }
        }

        // A value picked: on all of them, one step - an empty one too, which nothing typed could say.
        private void PickDifferingValue(ComponentPaneDifferingValue picked)
        {
            if (this.thisDiffersField is not { } field)
            {
                return;
            }

            this.CloseDifferingValues();
            this.thisShowing = true;

            try
            {
                this.SetBoxText(field, picked.Value);
            }
            finally
            {
                this.thisShowing = false;
            }

            this.CommitFieldOfSeveral(field, evenIfShown: true);
        }

        private static string? HintOf(Control box) => box switch
        {
            TextBox text => text.PlaceholderText,
            ComboBox list => list.PlaceholderText,
            AutoCompleteBox complete => complete.PlaceholderText,
            _ => null,
        };

        private static void SetHint(Control box, string? hint)
        {
            switch (box)
            {
                case TextBox text:
                    text.PlaceholderText = hint;
                    break;
                case ComboBox list:
                    list.PlaceholderText = hint;
                    break;
                case AutoCompleteBox complete:
                    complete.PlaceholderText = hint;
                    break;
            }
        }

        // Back to one component: the boxes as they always are.
        private void LetGoOfSeveral()
        {
            foreach ((Control box, string? hint) in this.thisOwnHints)
            {
                ComponentPane.SetHint(box, hint);
            }

            this.CloseDifferingValues();
            this.thisDifferingValues.Clear();
            this.thisCommonPlace.Clear();
            this.ShowEmptyLinesOfSeveral(several: false);
            this.BoardLabelBox.IsEnabled = true;
            this.SeveralBar.IsVisible = false;
        }

        // ###########################################################################################
        // The bar: who they are, what the page shown holds of them, and "Make all like" where it can be.
        // ###########################################################################################
        private void ShowSeveralBar()
        {
            this.SeveralBar.IsVisible = this.IsSeveral && this.SectionsScrollViewer.IsVisible;

            if (!this.SeveralBar.IsVisible || this.thisSession is not { } session)
            {
                return;
            }

            BoardData board = session.Content.Board;
            IReadOnlyList<string> labels = this.thisSeveral;

            BoldRuns.ShowMarked(this.SeveralText, SeveralComponentsWording.Editing(labels));

            bool packagesDiffer = this.PackagesDiffer(board);

            // What only some have on a list page, and why a page is held back - nothing on "Basics", whose fields say
            // "Differs" themselves (owner requests, 2026-11-09: the line "On "Basics" they are all the same" made way for what
            // the button's "*1" points at, then "remove the text "1 field here differs between..." - no need for that").
            this.SeveralPageText.Text = this.thisPage switch
            {
                ComponentPage.Basics => string.Empty,
                ComponentPage.Baselines => SeveralComponentsWording.BaselinesHeld,
                ComponentPage.Pinout when packagesDiffer =>
                    SeveralComponentsWording.PackagesDiffer(ComponentBulk.ValuesOf(board, labels, this.thisRegion, ComponentField.Package)),
                ComponentPage.Pinout => this.RowsLine(board, ComponentBulk.Pins, SeveralComponentsWording.Row),
                ComponentPage.Replacements => this.RowsLine(board, ComponentBulk.Replacements, SeveralComponentsWording.Row),
                ComponentPage.Pictures => this.RowsLine(board, ComponentBulk.Pictures, SeveralComponentsWording.Row),
                ComponentPage.Files => this.RowsLine(board, ComponentBulk.Files, SeveralComponentsWording.Row),
                _ => this.RowsLine(board, ComponentBulk.Links, SeveralComponentsWording.Row),
            };

            this.SeveralPageText.IsVisible = this.SeveralPageText.Text?.Length > 0;

            bool canMakeAlike = this.thisPage != ComponentPage.Baselines && !(this.thisPage == ComponentPage.Pinout && packagesDiffer);

            this.MakeAllLikePanel.IsVisible = canMakeAlike;
            this.MakeAllLikeNote.IsVisible = canMakeAlike;
            this.ShowMakeAllLikeButton();
        }

        // ###########################################################################################
        // "*1" as the bar marks it - small, bold and raised - after `text`, on the button and before the note under it.
        // ###########################################################################################
        private static Run FootnoteMark() => new(SeveralComponentsWording.FootnoteMark)
        {
            FontWeight = FontWeight.Bold,
            FontSize = 9,
            BaselineAlignment = BaselineAlignment.Superscript,
        };

        private static TextBlock WithMarkAfter(string text)
        {
            var block = new TextBlock();
            block.Inlines!.Add(new Run(text + " "));
            block.Inlines.Add(ComponentPane.FootnoteMark());
            return block;
        }

        private string RowsLine<T>(BoardData board, ComponentBulk.Section<T> section, Func<T, string> rowText) =>
            SeveralComponentsWording.RowsLine(
                this.thisPage,
                ComponentBulk.RowsOfSome(board, this.thisSeveral, this.thisRegion, section).Select(each => (each.Label, rowText(each.Row))).ToList());

        private bool PackagesDiffer(BoardData board) =>
            ComponentBulk.CommonValue(board, this.thisSeveral, this.thisRegion, ComponentField.Package) is null;

        // ###########################################################################################
        // The pages that are no page of several: Oscilloscope baseline always, and Pinout's PINS while their packages differ -
        // the bar says why. The package box stays (code review, 2026-11-09: the whole page went, the box with it, and the bar
        // sent them to "Basics", which has none): one package typed or picked there is given to all, and the pins come back.
        // Called as a page is shown, and as the package changes.
        // ###########################################################################################
        private void HoldBackPagesOfOne()
        {
            if (!this.IsSeveral || this.thisSession is not { } session)
            {
                return;
            }

            if (this.thisPage == ComponentPage.Baselines)
            {
                this.BaselinesSection.IsVisible = false;
            }

            bool pinsHeldBack = this.thisPage == ComponentPage.Pinout && this.PackagesDiffer(session.Content.Board);

            this.PinsArea.IsVisible = !pinsHeldBack;

            if (pinsHeldBack)
            {
                this.PinsRegionalCheck.IsVisible = false;
            }

            this.ShowSeveralBar();
        }

        // ---- "Make all like" -----------------------------------------------------------------------

        private void ShowMakeAllLikeChoices()
        {
            string? picked = this.MakeAllLikeSourceBox.SelectedItem as string;

            // Refilled only when the components are others.
            if (this.MakeAllLikeSourceBox.ItemsSource is not IReadOnlyList<string> listed || !listed.SequenceEqual(this.thisSeveral, StringComparer.Ordinal))
            {
                this.MakeAllLikeSourceBox.ItemsSource = this.thisSeveral.ToList();
            }
            this.MakeAllLikeSourceBox.SelectedItem = this.thisSeveral.FirstOrDefault(label => string.Equals(label, picked, StringComparison.OrdinalIgnoreCase))
                                                    ?? this.thisSeveral.FirstOrDefault();
        }

        private void OnMakeAllLikeSourceChanged(object? sender, SelectionChangedEventArgs e) => this.ShowMakeAllLikeButton();

        private void ShowMakeAllLikeButton()
        {
            string source = this.MakeAllLikeSourceBox.SelectedItem as string ?? string.Empty;

            this.MakeAllLikeButton.Content = ComponentPane.WithMarkAfter(SeveralComponentsWording.MakeAllLike(source));
            this.MakeAllLikeButton.IsEnabled = source.Length > 0;
            ToolTip.SetTip(this.MakeAllLikeButton, SeveralComponentsWording.MakeAllLikeTip(this.thisPage, source));

            this.MakeAllLikeNote.Inlines!.Clear();
            this.MakeAllLikeNote.Inlines.Add(ComponentPane.FootnoteMark());
            this.MakeAllLikeNote.Inlines.Add(new Run(" " + SeveralComponentsWording.MakeAllLikeNote(this.thisPage)));
        }

        internal string MakeAllLikeButtonTextForTests => BoldRuns.TextOf((TextBlock)this.MakeAllLikeButton.Content!);

        private void OnMakeAllLikeClick(object? sender, RoutedEventArgs e) => this.MakeAllLike(this.MakeAllLikeSourceBox.SelectedItem as string ?? string.Empty);

        // ###########################################################################################
        // The page shown made the same on all of them as `source` has it - one step - and read back.
        // ###########################################################################################
        internal void MakeAllLike(string source)
        {
            if (!this.IsSeveral || source.Length == 0)
            {
                return;
            }

            this.CommitHeld();

            IReadOnlyList<string> labels = this.thisSeveral.ToList();
            string region = this.thisRegion;
            ComponentPage page = this.thisPage;

            Func<BoardData, BoardData>? transform = page switch
            {
                ComponentPage.Basics => board => ComponentBulk.FieldsMadeLike(board, source, labels, region),

                // A list in every region, as [source] files it (ComponentBulk.MadeLike).
                ComponentPage.Replacements => board => ComponentBulk.MadeLike(board, source, labels, ComponentBulk.Replacements),
                ComponentPage.Pinout => board => ComponentBulk.MadeLike(board, source, labels, ComponentBulk.Pins),
                ComponentPage.Pictures => board => ComponentBulk.MadeLike(board, source, labels, ComponentBulk.Pictures),
                ComponentPage.Files => board => ComponentBulk.MadeLike(board, source, labels, ComponentBulk.Files),
                ComponentPage.Links => board => ComponentBulk.MadeLike(board, source, labels, ComponentBulk.Links),
                _ => null,
            };

            if (transform is null)
            {
                return;
            }

            this.Apply(DraftStepWording.ComponentsMadeLike(source, ComponentPaneWording.PageName(page)), transform);
            this.ShowComponent();
        }

        // ---- the commits of several ----------------------------------------------------------------

        // A field typed or picked: on all of them. Nothing while the box still says what was shown - empty where they differ -
        // unless a value was picked from the list under it (`evenIfShown`: the empty one is a value there).
        private void CommitFieldOfSeveral(ComponentField field, bool evenIfShown = false)
        {
            string text = this.BoxText(field);
            string wanted = ComponentPaneEdits.Normalised(field, text);

            if (!evenIfShown && this.thisShownValues.TryGetValue(field, out string? shown) && string.Equals(shown, wanted, StringComparison.Ordinal))
            {
                return;
            }

            this.thisShownValues[field] = wanted;

            IReadOnlyList<string> labels = this.thisSeveral.ToList();
            string region = this.thisRegion;

            this.Apply(
                DraftStepWording.ComponentField(this.StepLabel, ComponentPaneEdits.NameOf(field)),
                board => ComponentBulk.WithField(board, labels, region, field, text));

            // Alike now.
            this.ShowDiffers(field, null);

            // One package for all: their pins can be shown again (HoldBackPagesOfOne) - only what shows, never the boxes, as
            // the other fields of this commit are still to come.
            if (field == ComponentField.Package && this.thisPage == ComponentPage.Pinout && this.thisSession is { } now)
            {
                this.ShowRegionChoices(now.Content.Board);
                this.ShowCurrentPage();
            }

            this.ShowSeveralBar();
        }

        // ###########################################################################################
        // A list page's rows into the draft: what all of them have, changed as the page now holds it, on every one - then
        // the page's rows matched again to what they now all have, so the next change is worked out against that.
        // ###########################################################################################
        private List<T> CommitCommon<T, TRow>(
            ComponentBulk.Section<T> section,
            IEnumerable<TRow> rows,
            Func<TRow, T> entryOf,
            List<T> before,
            string description)
            where TRow : class
        {
            List<TRow> shown = rows.ToList();
            List<(T Row, int From)> after = shown.Select(row => (entryOf(row), this.thisCommonPlace.TryGetValue(row, out int at) ? at : -1)).ToList();

            IReadOnlyList<string> labels = this.thisSeveral.ToList();
            string region = this.thisRegion;

            this.Apply(description, board => ComponentBulk.WithCommonRows(board, labels, region, section, before, after));

            if (this.thisSession is not { } session)
            {
                return before;
            }

            List<T> now = ComponentBulk.CommonRows(session.Content.Board, labels, region, section);
            var taken = new bool[now.Count];

            foreach (TRow row in shown)
            {
                string key = section.Key(entryOf(row));
                int at = Enumerable.Range(0, now.Count).FirstOrDefault(index => !taken[index] && string.Equals(section.Key(now[index]), key, StringComparison.Ordinal), -1);

                if (at >= 0)
                {
                    taken[at] = true;
                    this.thisCommonPlace[row] = at;
                }
                else
                {
                    this.thisCommonPlace.Remove(row);
                }
            }

            this.ShowSeveralBar();
            return now;
        }

        private void CommitPinsOfSeveral()
        {
            // Not while their packages differ: the page shows none of their pins.
            if (this.thisSession is { } session && this.PackagesDiffer(session.Content.Board))
            {
                return;
            }

            this.thisCommonPins = this.CommitCommon(
                ComponentBulk.Pins, this.thisPins, row => row.ToEntry(), this.thisCommonPins, DraftStepWording.ComponentPins(this.StepLabel));
        }

        private void CommitReplacementsOfSeveral() =>
            this.thisCommonReplacements = this.CommitCommon(
                ComponentBulk.Replacements,
                this.thisReplacements.Where(row => !this.IsWaitingToBeTyped(row)),
                row => row.ToEntry(),
                this.thisCommonReplacements,
                DraftStepWording.ComponentReplacements(this.StepLabel));

        private void CommitLinksOfSeveral() =>
            this.thisCommonLinks = this.CommitCommon(
                ComponentBulk.Links, this.thisLinks, row => row.ToEntry(), this.thisCommonLinks, DraftStepWording.ComponentLinks(this.StepLabel));

        private void CommitPicturesOfSeveral() =>
            this.thisCommonPictures = this.CommitCommon(
                ComponentBulk.Pictures, this.thisPictures, row => row.ToEntry(), this.thisCommonPictures, DraftStepWording.ComponentPictures(this.StepLabel));

        private void CommitFilesOfSeveral() =>
            this.thisCommonFiles = this.CommitCommon(
                ComponentBulk.Files, this.thisFiles, row => row.ToEntry(), this.thisCommonFiles, DraftStepWording.ComponentFiles(this.StepLabel));
    }
}
