using System.Collections.ObjectModel;
using System.Reflection;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The "Add new component" path through the contribution editor: the window opened on a component
// that exists nowhere yet, so the contributor names it instead of the board data.
//
// Two things make this worth testing through the real window rather than only through
// ContributionPackaging.ValidateNewComponentBoardLabel:
//
//   1. What gets PRELOADED. The board-wide sections (board local files, board links) are saved as
//      whole lists (see ComponentDraftWriter.ApplyBoardScopedSection), so a new component that
//      came up with them empty would draft a removal of every one of them. The component-scoped
//      sections must be the opposite - empty, and not quietly filled from data rows that carry a
//      blank board label.
//   2. Where the typed label ENDS UP. Rows added before the label is typed are stamped with the
//      board label as it was then, which is nothing, so the save has to fill them in
//      (BuildComponentImageDraftRows and friends, via ResolveRowBoardLabel).
//
// Private members are reached by reflection, the same approach and reasoning as
// ComponentContributionValidationTests: the logic is welded to a Window.
//
// Session 2c note: see ComponentContributionValidationTests' own header - this window now saves
// into a local draft instead of posting to the contribution server, so there is no more
// BuildPayload/email/mandatory-comment to test here. What a real SAVE writes to disk is not
// exercised at this layer (see that header for why); the row-building helpers that feed it
// (BuildComponentDraftRows etc) are exercised directly instead, since they are reachable with no
// file I/O at all.
[Collection("HeadlessUi")]
public class NewComponentContributionTests
{
    [Fact]
    public void A_new_component_opens_with_one_blank_row_carrying_the_current_region()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            var componentRows = GetRows<ContributionComponentRow>(window, "thisComponentRows");

            var row = Assert.Single(componentRows);
            Assert.Equal(string.Empty, row.BoardLabel);

            // The region is the one thing the window can fill in by itself - the contribution is
            // made from a board already loaded for a region.
            Assert.Equal("PAL", row.Region);
        });
    }

    // The two halves of what a new component may bring with it. The board-wide rows have to be
    // there (see the header note), and the component-scoped ones must not - including the rows
    // whose board label is blank, which a plain comparison against an empty label would match.
    [Fact]
    public void A_new_component_preloads_the_boards_own_files_and_links_but_no_component_rows()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            Assert.Single(GetRows<ContributionComponentRow>(window, "thisComponentRows"));
            Assert.Empty(GetRows<ContributionComponentImageRow>(window, "thisComponentImageRows"));
            Assert.Empty(GetRows<ContributionComponentLocalFileRow>(window, "thisComponentLocalFileRows"));
            Assert.Empty(GetRows<ContributionComponentLinkRow>(window, "thisComponentLinkRows"));

            Assert.Equal(2, GetRows<ContributionBoardLocalFileRow>(window, "thisBoardLocalFileRows").Count);
            Assert.Single(GetRows<ContributionBoardLinkRow>(window, "thisBoardLinkRows"));
        });
    }

    // The board label is the only thing identifying a component that is not in the board data, so
    // sending without one is refused - and the refusal has to be visible on the row itself, not
    // just in the status line at the far end of the window.
    [Fact]
    public void Sending_a_new_component_with_no_board_label_marks_the_row_and_stops()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());
            window.Show();

            Submit(window);

            // ShowStatus posts its update, so the status line is empty until the queue is drained.
            PumpLayout(window);

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();
            Assert.True(row.HasBoardLabelError);
            Assert.Equal("A board label is required", row.BoardLabelErrorText);
            Assert.Contains("needs a board label", StatusText(window));

            window.Close();
        });
    }

    // Reusing a label the board already has would make the server resolve this contribution
    // against the EXISTING component and propose deleting everything the new one does not repeat.
    // The comparison ignores case, and the message names the label so the user can find it in the
    // component list.
    [Theory]
    [InlineData("U1")]
    [InlineData("u1")]
    public void A_board_label_the_board_already_has_is_refused_and_named(string typedLabel)
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());
            window.Show();

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();
            row.BoardLabel = typedLabel;

            // No category either, and it is still the label that is reported: it decides what the
            // whole contribution is about, so it is judged first.
            Submit(window);

            // ShowStatus posts its update, so the status line is empty until the queue is drained.
            PumpLayout(window);

            Assert.True(row.HasBoardLabelError);
            Assert.Equal("This board label is already taken", row.BoardLabelErrorText);
            Assert.Contains(typedLabel, StatusText(window));
            Assert.Contains("component list", StatusText(window));

            window.Close();
        });
    }

    // A label taken by a component of ANOTHER region is taken here too: the server resolves a
    // contribution by board label alone and never consults the region, so the collision is real
    // even though the component list on screen never showed that component.
    [Fact]
    public void A_board_label_used_only_by_another_regions_component_is_taken_as_well()
    {
        UiTest.Run(() =>
        {
            var boardData = new BoardData
            {
                Components =
                {
                    new ComponentEntry { BoardLabel = "U1", Region = "NTSC" }
                }
            };

            var window = new ComponentContributionWindow();
            LoadNewComponent(window, boardData);

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();
            row.BoardLabel = "U1";

            Assert.NotNull(Validate(window));
            Assert.True(row.HasBoardLabelError);
        });
    }

    // A component with no category is merged into the board data and is then unreachable - the
    // main window builds its category filter from the categories present and skips blank ones - so
    // it is refused as firmly as a missing board label, and marked on its own box.
    [Fact]
    public void Sending_a_new_component_with_no_category_marks_the_category_and_stops()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());
            window.Show();

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();
            row.BoardLabel = "U99";

            Submit(window);

            // ShowStatus posts its update, so the status line is empty until the queue is drained.
            PumpLayout(window);

            Assert.True(row.HasCategoryError);
            Assert.Equal("A category is required", row.CategoryErrorText);

            // The label is fine, so nothing may be marked on it.
            Assert.False(row.HasBoardLabelError);
            Assert.Contains("needs a category", StatusText(window));

            window.Close();
        });
    }

    // The mark has to go as soon as the box stops being empty, or it sits there contradicting what
    // the user has already typed. Both fields behave the same way, and only once BOTH are filled in
    // does validation pass - a board label on its own used to be enough, and produced exactly the
    // invisible component this check now prevents.
    [Fact]
    public void Typing_into_a_marked_box_clears_it_and_a_complete_row_validates()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();

            Assert.NotNull(Validate(window));
            Assert.True(row.HasBoardLabelError);

            row.BoardLabel = "U99";

            Assert.False(row.HasBoardLabelError);
            Assert.Equal(string.Empty, row.BoardLabelErrorText);

            // The label alone is not enough any more.
            Assert.NotNull(Validate(window));
            Assert.True(row.HasCategoryError);

            row.Category = "IC";

            Assert.False(row.HasCategoryError);
            Assert.Equal(string.Empty, row.CategoryErrorText);
            Assert.Null(Validate(window));
        });
    }

    // Typing the category by hand is what produced the split groups this suggestion list exists to
    // prevent: the main window groups and filters on the exact string, so "Capacitors" beside
    // "Capacitor" makes two groups out of one. The board's own categories are offered, de-duplicated
    // without regard to case and sorted, with blank ones left out.
    [Fact]
    public void The_category_field_suggests_the_categories_the_board_already_uses()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();

            Assert.Equal(new[] { "Capacitor", "IC" }, row.AvailableCategories);
        });
    }

    // The suggestions are offered on an existing component's row too - the same typo splits the
    // same group whether the component is new or being corrected.
    [Fact]
    public void An_existing_components_row_is_offered_the_same_categories()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            window.LoadComponent(BoardWithOneComponent(), string.Empty, "Commodore 64", "250407", "PAL", "U1", "board.xlsx");

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();

            Assert.Equal(new[] { "Capacitor", "IC" }, row.AvailableCategories);
        });
    }

    // Suggestions, not a closed list: the first component of a kind the board has never carried has
    // to be contributable, or a new category could never be introduced at all.
    [Fact]
    public void A_category_the_board_has_never_used_is_still_accepted()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();
            row.BoardLabel = "Y1";
            row.Category = "Crystal";

            Assert.DoesNotContain("Crystal", row.AvailableCategories);
            Assert.Null(Validate(window));
        });
    }

    // The visual half: each model flag has to reach its own box as a style class and repaint it,
    // or submission is blocked with nothing on screen to point at. The two are bound separately, so
    // marking one must leave the other alone - a binding pointing at the wrong flag would light up
    // the box the user has already filled in correctly.
    [Fact]
    public void The_mark_turns_the_offending_box_red_and_leaves_the_other_alone()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            window.Show();
            PumpLayout(window);

            var (boardLabelFrame, categoryFrame) = FindComponentRowFrames(window);

            // Resolved against the window's own theme variant - a plain FindResource comes back
            // unset here, and an unset expectation would make the comparison meaningless.
            Assert.True(window.TryFindResource("Text_Fail_Fg", window.ActualThemeVariant, out object? failBrush));

            var unmarkedBrush = boardLabelFrame.BorderBrush;
            Assert.NotEqual(failBrush, unmarkedBrush);

            // Nothing typed at all: the board label is what is reported.
            Validate(window);
            PumpLayout(window);

            Assert.Contains("HasError", boardLabelFrame.Classes);
            Assert.Equal(failBrush, boardLabelFrame.BorderBrush);
            Assert.DoesNotContain("HasError", categoryFrame.Classes);

            GetRows<ContributionComponentRow>(window, "thisComponentRows").Single().BoardLabel = "U99";
            Validate(window);
            PumpLayout(window);

            Assert.DoesNotContain("HasError", boardLabelFrame.Classes);
            Assert.Equal(unmarkedBrush, boardLabelFrame.BorderBrush);
            Assert.Contains("HasError", categoryFrame.Classes);
            Assert.Equal(failBrush, categoryFrame.BorderBrush);

            // The frame is what carries the mark, so its thickness must never change with it - a
            // frame that appears and disappears would shift the whole row as the user typed.
            Assert.Equal(boardLabelFrame.BorderThickness, categoryFrame.BorderThickness);

            window.Close();
        });
    }

    // Everything the draft save has to carry for a component the board has never seen: the typed
    // label reaches ResolveEffectiveBoardLabel, and the same label reaches a row that was added
    // BEFORE it was typed - those were stamped with the board label as it stood then, which was
    // nothing (BuildComponentImageDraftRows/etc resolve it against the effective label, the same
    // way BuildPayload used to via ResolveRowBoardLabel).
    [Fact]
    public void The_typed_label_reaches_new_rows_and_the_rows_that_had_none()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            // Added the way the user adds it - before the component has been given a name.
            Invoke(window, "OnAddComponentImageRowClick", null, new Avalonia.Interactivity.RoutedEventArgs());
            var imageRow = GetRows<ContributionComponentImageRow>(window, "thisComponentImageRows").Single();
            Assert.Equal(string.Empty, imageRow.BoardLabel);

            var componentRow = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();
            componentRow.BoardLabel = "  U99  ";
            componentRow.FriendlyName = "Colour clock buffer";

            string effectiveLabel = (string)Invoke(window, "ResolveEffectiveBoardLabel")!;
            Assert.Equal("U99", effectiveLabel);

            var componentDraftRows = (IReadOnlyList<ComponentDraftWriter.ComponentDraftRow>)
                Invoke(window, "BuildComponentDraftRows", effectiveLabel)!;
            Assert.Equal("U99", Assert.Single(componentDraftRows).BoardLabel);

            var componentImageDraftRows = (IReadOnlyList<ComponentDraftWriter.ComponentImageDraftRow>)
                Invoke(window, "BuildComponentImageDraftRows", effectiveLabel)!;
            Assert.Equal("U99", Assert.Single(componentImageDraftRows).BoardLabel);

            // The summary the reviewer sees.
            string displayText = (string)Invoke(window, "ResolveComponentDisplayText")!;
            Assert.Equal("U99 | Colour clock buffer", displayText);
        });
    }

    // The counterpart to the new-component path: opening an existing component must still load
    // that component's own rows. Both paths share LoadRows, and the filter that keeps a new
    // component from picking up existing data is the same one that has to let these through.
    [Fact]
    public void An_existing_component_still_loads_its_own_rows_and_shows_no_new_component_hint()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            window.LoadComponent(BoardWithOneComponent(), string.Empty, "Commodore 64", "250407", "PAL", "U1", "board.xlsx");

            Assert.Single(GetRows<ContributionComponentRow>(window, "thisComponentRows"));
            Assert.Single(GetRows<ContributionComponentImageRow>(window, "thisComponentImageRows"));
            Assert.Single(GetRows<ContributionComponentLinkRow>(window, "thisComponentLinkRows"));

            // The notice is on either way; what it says is what changes.
            Assert.Contains("modifying an existing component", HeadingText(window));
            Assert.True(window.FindControl<TextBlock>("ExistingComponentNoticeTextBlock")!.IsVisible);
            Assert.False(window.FindControl<TextBlock>("NewComponentNoticeTextBlock")!.IsVisible);

            // Nothing to validate on an existing component - its label came from the board data.
            Assert.Null(Validate(window));
        });
    }

    // The window has to say which of the two things it is doing. Opened for a new component it
    // announces it and opens the section that must be filled in.
    [Fact]
    public void A_new_component_announces_itself_and_opens_the_component_section()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            Assert.Contains("adding a component this board does not have yet", HeadingText(window));
            Assert.True(window.FindControl<TextBlock>("NewComponentNoticeTextBlock")!.IsVisible);
            Assert.False(window.FindControl<TextBlock>("ExistingComponentNoticeTextBlock")!.IsVisible);

            Assert.True(window.FindControl<Expander>("ComponentExpander")!.IsExpanded);
            Assert.Contains("new component", window.Title!, StringComparison.OrdinalIgnoreCase);
        });
    }

    // ---------------------------------------------------------------- after it has been saved

    // A draft save takes effect locally right away - unlike the old network submission, a new
    // component now appears on this board immediately. The success line is what says so, since
    // without it a save appears to have done nothing new relative to an ordinary edit.
    [Fact]
    public void The_success_line_says_a_new_component_now_appears_on_the_board()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            LoadNewComponent(window, BoardWithOneComponent());

            var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();
            row.BoardLabel = "U99";
            row.Category = "IC";

            string text = SuccessText(window);

            Assert.Contains("Saved to your local draft", text);
            Assert.Contains("U99", text);
            Assert.Contains("now appears on this board", text);
        });
    }

    // An edit of a component that is already on the board has nothing of the sort to explain - the
    // component is right there in the list either way, and only its edit takes effect.
    [Fact]
    public void An_edit_of_an_existing_component_gets_the_plain_success_line()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            window.LoadComponent(BoardWithOneComponent(), string.Empty, "Commodore 64", "250407", "PAL", "U1", "board.xlsx");

            Assert.Equal(
                "Saved to your local draft. Your changes are visible on this board right away.",
                SuccessText(window));
        });
    }

    // ---------------------------------------------------------------- the button that opens it

    // The Contribute tab lists only the components the board data already has, so the button is
    // the only way into this flow - and it must not offer it before there is a board to add to.
    [Fact]
    public void The_contribute_tab_offers_the_button_only_once_a_board_is_loaded()
    {
        UiTest.Run(() =>
        {
            var tab = new TabContribute();
            var button = tab.FindControl<Button>("AddNewComponentButton")!;

            // Before any board has been chosen, and after one is unloaded again.
            tab.LoadData(null, "PAL");
            Assert.False(button.IsEnabled);

            tab.LoadData(BoardWithOneComponent(), "PAL");
            Assert.True(button.IsEnabled);

            tab.LoadData(null, "PAL");
            Assert.False(button.IsEnabled);
        });
    }

    // ---------------------------------------------------------------- board data

    // One component with a row in every component-scoped section, a second component row whose
    // board label is BLANK (real board data does contain these), and board-wide rows that every
    // contribution has to carry back unchanged.
    private static BoardData BoardWithOneComponent()
    {
        return new BoardData
        {
            RevisionDate = "2026-01-15",
            Components =
            {
                new ComponentEntry { BoardLabel = "U1", FriendlyName = "VIC-II", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },

                // The same category in another spelling of case, and a component filed under none
                // at all - neither may reach the suggestion list as a separate entry.
                new ComponentEntry { BoardLabel = "C2", Category = "capacitor" },
                new ComponentEntry { BoardLabel = "R1", Category = "" },
                new ComponentEntry { BoardLabel = "", FriendlyName = "Stray unlabelled row" }
            },
            ComponentImages =
            {
                new ComponentImageEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "Pin 1" },
                new ComponentImageEntry { BoardLabel = "", Name = "Stray unlabelled image" }
            },
            ComponentLocalFiles =
            {
                new ComponentLocalFileEntry { BoardLabel = "", Name = "Stray unlabelled file" }
            },
            ComponentLinks =
            {
                new ComponentLinkEntry { BoardLabel = "U1", Name = "Datasheet", Url = "https://example.com/vic" }
            },
            BoardLocalFiles =
            {
                new BoardLocalFileEntry { Category = "Schematics", Name = "Sheet 1" },
                new BoardLocalFileEntry { Category = "Schematics", Name = "Sheet 2" }
            },
            BoardLinks =
            {
                new BoardLinkEntry { Category = "Repairs", Name = "Repair log", Url = "https://example.com/log" }
            }
        };
    }

    // ---------------------------------------------------------------- helpers

    // An empty data root on purpose: the folder scan then finds nothing, and no test here needs a
    // file to exist - so nothing touches the user's real data.
    private static void LoadNewComponent(ComponentContributionWindow window, BoardData boardData)
    {
        window.LoadNewComponent(boardData, string.Empty, "Commodore 64", "250407", "PAL", "board.xlsx");
    }

    private static ObservableCollection<T> GetRows<T>(ComponentContributionWindow window, string fieldName)
    {
        var field = typeof(ComponentContributionWindow).GetField(
            fieldName,
            BindingFlags.Instance | BindingFlags.NonPublic);

        return (ObservableCollection<T>)field!.GetValue(window)!;
    }

    // Returns the (Row, Message) tuple ValidateNewComponentRow produced, or null when it found
    // nothing wrong. Boxed as object because the tuple type itself is private to the window.
    private static object? Validate(ComponentContributionWindow window)
    {
        return Invoke(window, "ValidateNewComponentRow");
    }

    private static string SuccessText(ComponentContributionWindow window)
    {
        return (string)Invoke(window, "BuildSaveSuccessText")!;
    }

    private static object? Invoke(ComponentContributionWindow window, string methodName, params object?[] arguments)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            methodName,
            BindingFlags.Instance | BindingFlags.NonPublic);

        return method!.Invoke(window, arguments.Length == 0 ? null : arguments);
    }

    // Presses "Save to draft" through the real handler. Every test using this must be sure
    // validation rejects the form, or the handler would go on to try a real draft save - the
    // board label check sits before the first await, so an unusable label stops it here.
    private static void Submit(ComponentContributionWindow window)
    {
        Invoke(window, "OnSubmitClick", null, new Avalonia.Interactivity.RoutedEventArgs());
    }

    private static string HeadingText(ComponentContributionWindow window)
    {
        return window.FindControl<TextBlock>("ContributionNoticeHeadingTextBlock")!.Text ?? string.Empty;
    }

    private static string StatusText(ComponentContributionWindow window)
    {
        return window.FindControl<TextBlock>("StatusTextBlock")!.Text ?? string.Empty;
    }

    // The two marked frames, in the order the row's grid declares them: board label first,
    // category second. Nothing inside a DataTemplate can be reached by name, so position in the
    // visual tree is what there is - and which is which is checked here rather than assumed, by
    // requiring each frame to hold the control it is supposed to be marking.
    private static (Border BoardLabel, Border Category) FindComponentRowFrames(ComponentContributionWindow window)
    {
        var itemsControl = window.FindControl<ItemsControl>("ComponentRowsItemsControl")!;

        var frames = itemsControl.GetVisualDescendants()
            .OfType<Border>()
            .Where(border => border.Classes.Contains("ErrorFrame"))
            .Take(2)
            .ToList();

        Assert.Equal(2, frames.Count);

        var row = GetRows<ContributionComponentRow>(window, "thisComponentRows").Single();

        var boardLabelBox = Assert.IsType<TextBox>(frames[0].Child);
        Assert.Equal(row.BoardLabel, boardLabelBox.Text ?? string.Empty);

        var categoryBox = Assert.IsType<AutoCompleteBox>(frames[1].Child);
        Assert.Equal(row.Category, categoryBox.Text ?? string.Empty);

        // The suggestion list has to be on the control, not just on the row model.
        Assert.Equal(row.AvailableCategories, categoryBox.ItemsSource);

        return (frames[0], frames[1]);
    }

    // Headless windows do not lay out on their own, and an ItemsControl builds no containers until
    // it has been measured - so the dispatcher is drained and a layout pass forced by hand.
    private static void PumpLayout(Window window)
    {
        Dispatcher.UIThread.RunJobs();
        window.Measure(window.ClientSize);
        window.Arrange(new Avalonia.Rect(window.ClientSize));
        Dispatcher.UIThread.RunJobs();
    }
}
