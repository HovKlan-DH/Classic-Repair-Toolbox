using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A component's pages, its label, the region shown, and "Board: Regions" and "Board: Categories" -
// what the owner asked of the Contribute tab on 2026-11-01.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentPanePagesTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData Board(bool regions = true) => new()
    {
        Regions = regions ? ["PAL", "NTSC"] : [],
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510", Category = "IC", Package = "DIP-40" },
            new ComponentEntry { BoardLabel = "U2", FriendlyName = "SID", Category = "IC" },
            new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", Direction = "Ground" },
            new ComponentPinEntry { BoardLabel = "U8", Pin = "40", Name = "RES" },
        ],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Pin = "1", Name = "VSS", File = "Scope baseline/a.png", TimeDiv = "1ms", Note = "First line\nSecond line" },
            new ComponentImageEntry { BoardLabel = "U8", Region = "NTSC", Pin = "1", Name = "VSS", File = "Scope baseline/b.png", TimeDiv = "1ms" },
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/c.png" },
        ],
        ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
    };

    private DraftSession OpenDraft(BoardData? board = null)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPanePagesTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board ?? Board());

        return DraftSession.Open(this.DraftsRoot, ComponentPanePagesTests.BoardKey)!;
    }

    private static (ComponentPane Pane, Window Window) Shown(DraftSession session, string label = "U8", string region = "PAL")
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 1400 };
        window.Show();

        pane.Show(session, label, region, Board());
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    private static (ComponentPane Pane, Window Window) ShownPart(DraftSession session, PaneBoardPart part)
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 900 };
        window.Show();

        pane.ShowBoardPart(session, part, Board(), string.Empty, "PAL");
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    private static T Get<T>(ComponentPane pane, string name) where T : Control => pane.GetControl<T>(name);

    private static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    private static void Tick(CheckBox box, bool on)
    {
        box.IsChecked = on;
        box.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
    }

    // ---- replacement parts --------------------------------------------------------------------------

    // A row's boxes as drawn, typed into as a person types: the box holds the text, then it is left.
    private static void Type(ComponentPane pane, string list, int row, string text)
    {
        TextBox box = Get<ItemsControl>(pane, list).GetRealizedContainers().ElementAt(row).GetVisualDescendants().OfType<TextBox>().First();

        box.Focus();
        box.Text = text;
        Dispatcher.UIThread.RunJobs();
        Get<Button>(pane, "AddReplacementButton").Focus();
        Dispatcher.UIThread.RunJobs();
    }

    // *** OWNER REPORT, 2026-11-01: "adding another replacement part will just overwrite the first one -
    // it will not add one". On a component told apart per region, as U1 was, and on one that is not. ***
    [Theory]
    [InlineData("U8", "PAL")]
    [InlineData("U1", "NTSC")]
    [InlineData("U1", "PAL")]
    public void Another_replacement_part_is_added_beside_the_first(string label, string region)
    {
        BoardData board = Board();
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Region = "PAL", Category = "IC" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Region = "NTSC", Category = "IC" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, label, region);

            pane.ShowPage(ComponentPage.Replacements);
            Dispatcher.UIThread.RunJobs();

            Click(Get<Button>(pane, "AddReplacementButton"));
            Dispatcher.UIThread.RunJobs();
            Type(pane, "ReplacementsItemsControl", 0, "8500");

            Click(Get<Button>(pane, "AddReplacementButton"));
            Dispatcher.UIThread.RunJobs();
            Type(pane, "ReplacementsItemsControl", 1, "6510");

            Assert.Equal(
                ["8500", "6510"],
                ComponentPaneEdits.ReplacementsOf(session.Content.Board, label, region).Select(row => row.PartNumber));

            window.Close();
        });
    }

    // A part's one note holds several lines, and is a box of several lines while empty too.
    [Fact]
    public void A_replacement_parts_note_is_one_box_of_several_lines()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Replacements);
            Dispatcher.UIThread.RunJobs();
            Click(Get<Button>(pane, "AddReplacementButton"));
            Dispatcher.UIThread.RunJobs();

            Control row = Get<ItemsControl>(pane, "ReplacementsItemsControl").GetRealizedContainers().Single();
            TextBox note = row.GetVisualDescendants().OfType<TextBox>().Single(box => box.AcceptsReturn);

            Assert.True(note.TextWrapping == Avalonia.Media.TextWrapping.Wrap);
            Assert.True(note.MinHeight >= 60);
            Assert.Equal(["Part number", "Kind", "Note", "Applies to"], FieldLabels(row));

            note.Text = "Add a resistor\nSee the forum";
            pane.CommitHeld();

            Assert.Equal("Add a resistor\nSee the forum", ComponentPaneEdits.ReplacementsOf(session.Content.Board, "U8", "PAL").Single().Note);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Applies to" picked first, on a part with nothing in it yet: it is in no draft to move, so it stays
    // on the page until something is typed - it used to go, with nothing kept (2026-11-01). Then it moves
    // as any row does.
    // ###########################################################################################
    [Fact]
    public void A_new_part_with_nothing_in_it_stays_when_its_region_is_picked_first()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Replacements);
            Dispatcher.UIThread.RunJobs();
            Click(Get<Button>(pane, "AddReplacementButton"));

            ItemsControl list = Get<ItemsControl>(pane, "ReplacementsItemsControl");
            ComponentPaneReplacementRow row = list.Items.OfType<ComponentPaneReplacementRow>().Single();

            row.AppliesTo = "NTSC";
            pane.CommitHeld();

            Assert.Same(row, list.Items.OfType<ComponentPaneReplacementRow>().Single());
            Assert.Empty(session.Content.Board.ComponentReplacements);

            row.PartNumber = "8565";
            pane.CommitHeld();

            Assert.Empty(list.Items.OfType<ComponentPaneReplacementRow>());
            Assert.Equal("NTSC", session.Content.Board.ComponentReplacements.Single().Region);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A NEW ROW FOR ANOTHER REGION STAYS WHILE IT IS TYPED IN (code review, 2026-11-02). *** "Add
    // link", "NTSC" picked, a name typed and a pause: it went into the draft half typed and off the page,
    // the box from under the caret, and the address could not be typed there. So it waits until the focus
    // leaves the row - neither the typing pause nor the Tab to its next box - and then moves, once.
    // ###########################################################################################
    [Fact]
    public void A_new_link_for_another_region_stays_on_the_page_while_it_is_typed_in()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Links);
            Dispatcher.UIThread.RunJobs();
            Click(Get<Button>(pane, "AddLinkButton"));
            Dispatcher.UIThread.RunJobs();

            ItemsControl list = Get<ItemsControl>(pane, "LinksItemsControl");
            ComponentPaneLinkRow row = list.Items.OfType<ComponentPaneLinkRow>().Single();

            row.AppliesTo = "NTSC";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            List<TextBox> boxes = list.GetRealizedContainers().Single().GetVisualDescendants().OfType<TextBox>().ToList();

            boxes[0].Focus();
            boxes[0].Text = "Datasheet";
            Dispatcher.UIThread.RunJobs();

            // The typing pause.
            pane.CommitHeld(fromTyping: true);
            Dispatcher.UIThread.RunJobs();

            Assert.Same(row, list.Items.OfType<ComponentPaneLinkRow>().Single());
            Assert.Empty(session.Content.Board.ComponentLinks);

            // On to the row's next box.
            boxes[1].Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.Same(row, list.Items.OfType<ComponentPaneLinkRow>().Single());
            Assert.Empty(session.Content.Board.ComponentLinks);

            boxes[1].Text = "https://example.org/d";
            Dispatcher.UIThread.RunJobs();

            // Out of the row: into the draft once, and off this region's page.
            Get<Button>(pane, "AddLinkButton").Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(list.Items.OfType<ComponentPaneLinkRow>());
            Assert.Equal(
                ["Datasheet|https://example.org/d|NTSC"],
                session.Content.Board.ComponentLinks.Select(link => $"{link.Name}|{link.Url}|{link.Region}"));

            window.Close();
        });
    }

    // A commit nobody can wait out - CRT closing, the pane handed over - keeps what was typed.
    [Fact]
    public void A_row_for_another_region_being_typed_in_goes_into_the_draft_when_the_pane_is_handed_over()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Replacements);
            Dispatcher.UIThread.RunJobs();
            Click(Get<Button>(pane, "AddReplacementButton"));
            Dispatcher.UIThread.RunJobs();

            ItemsControl list = Get<ItemsControl>(pane, "ReplacementsItemsControl");
            ComponentPaneReplacementRow row = list.Items.OfType<ComponentPaneReplacementRow>().Single();

            row.AppliesTo = "NTSC";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            TextBox box = list.GetRealizedContainers().Single().GetVisualDescendants().OfType<TextBox>().First();
            box.Focus();
            box.Text = "8565";
            Dispatcher.UIThread.RunJobs();

            pane.CommitHeld(fromTyping: true);
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(session.Content.Board.ComponentReplacements);

            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("8565|NTSC", session.Content.Board.ComponentReplacements.Select(part => $"{part.PartNumber}|{part.Region}").Single());
            Assert.Empty(list.Items.OfType<ComponentPaneReplacementRow>());

            window.Close();
        });
    }

    // ---- what stops Submit, on its tab ----------------------------------------------------------------

    private static bool TabIsRed(ComponentPane pane, string label) => Get<TextBlock>(pane, label).Classes.Contains("error");

    // ###########################################################################################
    // *** EVERY ERROR ON A PAGE MARKS ITS TAB (code review, 2026-11-02). *** Only a pin past the case did
    // on the Baselines tab, and nothing on the Links tab, so from another page there was no sign until
    // Submit refused: a baseline with no signal name (typed in the table) and a link's "www.lemon64.com".
    // ###########################################################################################
    [Fact]
    public void A_baseline_lacking_what_it_needs_marks_its_tab_and_says_why()
    {
        BoardData board = Board();
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U2", Pin = "1", Name = "", File = "Scope baseline/x.png" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, label: "U2");

            pane.ShowPage(ComponentPage.Basics);
            Dispatcher.UIThread.RunJobs();

            Assert.True(TabIsRed(pane, "BaselinesPageLabel"));
            Assert.Contains("its signal name", pane.BaselinesForTests.Single().Problem, StringComparison.Ordinal);

            window.Close();
        });
    }

    [Fact]
    public void A_baseline_with_all_it_needs_leaves_its_tab_plain()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.False(TabIsRed(pane, "BaselinesPageLabel"));
            Assert.False(pane.BaselinesForTests.Single().HasProblem);

            window.Close();
        });
    }

    [Fact]
    public void A_link_address_Submit_would_refuse_marks_its_tab_until_it_is_put_right()
    {
        BoardData board = Board();
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Forum", Url = "www.lemon64.com" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.True(TabIsRed(pane, "LinksPageLabel"));

            pane.ShowPage(ComponentPage.Links);
            Dispatcher.UIThread.RunJobs();

            pane.GetLogicalDescendants().OfType<ItemsControl>().Single(list => list.Name == "LinksItemsControl")
                .Items.OfType<ComponentPaneLinkRow>().Single().Url = "https://www.lemon64.com";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            Assert.False(TabIsRed(pane, "LinksPageLabel"));

            window.Close();
        });
    }

    // ---- what an entry shows ------------------------------------------------------------------------

    // The field labels of a row as drawn, top to bottom.
    private static List<string> FieldLabels(Control row) =>
        row.GetVisualDescendants().OfType<TextBlock>()
            .Where(text => text.Classes.Contains("FieldLabel") && text.IsEffectivelyVisible)
            .Select(text => text.Text ?? string.Empty)
            .ToList();

    // ###########################################################################################
    // Owner's list, 2026-11-01: each entry in a dark gray frame, "Applies to" last, no "Where it comes
    // from", and a picture's note a box of several lines while empty too.
    // ###########################################################################################
    [Fact]
    public void A_baseline_and_a_picture_are_framed_and_say_what_they_apply_to_last()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Baselines);
            Dispatcher.UIThread.RunJobs();

            Control baseline = Get<ItemsControl>(pane, "BaselinesItemsControl").GetRealizedContainers().First();

            Assert.Equal(
                ["Pin", "Signal name", "Expected reading", "T/DIV", "V/DIV", "Trigger level", "Machine state", "Note", "Applies to"],
                FieldLabels(baseline));
            Assert.Contains(baseline.GetVisualDescendants().OfType<Border>(), border => border.Classes.Contains("PaneEntry"));

            pane.ShowPage(ComponentPage.Pictures);
            Dispatcher.UIThread.RunJobs();

            Control picture = Get<ItemsControl>(pane, "PicturesItemsControl").GetRealizedContainers().First();
            TextBox note = picture.GetVisualDescendants().OfType<TextBox>().Single(box => box.AcceptsReturn);

            Assert.Equal(["Name", "Note", "Applies to"], FieldLabels(picture));
            Assert.Contains(picture.GetVisualDescendants().OfType<Border>(), border => border.Classes.Contains("PaneEntry"));
            Assert.Equal(string.Empty, note.Text ?? string.Empty);
            Assert.True(note.MinHeight >= 60);

            window.Close();
        });
    }

    // ---- the pages --------------------------------------------------------------------------------

    // Owner's list, in the owner's order and words.
    [Fact]
    public void A_component_has_its_seven_pages_as_tabs_each_saying_how_many_it_holds()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.True(Get<Border>(pane, "PageTabs").IsVisible);
            Assert.Equal(
                ["Basics", "Replacement parts", "Pinout (2)", "Pictures (1)", "Oscilloscope baseline (1)", "Files", "Links"],
                Get<Border>(pane, "PageTabs").GetLogicalDescendants().OfType<TextBlock>().Select(label => label.Text));

            window.Close();
        });
    }

    [Theory]
    [InlineData("Basics", "BasicsSection")]
    [InlineData("Replacements", "ReplacementsSection")]
    [InlineData("Pinout", "PinoutSection")]
    [InlineData("Pictures", "PicturesSection")]
    [InlineData("Baselines", "BaselinesSection")]
    [InlineData("Files", "FilesSection")]
    [InlineData("Links", "LinksSection")]
    public void One_page_shows_at_a_time(string page, string shown)
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(Enum.Parse<ComponentPage>(page));

            string[] sections = ["BasicsSection", "ReplacementsSection", "PinoutSection", "PicturesSection", "BaselinesSection", "FilesSection", "LinksSection"];

            Assert.Equal([shown], sections.Where(name => Get<StackPanel>(pane, name).IsVisible));

            window.Close();
        });
    }

    // Walking a board's chips stays on one page.
    [Fact]
    public void The_page_stays_as_another_component_is_shown()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Pinout);
            pane.Show(session, "U2", "PAL", Board());
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(ComponentPage.Pinout, pane.Page);
            Assert.True(Get<StackPanel>(pane, "PinoutSection").IsVisible);

            window.Close();
        });
    }

    // Every section stays filled behind its tab: a page not on screen is still in the draft.
    [Fact]
    public void What_a_page_holds_goes_into_the_draft_from_any_page()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "PartNumberBox").Text = "906108";
            pane.ShowPage(ComponentPage.Links);

            Assert.Equal("906108", ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!.PartNumber);

            window.Close();
        });
    }

    [Fact]
    public void A_board_part_and_the_new_components_form_have_no_tabs()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);
            Assert.False(Get<Border>(pane, "PageTabs").IsVisible);

            pane.ShowNew(session, "PAL", Board());
            Assert.False(Get<Border>(pane, "PageTabs").IsVisible);

            window.Close();
        });
    }

    // The placeholder texts the owner gave, 2026-11-01 ("manufacturer" corrected, as asked).
    [Fact]
    public void The_basics_boxes_say_what_goes_in_them()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal("What will people also know this as", Get<TextBox>(pane, "FriendlyNameBox").PlaceholderText);
            Assert.Equal("Details printed on the component or its value", Get<TextBox>(pane, "TechnicalNameBox").PlaceholderText);
            Assert.Equal("Orderable part number from the manufacturer", Get<TextBox>(pane, "PartNumberBox").PlaceholderText);
            Assert.Equal("Other relevant details", Get<TextBox>(pane, "DescriptionBox").PlaceholderText);

            window.Close();
        });
    }

    // ###########################################################################################
    // THE PIN'S BOX FITS ITS COLUMN (owner report, 2026-11-01: "the left-most border of the input field
    // for 'Pin' is not visible and has been cut-off"). A text box takes at least 64 wide in this theme;
    // in a column of 60 it hung over both edges, and the layout clipped it there.
    // ###########################################################################################
    [Fact]
    public void A_pins_box_lies_wholly_inside_its_column()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Pinout);
            Dispatcher.UIThread.RunJobs();

            Control row = Get<ItemsControl>(pane, "PinsItemsControl").GetRealizedContainers().First();
            TextBox pin = row.GetVisualDescendants().OfType<TextBox>().First();

            Assert.True(pin.Bounds.Width > 0);
            Assert.True(pin.Bounds.X >= 0, $"The pin's box starts at {pin.Bounds.X}, left of its column - its left border is clipped.");

            window.Close();
        });
    }

    // Owner request, 2026-11-01: "the Note should be a textarea, having multiple lines".
    [Fact]
    public void A_baselines_note_is_a_box_of_several_lines()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Baselines);
            Dispatcher.UIThread.RunJobs();

            TextBox note = Get<ItemsControl>(pane, "BaselinesItemsControl").GetRealizedContainers().First()
                .GetVisualDescendants().OfType<TextBox>().Single(box => box.AcceptsReturn);

            Assert.Equal("First line\nSecond line", note.Text);
            Assert.True(note.TextWrapping == Avalonia.Media.TextWrapping.Wrap);
            Assert.True(note.MinHeight >= 60);

            window.Close();
        });
    }

    // ---- its label --------------------------------------------------------------------------------

    [Fact]
    public void Its_label_is_renamed_with_everything_filed_under_it()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal("U8", Get<TextBox>(pane, "BoardLabelBox").Text);

            Get<TextBox>(pane, "BoardLabelBox").Text = "U80";
            pane.CommitHeld();

            BoardData board = session.Content.Board;

            Assert.Equal("U80", pane.Label);
            Assert.Null(ComponentPaneEdits.RowShown(board, "U8", "PAL"));
            Assert.NotNull(ComponentPaneEdits.RowShown(board, "U80", "PAL"));
            Assert.Equal(2, board.ComponentPins.Count(pin => pin.BoardLabel == "U80"));
            Assert.Equal(3, board.ComponentImages.Count(image => image.BoardLabel == "U80"));
            Assert.Equal("U80", Assert.Single(board.ComponentHighlights).BoardLabel);
            Assert.Equal("Component: U80 - CPU | 6510", Get<TextBlock>(pane, "TitleText").Text);

            window.Close();
        });
    }

    // A pause mid-word would rename the component to half a label.
    [Fact]
    public void A_typing_pause_does_not_rename()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "BoardLabelBox").Text = "U";

            pane.CommitHeld(fromTyping: true);

            Assert.Equal("U8", pane.Label);
            Assert.NotNull(ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL"));

            window.Close();
        });
    }

    [Theory]
    [InlineData("U2", "already has [U2]")]
    [InlineData("u2", "already has [u2]")]
    [InlineData("   ", "needs a label")]
    public void A_label_that_cannot_be_taken_stays_in_its_box_with_the_reason(string typed, string said)
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "BoardLabelBox").Text = typed;
            pane.CommitHeld();

            Assert.Equal("U8", pane.Label);
            Assert.Equal(typed, Get<TextBox>(pane, "BoardLabelBox").Text);
            Assert.True(Get<TextBlock>(pane, "BoardLabelProblemText").IsVisible);
            Assert.Contains(said, Get<TextBlock>(pane, "BoardLabelProblemText").Text, StringComparison.Ordinal);
            Assert.False(session.HasUnsavedChanges);

            window.Close();
        });
    }

    // One step, and the pane follows the component back and forth as the steps are.
    [Fact]
    public void Undo_and_redo_of_a_rename_keep_the_component_open()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "BoardLabelBox").Text = "U80";
            pane.CommitHeld();

            session.Undo();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("U8", pane.Label);
            Assert.Equal("U8", Get<TextBox>(pane, "BoardLabelBox").Text);
            Assert.True(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);

            session.Redo();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("U80", pane.Label);
            Assert.True(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);

            window.Close();
        });
    }

    // ---- what Submit would refuse -----------------------------------------------------------------

    [Fact]
    public void A_category_emptied_is_marked_and_the_page_says_so()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.False(Get<TextBlock>(pane, "CategoryRequiredText").IsVisible);

            Get<AutoCompleteBox>(pane, "CategoryBox").Text = string.Empty;
            pane.CommitHeld();

            // Unfinished work is saved - it is the submission that needs one.
            Assert.Equal(string.Empty, ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!.Category);
            Assert.True(Get<TextBlock>(pane, "CategoryRequiredText").IsVisible);
            Assert.Contains("Missing", Get<AutoCompleteBox>(pane, "CategoryBox").Classes);
            Assert.Contains("error", Get<TextBlock>(pane, "BasicsPageLabel").Classes);

            window.Close();
        });
    }

    [Fact]
    public void A_new_component_needs_a_category()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 900 };
            window.Show();

            pane.ShowNew(session, "PAL", Board());

            Get<TextBox>(pane, "NewLabelBox").Text = "U99";
            Dispatcher.UIThread.RunJobs();

            Assert.False(Get<Button>(pane, "AddNewComponentButton").IsEnabled);
            Assert.True(Get<TextBlock>(pane, "NewCategoryNoteText").IsVisible);
            Assert.False(pane.AddNewComponent());

            Get<AutoCompleteBox>(pane, "NewCategoryBox").Text = "IC";
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<Button>(pane, "AddNewComponentButton").IsEnabled);
            Assert.True(pane.AddNewComponent());

            window.Close();
        });
    }

    // Owner request, 2026-11-01: "if I e.g. put pin 48 on a 24 pin package, this is an error".
    [Fact]
    public void A_baselines_pin_the_case_does_not_have_is_marked_on_its_box()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Baselines);
            ComponentPaneBaselineRow row = pane.BaselinesForTests[0];

            row.Pin = "48";
            pane.CommitHeld();

            Assert.True(row.PinOutside);
            Assert.Contains("DIP-40", row.PinProblem, StringComparison.Ordinal);
            Assert.Contains("error", Get<TextBlock>(pane, "BaselinesPageLabel").Classes);

            row.Pin = "40";
            pane.CommitHeld();

            Assert.False(row.PinOutside);
            Assert.DoesNotContain("error", Get<TextBlock>(pane, "BaselinesPageLabel").Classes);

            window.Close();
        });
    }

    [Fact]
    public void A_new_baselines_pin_the_case_does_not_have_holds_it_back()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Baselines);
            pane.OpenNewBaseline();

            Get<TextBox>(pane, "NewBaselinePinBox").Text = "48";
            Get<TextBox>(pane, "NewBaselineTriggerBox").Text = "1V";
            Dispatcher.UIThread.RunJobs();

            Assert.False(Get<Button>(pane, "AddNewBaselineButton").IsEnabled);

            window.Close();
        });
    }

    // Changing the case judges what is there again.
    [Fact]
    public void A_smaller_package_marks_the_pins_it_does_not_have()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Pinout);
            Assert.DoesNotContain(pane.PinsForTests, row => row.PinOutside);

            Get<AutoCompleteBox>(pane, "PackageBox").Text = "DIP-24";
            pane.CommitHeld();

            Assert.Equal(["40"], pane.PinsForTests.Where(row => row.PinOutside).Select(row => row.Pin));
            Assert.Contains("error", Get<TextBlock>(pane, "PinoutPageLabel").Classes);

            window.Close();
        });
    }

    // Owner request, 2026-11-01.
    [Fact]
    public void Unknown_is_the_last_direction_a_pin_offers()
    {
        Assert.Equal("Unknown", new ComponentPanePinRow().DirectionChoices[^1]);
    }

    // ---- the region shown --------------------------------------------------------------------------

    [Fact]
    public void The_switch_stands_under_its_tick_and_shows_only_while_it_is_ticked()
    {
        BoardData board = Board();
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Region = "PAL", Category = "IC" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Region = "NTSC", Category = "IC" });
        board.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "VCC" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            // Told apart per region, in both: each switch under its tick, saying the region shown.
            (ComponentPane pane, Window window) = Shown(session, "U1");

            Assert.Equal("This data is different per region", Get<CheckBox>(pane, "BasicsRegionalCheck").Content);
            Assert.Equal("This data is different per region", Get<CheckBox>(pane, "PinsRegionalCheck").Content);
            Assert.True(Get<StackPanel>(pane, "BasicsRegionSwitch").IsVisible);
            Assert.True(Get<StackPanel>(pane, "PinsRegionSwitch").IsVisible);
            Assert.Contains("active", Get<PaneRegionSwitch>(pane, "BasicsRegionSwitch").PalButton.Classes);
            Assert.DoesNotContain("active", Get<PaneRegionSwitch>(pane, "BasicsRegionSwitch").NtscButton.Classes);
            Assert.Contains("active", Get<PaneRegionSwitch>(pane, "PinsRegionSwitch").PalButton.Classes);

            StackPanel section = Get<StackPanel>(pane, "BasicsSection");
            Assert.Equal(section.Children.IndexOf(Get<CheckBox>(pane, "BasicsRegionalCheck")) + 1, section.Children.IndexOf(Get<StackPanel>(pane, "BasicsRegionSwitch")));

            // One picks for both - and for all of CRT.
            Click(Get<PaneRegionSwitch>(pane, "PinsRegionSwitch").NtscButton);
            Assert.Contains("active", Get<PaneRegionSwitch>(pane, "BasicsRegionSwitch").NtscButton.Classes);

            // A component that is one and the same everywhere: the ticks, and no switch.
            pane.Show(session, "U8", "PAL", board);
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<CheckBox>(pane, "BasicsRegionalCheck").IsVisible);
            Assert.False(Get<StackPanel>(pane, "BasicsRegionSwitch").IsVisible);
            Assert.False(Get<StackPanel>(pane, "PinsRegionSwitch").IsVisible);

            // Ticked, it comes.
            Tick(Get<CheckBox>(pane, "BasicsRegionalCheck"), on: true);
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<StackPanel>(pane, "BasicsRegionSwitch").IsVisible);
            Assert.False(Get<StackPanel>(pane, "PinsRegionSwitch").IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void A_board_without_regions_offers_no_switch_no_ticks_and_no_applies_to()
    {
        DraftSession session = this.OpenDraft(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "a.png" }],
        });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.False(Get<StackPanel>(pane, "BasicsRegionSwitch").IsVisible);
            Assert.False(Get<StackPanel>(pane, "PinsRegionSwitch").IsVisible);
            Assert.False(Get<CheckBox>(pane, "BasicsRegionalCheck").IsVisible);
            Assert.False(Get<CheckBox>(pane, "PinsRegionalCheck").IsVisible);
            Assert.False(pane.PicturesForTests[0].ShowsAppliesTo);

            window.Close();
        });
    }

    [Fact]
    public void Picking_a_region_shows_what_it_shows_and_asks_CRT_to_show_it()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            var picked = new List<string>();
            pane.RegionPicked += (_, region) => picked.Add(region);

            Assert.Equal("Scope baseline/a.png", Assert.Single(pane.BaselinesForTests).File);

            Click(Get<PaneRegionSwitch>(pane, "BasicsRegionSwitch").NtscButton);

            Assert.Equal(["NTSC"], picked);
            Assert.Equal("Scope baseline/b.png", Assert.Single(pane.BaselinesForTests).File);
            Assert.Contains("active", Get<PaneRegionSwitch>(pane, "BasicsRegionSwitch").NtscButton.Classes);

            // The one shown already, from either switch: nothing to do.
            Click(Get<PaneRegionSwitch>(pane, "BasicsRegionSwitch").NtscButton);
            pane.UseRegion("NTSC");
            Assert.Single(picked);

            window.Close();
        });
    }

    // A row left behind would be filed again beside itself at the next commit.
    [Fact]
    public void A_row_moved_to_the_other_region_leaves_the_page_once_and_stays_in_the_draft_once()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            ComponentPanePictureRow picture = Assert.Single(pane.PicturesForTests);
            Assert.True(picture.ShowsAppliesTo);
            Assert.Equal(ComponentPaneWording.AllRegions, picture.AppliesTo);

            picture.AppliesTo = "NTSC";
            pane.CommitHeld();

            Assert.Empty(pane.PicturesForTests);
            Assert.True(Get<TextBlock>(pane, "PaneMessageText").IsVisible);

            pane.CommitHeld();

            ComponentImageEntry moved = Assert.Single(session.Content.Board.ComponentImages, row => row.Name == "Top");
            Assert.Equal("NTSC", moved.Region);

            window.Close();
        });
    }

    [Fact]
    public void A_row_narrowed_to_the_region_shown_stays_on_the_page()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            ComponentPanePictureRow picture = Assert.Single(pane.PicturesForTests);
            picture.AppliesTo = "PAL";
            pane.CommitHeld();

            Assert.Same(picture, Assert.Single(pane.PicturesForTests));
            Assert.Equal("PAL", Assert.Single(session.Content.Board.ComponentImages, row => row.Name == "Top").Region);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A BASELINE HELD BACK STAYS WHERE THE DRAFT HAS IT (code review, 2026-11-01). *** An edit that
    // would make it no baseline leaves the draft's row as it was, in its region. Let go of for the
    // region its list now says, it read "Moved to NTSC" - and the next commit, which replaces the
    // region's rows with the page's, deleted it from the draft without a word.
    // ###########################################################################################
    [Fact]
    public void A_baseline_held_back_stays_on_its_page_and_moves_once_it_is_a_baseline_again()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            ComponentPaneBaselineRow baseline = Assert.Single(pane.BaselinesForTests);

            baseline.Pin = string.Empty;
            baseline.AppliesTo = "NTSC";
            pane.CommitHeld();

            Assert.Same(baseline, Assert.Single(pane.BaselinesForTests));
            Assert.True(baseline.HasProblem);
            Assert.False(Get<TextBlock>(pane, "PaneMessageText").IsVisible);

            // Another box left: still in the draft, as it was.
            pane.CommitHeld();

            ComponentImageEntry kept = Assert.Single(session.Content.Board.ComponentImages, row => row.File == "Scope baseline/a.png");
            Assert.Equal("PAL", kept.Region);
            Assert.Equal("1", kept.Pin);

            // Given its pin back it is a baseline again, and moves as any row does.
            baseline.Pin = "1";
            pane.CommitHeld();

            Assert.Empty(pane.BaselinesForTests);
            Assert.True(Get<TextBlock>(pane, "PaneMessageText").IsVisible);
            Assert.Equal("NTSC", Assert.Single(session.Content.Board.ComponentImages, row => row.File == "Scope baseline/a.png").Region);

            window.Close();
        });
    }

    // An empty write from a list being built is no change, as a pin's direction.
    [Fact]
    public void A_lists_empty_write_is_no_change_of_what_a_row_applies_to()
    {
        var row = new ComponentPanePictureRow { Region = "PAL" };

        row.AppliesTo = null;

        Assert.Equal("PAL", row.Region);
    }

    [Fact]
    public void The_basics_tick_tells_the_component_apart_per_region_and_back()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            CheckBox tick = Get<CheckBox>(pane, "BasicsRegionalCheck");
            Assert.False(tick.IsChecked);

            Tick(tick, true);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["PAL", "NTSC"], session.Content.Board.Components.Where(row => row.BoardLabel == "U8").Select(row => row.Region));
            Assert.Equal(DraftStepWording.ComponentRegional("U8", true), session.UndoDescription);

            Tick(tick, false);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(string.Empty, Assert.Single(session.Content.Board.Components, row => row.BoardLabel == "U8").Region);
            Assert.True(Get<TextBlock>(pane, "PaneMessageText").IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void The_pinout_tick_tells_the_pins_apart_and_a_new_pin_joins_the_set_it_stands_with()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Tick(Get<CheckBox>(pane, "PinsRegionalCheck"), true);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(4, session.Content.Board.ComponentPins.Count);
            Assert.Equal(2, pane.PinsForTests.Count);

            Click(Get<Button>(pane, "AddPinButton"));
            Assert.Equal("PAL", pane.PinsForTests[^1].Region);

            window.Close();
        });
    }

    // ###########################################################################################
    // A component with no pins - most new ones - has nothing to tell apart yet (code review, 2026-11-01):
    // the tick made an empty step and unticked itself, and a draft only looked at then showed as a
    // draft with no change in it. No step now, the tick reads what the draft says, and it says why.
    // ###########################################################################################
    [Fact]
    public void The_pinout_tick_on_a_component_with_no_pins_makes_no_step_and_says_why()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U2");

            CheckBox tick = Get<CheckBox>(pane, "PinsRegionalCheck");

            Tick(tick, true);
            Dispatcher.UIThread.RunJobs();

            Assert.False(session.CanUndo);
            Assert.False(tick.IsChecked);
            Assert.Equal(ComponentPaneWording.NoPinsToTellApart, Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // Taken off while the region shown has no pins of its own, there is nothing to make the same (code
    // review, 2026-11-01): every pin of the component went, under "the same in every region, as NTSC had
    // it". Nothing changes, the tick stays on, and it says what to do.
    // ###########################################################################################
    [Fact]
    public void The_pinout_tick_taken_off_in_a_region_with_no_pins_deletes_nothing_and_says_why()
    {
        DraftSession session = this.OpenDraft(new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U8", Region = "PAL", Pin = "1", Name = "VSS" }],
        });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U8", "NTSC");

            CheckBox tick = Get<CheckBox>(pane, "PinsRegionalCheck");
            Assert.True(tick.IsChecked);

            Tick(tick, false);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("PAL", Assert.Single(session.Content.Board.ComponentPins).Region);
            Assert.True(tick.IsChecked);
            Assert.Equal(ComponentPaneWording.NoPinsInRegion("NTSC"), Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // A rename reads back what goes by the label, not the rows (code review, 2026-11-01): a pin just added
    // and still blank, and a baseline's edit held back with its reason, went without a word.
    // ###########################################################################################
    [Fact]
    public void A_rename_keeps_a_pin_just_added_and_a_baseline_held_back()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Click(Get<Button>(pane, "AddPinButton"));
            int pins = pane.PinsForTests.Count;

            ComponentPaneBaselineRow baseline = Assert.Single(pane.BaselinesForTests);
            baseline.Pin = string.Empty;
            pane.CommitHeld();
            Assert.True(baseline.HasProblem);

            Get<TextBox>(pane, "BoardLabelBox").Text = "U80";
            pane.CommitHeld();

            Assert.Equal(pins, pane.PinsForTests.Count);
            Assert.Same(baseline, Assert.Single(pane.BaselinesForTests));
            Assert.True(baseline.HasProblem);
            Assert.Equal(string.Empty, baseline.Pin);
            Assert.Equal("U80", Get<TextBox>(pane, "BoardLabelBox").Text);
            Assert.Contains("U80", Get<TextBlock>(pane, "TitleText").Text, StringComparison.Ordinal);

            ComponentImageEntry kept = Assert.Single(session.Content.Board.ComponentImages, row => row.File == "Scope baseline/a.png");
            Assert.Equal("U80", kept.BoardLabel);
            Assert.Equal("1", kept.Pin);

            window.Close();
        });
    }

    // ---- Board: Regions ----------------------------------------------------------------------------

    [Fact]
    public void Board_regions_ticks_and_unticks_as_one_step()
    {
        DraftSession session = this.OpenDraft(new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Regions);

            CheckBox tick = Get<CheckBox>(pane, "RegionsCheckBox");

            Assert.False(tick.IsChecked);
            Assert.True(tick.IsEnabled);

            Tick(tick, true);

            Assert.Equal(["PAL", "NTSC"], session.Content.Board.Regions);
            Assert.Equal(DraftStepWording.Regions, session.UndoDescription);

            Tick(tick, false);

            Assert.Empty(session.Content.Board.Regions);

            window.Close();
        });
    }

    [Fact]
    public void Board_regions_cannot_be_unticked_while_a_row_says_a_region_and_says_why()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Regions);

            Assert.True(Get<CheckBox>(pane, "RegionsCheckBox").IsChecked);
            Assert.False(Get<CheckBox>(pane, "RegionsCheckBox").IsEnabled);
            Assert.True(Get<TextBlock>(pane, "RegionsInUseText").IsVisible);
            Assert.Contains("2 rows", Get<TextBlock>(pane, "RegionsInUseText").Text, StringComparison.Ordinal);

            window.Close();
        });
    }

    // ---- Board: Categories -------------------------------------------------------------------------

    private static ComponentPaneCategoryRow Category(ComponentPane pane, string name) =>
        pane.CategoriesForTests.Single(row => row.Name == name);

    [Fact]
    public void Board_categories_lists_each_with_how_many_components_it_has()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            Assert.Equal(["IC", "Capacitor"], pane.CategoriesForTests.Select(row => row.Name));
            Assert.Equal("2 components", Category(pane, "IC").CountText);
            Assert.Equal("1 component", Category(pane, "Capacitor").CountText);
            Assert.Equal("Board: Categories (2)", Get<TextBlock>(pane, "TitleText").Text);

            window.Close();
        });
    }

    [Fact]
    public void A_category_renamed_takes_its_components_along_as_one_step_and_keeps_its_row()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            ComponentPaneCategoryRow row = Category(pane, "IC");
            row.NameDraft = "Chip";
            pane.CommitCategoryName(row);

            Assert.Equal(["Chip", "Chip", "Capacitor"], session.Content.Board.Components.Select(component => component.Category));
            Assert.Equal(DraftStepWording.CategoryRenamed("IC", "Chip"), session.UndoDescription);

            // A name commits as its box loses focus - on the press of another row's button - and a row
            // made again would throw that click away.
            Assert.Same(row, Category(pane, "Chip"));

            window.Close();
        });
    }

    [Fact]
    public void A_category_named_as_another_joins_it()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            ComponentPaneCategoryRow row = Category(pane, "Capacitor");
            row.NameDraft = "ic";
            pane.CommitCategoryName(row);

            Assert.Equal(["IC", "IC", "IC"], session.Content.Board.Components.Select(component => component.Category));
            Assert.Equal(["IC"], pane.CategoriesForTests.Select(item => item.Name));

            window.Close();
        });
    }

    // Where a category's components went is said, so the filter beside the schematics can follow - a
    // category nothing uses has none to follow.
    [Fact]
    public void A_category_renamed_or_deleted_says_where_its_components_went()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);
            var moved = new List<(string From, string To)>();
            pane.CategoryMoved += (_, move) => moved.Add(move);

            Get<TextBox>(pane, "NewCategoryNameBox").Text = "Jumper";
            pane.AddCategory();
            ComponentPaneCategoryRow unused = Category(pane, "Jumper");
            unused.NameDraft = "Socket";
            pane.CommitCategoryName(unused);

            ComponentPaneCategoryRow chips = Category(pane, "IC");
            chips.NameDraft = "Chip";
            pane.CommitCategoryName(chips);

            ComponentPaneCategoryRow capacitors = Category(pane, "Capacitor");
            capacitors.MoveTo = "Chip";
            pane.DeleteCategory(capacitors);

            Assert.Equal([("IC", "Chip"), ("Capacitor", "Chip")], moved);

            window.Close();
        });
    }

    // Both setters are ComponentPaneRowBase.Set, which takes a null as nothing typed: a binding writing
    // one reaches no .Trim() (code review, 2026-11-01, checked: there was nothing to fix).
    [Fact]
    public void A_null_written_to_a_name_box_reads_as_nothing_typed()
    {
        var category = new ComponentPaneCategoryRow(new BoardCategoryRow("IC", 1)) { NameDraft = null! };
        var net = new ComponentPaneOtherNet("VCC") { DisplayName = null! };
        var suggestion = new ComponentPaneSignalSuggestion("Clock", "CLK") { DisplayName = null! };

        Assert.Equal(string.Empty, category.NameDraft);
        Assert.Equal(net.SuggestedName, net.NameToAdd);
        Assert.Equal("Clock", suggestion.NameToAdd);
    }

    // Two categories nothing uses yet, one named as the other: one row, as for a category with
    // components. Two rows of one name offered it twice, and deleting one took both.
    [Fact]
    public void An_unused_category_named_as_another_unused_one_leaves_one_row()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            foreach (string name in new[] { "Jumper", "Socket" })
            {
                Get<TextBox>(pane, "NewCategoryNameBox").Text = name;
                pane.AddCategory();
            }

            ComponentPaneCategoryRow row = Category(pane, "Jumper");
            row.NameDraft = "socket";
            pane.CommitCategoryName(row);

            Assert.Single(pane.CategoriesForTests, item => string.Equals(item.Name, "Socket", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(pane.CategoriesForTests, item => item.Name == "Jumper");

            window.Close();
        });
    }

    // ###########################################################################################
    // "Changed" on "Board: Categories" says its NAMES differ from the published board's (code review,
    // 2026-11-01): a component added to a category, or rows standing in another order, changes none.
    // ###########################################################################################
    [Fact]
    public void Board_categories_reads_changed_only_when_the_names_do()
    {
        BoardData published = Board();

        var added = new BoardData { Components = [.. published.Components, new ComponentEntry { BoardLabel = "U9", Category = "IC" }] };
        var moved = new BoardData { Components = [published.Components[1], published.Components[0], published.Components[2]] };
        var reordered = new BoardData { Components = [.. published.Components.AsEnumerable().Reverse()] };

        Assert.False(ComponentPane.BoardPartDiffers(PaneBoardPart.Categories, added, published));
        Assert.False(ComponentPane.BoardPartDiffers(PaneBoardPart.Categories, moved, published));

        // The categories' own order is a change since they can be put in order (2026-11-01).
        Assert.True(ComponentPane.BoardPartDiffers(PaneBoardPart.Categories, reordered, published));
        Assert.True(ComponentPane.BoardPartDiffers(PaneBoardPart.Categories, BoardCategoryEdits.Renamed(published, "IC", "Chip"), published));
        Assert.True(ComponentPane.BoardPartDiffers(PaneBoardPart.Categories, BoardCategoryEdits.Renamed(published, "IC", "ic"), published));
        Assert.True(ComponentPane.BoardPartDiffers(PaneBoardPart.Categories, published, null));
    }

    // ###########################################################################################
    // Owner request, 2026-11-01: "rearrange the category up/down, alike Files or Links" - one step that
    // "Undo" takes back, what the draft's components carry following. The components with no category
    // and one nothing uses have no rows to move, so no handle, and nothing moves past them.
    // ###########################################################################################
    [Fact]
    public void A_category_moves_up_and_down_as_one_step_and_undo_takes_it_back()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            ComponentPaneCategoryRow ic = pane.CategoriesForTests.Single(row => row.Name == "IC");
            ComponentPaneCategoryRow capacitor = pane.CategoriesForTests.Single(row => row.Name == "Capacitor");

            Assert.True(capacitor.CanMove);
            Assert.False(pane.MoveCategory(capacitor, +1));
            Assert.False(pane.MoveCategory(ic, -1));
            Assert.False(session.CanUndo);

            Assert.True(pane.MoveCategory(capacitor, -1));

            Assert.Equal(["Capacitor", "IC"], pane.CategoriesForTests.Select(row => row.Name));
            Assert.Same(capacitor, pane.CategoriesForTests[0]);
            Assert.Equal(["C1", "U8", "U2"], session.Content.Board.Components.Select(row => row.BoardLabel));
            Assert.Equal(["Capacitor", "IC"], ComponentListBuilder.BuildDistinctCategories(session.Content.Board));
            Assert.Equal(DraftStepWording.CategoriesOrdered, session.UndoDescription);
            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);

            Assert.True(session.Undo());
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["IC", "Capacitor"], pane.CategoriesForTests.Select(row => row.Name));

            window.Close();
        });
    }

    [Fact]
    public void A_category_with_no_components_and_the_components_with_none_do_not_move()
    {
        BoardData board = Board();
        board.Components.Add(new ComponentEntry { BoardLabel = "X1", Category = "" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            Get<TextBox>(pane, "NewCategoryNameBox").Text = "Unused";
            pane.AddCategory();

            Assert.Equal(["", "IC", "Capacitor", "Unused"], pane.CategoriesForTests.Select(row => row.Name));
            Assert.False(pane.CategoriesForTests[0].CanMove);
            Assert.False(pane.CategoriesForTests[3].CanMove);

            Assert.False(pane.MoveCategory(pane.CategoriesForTests[3], -1));
            Assert.False(pane.MoveCategory(pane.CategoriesForTests[1], -1));
            Assert.False(pane.MoveCategory(pane.CategoriesForTests[2], +1));

            window.Close();
        });
    }

    [Fact]
    public void A_category_given_no_name_is_refused_with_the_reason()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            ComponentPaneCategoryRow row = Category(pane, "IC");
            row.NameDraft = "  ";
            pane.CommitCategoryName(row);

            Assert.True(row.HasProblem);
            Assert.Equal("IC", row.Name);
            Assert.Equal("IC", session.Content.Board.Components[0].Category);

            window.Close();
        });
    }

    // No component may be left without one.
    [Fact]
    public void A_category_deleted_asks_where_its_components_go_and_moves_them_there()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            ComponentPaneCategoryRow row = Category(pane, "Capacitor");

            pane.AskDeleteCategory(row);
            Assert.True(row.IsDeleting);
            Assert.False(row.CanDeleteNow);
            Assert.Equal(["IC"], row.MoveChoices);
            Assert.Equal("Capacitor", session.Content.Board.Components[2].Category);

            row.MoveTo = "IC";
            pane.DeleteCategory(row);

            Assert.Equal("IC", session.Content.Board.Components[2].Category);
            Assert.Equal(["IC"], pane.CategoriesForTests.Select(item => item.Name));
            Assert.Equal(DraftStepWording.CategoryDeleted("Capacitor", "IC"), session.UndoDescription);

            window.Close();
        });
    }

    [Fact]
    public void The_only_category_has_nowhere_to_move_its_components_and_says_so()
    {
        DraftSession session = this.OpenDraft(new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            ComponentPaneCategoryRow row = Category(pane, "IC");
            pane.AskDeleteCategory(row);

            Assert.False(row.HasMoveChoices);
            Assert.False(row.CanDeleteNow);
            Assert.Equal(ComponentPaneWording.CategoryDeleteNeedsAnother, row.DeleteQuestion);

            pane.DeleteCategory(row);
            Assert.Equal("IC", session.Content.Board.Components[0].Category);

            window.Close();
        });
    }

    [Fact]
    public void The_components_with_no_category_are_a_row_that_needs_a_name_and_cannot_be_deleted()
    {
        DraftSession session = this.OpenDraft(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "" }, new ComponentEntry { BoardLabel = "U2", Category = "IC" }],
        });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            ComponentPaneCategoryRow row = pane.CategoriesForTests[0];

            Assert.True(row.NeedsOne);
            Assert.True(row.HasProblem);
            Assert.False(row.CanDelete);

            row.NameDraft = "Misc";
            pane.CommitCategoryName(row);

            Assert.Equal(["Misc", "IC"], session.Content.Board.Components.Select(component => component.Category));
            Assert.False(row.NeedsOne);

            window.Close();
        });
    }

    // Kept until first used (owner decision, 2026-11-01): no board file holds it.
    [Fact]
    public void A_category_added_is_offered_until_a_component_uses_it_and_changes_no_board()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            Get<TextBox>(pane, "NewCategoryNameBox").Text = "Jumper";
            pane.AddCategory();

            Assert.Equal("No component yet", Category(pane, "Jumper").CountText);
            Assert.Equal(string.Empty, Get<TextBox>(pane, "NewCategoryNameBox").Text);
            Assert.False(session.HasUnsavedChanges);

            // Offered in a component's box...
            pane.Show(session, "U8", "PAL", Board());
            Assert.Contains("Jumper", Assert.IsAssignableFrom<IEnumerable<string>>(Get<AutoCompleteBox>(pane, "CategoryBox").ItemsSource));

            // ...and, used, it is real data, listed once.
            Get<AutoCompleteBox>(pane, "CategoryBox").Text = "Jumper";
            pane.CommitHeld();
            pane.ShowBoardPart(session, PaneBoardPart.Categories, Board(), string.Empty, "PAL");

            Assert.Equal("1 component", Assert.Single(pane.CategoriesForTests, row => row.Name == "Jumper").CountText);

            window.Close();
        });
    }

    [Fact]
    public void A_category_added_nothing_uses_goes_at_once_when_deleted()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            Get<TextBox>(pane, "NewCategoryNameBox").Text = "Jumper";
            pane.AddCategory();

            pane.AskDeleteCategory(Category(pane, "Jumper"));

            Assert.DoesNotContain(pane.CategoriesForTests, row => row.Name == "Jumper");

            window.Close();
        });
    }

    [Theory]
    [InlineData("", "needs a name")]
    [InlineData("capacitor", "already has")]
    public void A_new_category_with_no_name_or_one_the_board_has_is_refused(string typed, string said)
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Categories);

            int before = pane.CategoriesForTests.Count;

            Get<TextBox>(pane, "NewCategoryNameBox").Text = typed;
            pane.AddCategory();

            Assert.Equal(before, pane.CategoriesForTests.Count);
            Assert.Contains(said, Get<TextBlock>(pane, "NewCategoryProblemText").Text, StringComparison.Ordinal);

            window.Close();
        });
    }

    // ---- TEMPORARY: pinouts made of scope baselines --------------------------------------------------

    [Fact]
    public void The_temporary_tool_makes_the_components_pins_as_one_step_and_the_whole_boards()
    {
        DraftSession session = this.OpenDraft(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }, new ComponentEntry { BoardLabel = "U2", Category = "IC" }],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "a.png", TimeDiv = "1ms" },
                new ComponentImageEntry { BoardLabel = "U2", Pin = "1", Name = "VDD", File = "b.png", TimeDiv = "1ms" },
            ],
        });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Pinout);
            Assert.True(Get<Border>(pane, "MigrationPanel").IsVisible);
            Assert.True(Get<Button>(pane, "MigrateComponentButton").IsEnabled);

            pane.MigrateComponent();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Ground", Assert.Single(pane.PinsForTests).Direction);
            Assert.Equal(DraftStepWording.PinsFromBaselines("U8"), session.UndoDescription);
            Assert.False(Get<Button>(pane, "MigrateComponentButton").IsEnabled);

            pane.MigrateBoard();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["U8", "U2"], session.Content.Board.ComponentPins.Select(pin => pin.BoardLabel));
            Assert.Equal(DraftStepWording.PinsFromBaselinesOfBoard(1), session.UndoDescription);

            // Nothing more to make is said, and no step.
            pane.MigrateBoard();
            Assert.Equal(ComponentPaneWording.MigrationNothing, Get<TextBlock>(pane, "PaneMessageText").Text);
            Assert.Equal(DraftStepWording.PinsFromBaselinesOfBoard(1), session.UndoDescription);

            window.Close();
        });
    }

    [Fact]
    public void The_temporary_tool_is_not_offered_on_a_board_with_no_baselines()
    {
        DraftSession session = this.OpenDraft(new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.ShowPage(ComponentPage.Pinout);
            Assert.False(Get<Border>(pane, "MigrationPanel").IsVisible);

            window.Close();
        });
    }

    // ---- what the owner asked on 2026-11-02 -----------------------------------------------------------

    // A board with regions: the switch on the five pages with no tick, always - "The checkbox is not needed".
    [Fact]
    public void The_pages_with_no_tick_show_the_region_switch_on_a_board_with_regions()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            foreach (string name in new[] { "ReplacementsRegionSwitch", "PicturesRegionSwitch", "BaselinesRegionSwitch", "FilesRegionSwitch", "LinksRegionSwitch" })
            {
                PaneRegionSwitch regionSwitch = Get<PaneRegionSwitch>(pane, name);

                Assert.True(regionSwitch.IsVisible, name);
                Assert.Contains("active", regionSwitch.PalButton.Classes);
                Assert.DoesNotContain("active", regionSwitch.NtscButton.Classes);
            }

            // Picking on one is picking for them all, and what the pages list follows.
            Click(Get<PaneRegionSwitch>(pane, "BaselinesRegionSwitch").NtscButton);

            Assert.Contains("active", Get<PaneRegionSwitch>(pane, "FilesRegionSwitch").NtscButton.Classes);
            Assert.Equal("Scope baseline/b.png", Assert.Single(pane.BaselinesForTests).File);

            window.Close();
        });
    }

    [Fact]
    public void A_board_without_regions_shows_no_switch_on_any_page()
    {
        DraftSession session = this.OpenDraft(new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            foreach (string name in new[] { "ReplacementsRegionSwitch", "PicturesRegionSwitch", "BaselinesRegionSwitch", "FilesRegionSwitch", "LinksRegionSwitch" })
            {
                Assert.False(Get<PaneRegionSwitch>(pane, name).IsVisible, name);
            }

            window.Close();
        });
    }

    // Files and links of a component: "Applies to" per row, what the region shows, moved one place.
    [Fact]
    public void A_components_files_and_links_say_what_they_apply_to_and_move_up_and_down()
    {
        BoardData board = Board();
        board.ComponentLocalFiles.AddRange(
        [
            new ComponentLocalFileEntry { BoardLabel = "U8", Name = "One", File = "a.pdf" },
            new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Two", File = "b.pdf" },
            new ComponentLocalFileEntry { BoardLabel = "U8", Name = "NTSC", File = "c.pdf", Region = "NTSC" },
        ]);
        board.ComponentLinks.AddRange(
        [
            new ComponentLinkEntry { BoardLabel = "U8", Name = "First", Url = "https://example.org/1" },
            new ComponentLinkEntry { BoardLabel = "U8", Name = "Second", Url = "https://example.org/2" },
        ]);

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal(["One", "Two"], pane.FilesForTests.Select(row => row.Name));
            Assert.All(pane.FilesForTests, row => Assert.True(row.ShowsAppliesTo));
            Assert.True(Get<TextBlock>(pane, "FilesAppliesToHeading").IsVisible);
            Assert.True(Get<TextBlock>(pane, "LinksAppliesToHeading").IsVisible);

            // Moved: the order is what the draft holds, one step.
            Assert.True(pane.MoveFile(pane.FilesForTests[1], -1));
            Assert.Equal(["Two", "One"], ComponentPaneEdits.FilesOf(session.Content.Board, "U8", "PAL").Select(row => row.Name));
            Assert.Equal(DraftStepWording.ComponentFiles("U8"), session.UndoDescription);
            Assert.False(pane.MoveFile(pane.FilesForTests[0], -1));

            ComponentPaneLinkRow second = Assert.Single(Get<ItemsControl>(pane, "LinksItemsControl").Items.OfType<ComponentPaneLinkRow>(), row => row.Name == "Second");
            Assert.True(pane.MoveLink(second, -1));
            Assert.Equal(["Second", "First"], ComponentPaneEdits.LinksOf(session.Content.Board, "U8", "PAL").Select(row => row.Name));

            // "Applies to" moves a file out of the region shown, and the page says so.
            ComponentPaneFileRow file = pane.FilesForTests[0];
            file.AppliesTo = "NTSC";
            pane.CommitHeld();

            Assert.Equal(["One"], pane.FilesForTests.Select(row => row.Name));
            Assert.Equal("NTSC", session.Content.Board.ComponentLocalFiles.Single(row => row.Name == "Two").Region);
            Assert.Equal(ComponentPaneWording.MovedToRegion("NTSC"), Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // The board's own files and links say it too, every row of the board shown whatever region CRT shows.
    [Fact]
    public void The_boards_files_and_links_have_a_region_column_while_the_board_has_regions()
    {
        BoardData board = Board();
        board.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = "a.pdf", Region = "NTSC" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);

            ComponentPaneBoardFileRow row = Assert.Single(pane.BoardFilesForTests);
            Assert.True(row.ShowsAppliesTo);
            Assert.Equal("NTSC", row.AppliesTo);
            Assert.True(Get<TextBlock>(pane, "BoardFilesAppliesToHeading").IsVisible);

            row.AppliesTo = ComponentPaneWording.AllRegions;
            pane.CommitHeld();

            Assert.Equal(string.Empty, Assert.Single(session.Content.Board.BoardLocalFiles).Region);

            window.Close();
        });
    }

    [Fact]
    public void Without_regions_a_boards_files_have_no_region_column()
    {
        BoardData board = Board(regions: false);
        board.ComponentImages.RemoveAll(row => row.Region.Length > 0);
        board.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = "a.pdf" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);

            Assert.False(Assert.Single(pane.BoardFilesForTests).ShowsAppliesTo);
            Assert.False(Get<TextBlock>(pane, "BoardFilesAppliesToHeading").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // "When adding a new pin and I give it a number that exists (and this has KiCad data), then it
    // should per default copy the Signal name" - the pinout's, else KiCad's; never over a name typed.
    // ###########################################################################################
    [Fact]
    public void A_new_baselines_signal_name_comes_with_its_pin()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            pane.UseKiCadPins(session, "U8", [new KiCadPin("7", "PA3", "Input", "/PA3")]);
            pane.OpenNewBaseline();

            TextBox pin = Get<TextBox>(pane, "NewBaselinePinBox");
            TextBox name = Get<TextBox>(pane, "NewBaselineNameBox");

            // The pinout's name, then KiCad's, and a pin neither names takes back what was offered.
            pin.Text = "40";
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("RES", name.Text);

            pin.Text = "7";
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("PA3", name.Text);

            pin.Text = "9";
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(string.Empty, name.Text);

            // A name typed stays, whatever the pin.
            name.Text = "My signal";
            pin.Text = "7";
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("My signal", name.Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // "There is no requirement that one of T/DIV or V/DIV or Trigger level must be set - they can be
    // empty, but the Pin + Signal name + image are mandatory". Saved without them it is still a baseline,
    // read back as one.
    // ###########################################################################################
    [Fact]
    public void An_oscilloscope_baseline_goes_in_with_no_settings_and_not_without_its_signal_name()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            ComponentPaneBaselineRow row = Assert.Single(pane.BaselinesForTests);

            row.TimeDiv = string.Empty;
            pane.CommitHeld();

            Assert.False(row.HasProblem);
            Assert.Single(ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL"), baseline => baseline.Pin == "1" && baseline.TimeDiv.Length == 0);

            row.Name = string.Empty;
            pane.CommitHeld();

            Assert.StartsWith("An oscilloscope baseline needs its signal name.", row.Problem, StringComparison.Ordinal);
            Assert.Equal("VSS", Assert.Single(ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL")).Name);

            window.Close();
        });
    }

    // Pictures move up and down as schematic images do: the ends greyed out, one step each.
    [Fact]
    public void A_components_pictures_move_up_and_down()
    {
        BoardData board = Board();
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U8", Name = "Bottom", File = "Component images/d.png" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal(["Top", "Bottom"], pane.PicturesForTests.Select(row => row.Name));
            Assert.False(pane.PicturesForTests[0].CanMoveUp);
            Assert.True(pane.PicturesForTests[0].CanMoveDown);
            Assert.True(pane.PicturesForTests[1].CanMoveUp);
            Assert.False(pane.PicturesForTests[1].CanMoveDown);

            Assert.True(pane.MoveComponentPicture(pane.PicturesForTests[1], -1));

            Assert.Equal(["Bottom", "Top"], pane.PicturesForTests.Select(row => row.Name));
            Assert.Equal(["Bottom", "Top"], ComponentPaneEdits.PicturesOf(session.Content.Board, "U8", "PAL").Select(row => row.Name));
            Assert.Equal(DraftStepWording.ComponentPictures("U8"), session.UndoDescription);
            Assert.False(pane.PicturesForTests[0].CanMoveUp);
            Assert.True(pane.PicturesForTests[1].CanMoveUp);

            Assert.False(pane.MoveComponentPicture(pane.PicturesForTests[0], -1));

            window.Close();
        });
    }

    // Repair notes went (owner request, 2026-11-02: "no need for it").
    [Fact]
    public void Basics_has_no_repair_notes_any_more()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.DoesNotContain("Repair notes", FieldLabels(Get<StackPanel>(pane, "BasicsSection")));

            window.Close();
        });
    }

    // ###########################################################################################
    // On a new board nothing is published to differ from, so nothing is marked (owner request,
    // 2026-11-02) - a published board's differing component says "Draft".
    // ###########################################################################################
    [Fact]
    public void A_new_boards_component_has_no_mark_and_nothing_to_undo_back_to()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            pane.Show(session, "U8", "PAL", published: null);
            Dispatcher.UIThread.RunJobs();

            Assert.False(pane.IsComponentChanged);
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.False(Get<Button>(pane, "UndoChangesButton").IsVisible);

            // One it does differ from: the mark, in the component list's word.
            pane.Show(session, "U8", "PAL", new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "Another" }] });
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.Equal("Draft", ((TextBlock)Get<Border>(pane, "ChangedMark").Child!).Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // ... BUT "ALSO ON OTHER BOARDS" STAYS (code review, 2026-11-02). The owner asked for no mark and no
    // undo on a new board; the offer rode on the same answer and went with them, though all a new board's
    // chip holds is what another board's same chip could take.
    // ###########################################################################################
    [Fact]
    public void A_new_boards_component_still_offers_also_on_other_boards()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            pane.Show(session, "U8", "PAL", published: null);
            Dispatcher.UIThread.RunJobs();

            Get<TextBox>(pane, "PartNumberBox").Text = "6510";
            pane.CommitHeld();

            Assert.True(Get<Button>(pane, "AlsoOnOtherBoardsButton").IsVisible);
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.False(Get<Button>(pane, "UndoChangesButton").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // A component's links only dragged into another order are a change of it (code review, 2026-11-02):
    // "Draft" and "Undo changes to", which puts them back - Submit and "Review changes" said so already.
    // ###########################################################################################
    [Fact]
    public void Its_links_only_moved_say_draft_and_undo_changes_puts_them_back()
    {
        BoardData Linked()
        {
            BoardData board = Board();
            board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "First", Url = "https://example.org/1" });
            board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Second", Url = "https://example.org/2" });
            return board;
        }

        DraftSession session = this.OpenDraft(Linked());

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            pane.Show(session, "U8", "PAL", Linked());
            Dispatcher.UIThread.RunJobs();

            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);

            ComponentPaneLinkRow second = Assert.Single(Get<ItemsControl>(pane, "LinksItemsControl").Items.OfType<ComponentPaneLinkRow>(), row => row.Name == "Second");
            Assert.True(pane.MoveLink(second, -1));

            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.True(Get<Button>(pane, "UndoChangesButton").IsVisible);

            _ = pane.UndoComponentChanges();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["First", "Second"], session.Content.Board.ComponentLinks.Select(row => row.Name));
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);

            window.Close();
        });
    }

    // ---- "Applies to", picked in the box itself ------------------------------------------------------

    private static ComboBox AppliesToBox(ComponentPane pane, string list, int row) =>
        Get<ItemsControl>(pane, list).ContainerFromIndex(row)!.GetVisualDescendants().OfType<ComboBox>().Single();

    // ###########################################################################################
    // *** A PICK HOLDS (code review, 2026-11-02). *** The row handed out a NEW list of choices each time
    // its region changed, and a list replaced in the middle of a pick makes the box put the item before it
    // back and write that: no row's region could be changed in the pane at all. The tests set
    // row.AppliesTo on the model, a path that keeps the value - so this one picks in the box.
    // ###########################################################################################
    [Fact]
    public void A_region_picked_in_applies_to_holds_in_the_box_and_goes_into_the_draft()
    {
        BoardData board = Board();
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.org/1" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.ShowPage(ComponentPage.Links);
            Dispatcher.UIThread.RunJobs();

            ComboBox box = AppliesToBox(pane, "LinksItemsControl", 0);
            Assert.Equal(ComponentPaneWording.AllRegions, box.SelectedItem);

            box.SelectedItem = BoardRegions.Pal;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardRegions.Pal, box.SelectedItem);

            pane.CommitHeld();
            Assert.Equal(BoardRegions.Pal, Assert.Single(session.Content.Board.ComponentLinks).Region);

            // The other region: it leaves the page, and the page says so.
            box = AppliesToBox(pane, "LinksItemsControl", 0);
            box.SelectedItem = BoardRegions.Ntsc;
            Dispatcher.UIThread.RunJobs();
            pane.CommitHeld();

            Assert.Equal(BoardRegions.Ntsc, Assert.Single(session.Content.Board.ComponentLinks).Region);
            Assert.Equal(ComponentPaneWording.MovedToRegion(BoardRegions.Ntsc), Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // A word the board holds that is no region stays listed, so the row can be put right - in one pick.
    [Fact]
    public void A_row_holding_another_word_can_be_picked_away_from_it()
    {
        BoardData board = Board();
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.org/1", Region = "PAL" });
        board.BoardLinks.Add(new BoardLinkEntry { Category = "Resources", Name = "Odd", Url = "https://example.org/2", Region = "Europe" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Links);

            ComboBox box = AppliesToBox(pane, "BoardLinksItemsControl", 0);
            Assert.Equal("Europe", box.SelectedItem);

            box.SelectedItem = BoardRegions.Ntsc;
            Dispatcher.UIThread.RunJobs();
            pane.CommitHeld();

            Assert.Equal(BoardRegions.Ntsc, box.SelectedItem);
            Assert.Equal(BoardRegions.Ntsc, Assert.Single(session.Content.Board.BoardLinks).Region);

            window.Close();
        });
    }

    // Alt+Down is what opens a drop-down list (code review, 2026-11-02: on "Applies to" it moved the row
    // and saved a reorder step). In a row's other boxes it still moves the row.
    [Fact]
    public void Alt_down_on_applies_to_opens_its_list_and_moves_no_row()
    {
        BoardData board = Board();
        board.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Resources", Name = "One", File = "a.pdf" });
        board.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Resources", Name = "Two", File = "b.pdf" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);

            ComboBox box = AppliesToBox(pane, "BoardFilesItemsControl", 0);
            box.Focus();
            Dispatcher.UIThread.RunJobs();

            window.KeyPress(Key.Down, RawInputModifiers.Alt, PhysicalKey.ArrowDown, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["One", "Two"], session.Content.Board.BoardLocalFiles.Select(row => row.Name));
            Assert.True(box.IsDropDownOpen);

            box.IsDropDownOpen = false;
            window.Close();
        });
    }

    // Given a name, the row of the components with no category stays - the same object - and has to say
    // that it can now be moved (code review, 2026-11-02: its handle stayed hidden until the part opened again).
    [Fact]
    public void The_row_of_no_category_given_a_name_gets_its_handle()
    {
        var row = new ComponentPaneCategoryRow(new BoardCategoryRow(string.Empty, 2));
        var raised = new List<string?>();
        row.PropertyChanged += (_, e) => raised.Add(e.PropertyName);

        Assert.False(row.CanMove);

        row.UseName("Misc");

        Assert.True(row.CanMove);
        Assert.Contains(nameof(ComponentPaneCategoryRow.CanMove), raised);
    }
}
