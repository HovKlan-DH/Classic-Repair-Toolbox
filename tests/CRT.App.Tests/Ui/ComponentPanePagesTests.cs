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
    // replacement part", "NTSC" picked, a part number typed and a pause: it went into the draft half typed
    // and off the page, the box from under the caret, and the note could not be typed there. So it waits
    // until the focus leaves the row - neither the typing pause nor the Tab to its next box - and then
    // moves, once. (Files and links had the same until they became the component's whole lists.)
    // ###########################################################################################
    [Fact]
    public void A_new_part_for_another_region_stays_on_the_page_while_it_is_typed_in()
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

            // The row's own boxes - a drop-down list has one inside it too.
            List<TextBox> boxes = list.GetRealizedContainers().Single().GetVisualDescendants().OfType<TextBox>()
                .Where(box => box.TemplatedParent is null)
                .ToList();

            boxes[0].Focus();
            boxes[0].Text = "8565";
            Dispatcher.UIThread.RunJobs();

            // The typing pause.
            pane.CommitHeld(fromTyping: true);
            Dispatcher.UIThread.RunJobs();

            Assert.Same(row, list.Items.OfType<ComponentPaneReplacementRow>().Single());
            Assert.Empty(session.Content.Board.ComponentReplacements);

            // On to the row's next box.
            boxes[1].Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.Same(row, list.Items.OfType<ComponentPaneReplacementRow>().Single());
            Assert.Empty(session.Content.Board.ComponentReplacements);

            boxes[1].Text = "Runs cooler";
            Dispatcher.UIThread.RunJobs();

            // Out of the row: into the draft once, and off this region's page.
            Get<Button>(pane, "AddReplacementButton").Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(list.Items.OfType<ComponentPaneReplacementRow>());
            Assert.Equal(
                ["8565|Runs cooler|NTSC"],
                session.Content.Board.ComponentReplacements.Select(part => $"{part.PartNumber}|{part.Note}|{part.Region}"));

            window.Close();
        });
    }

    // A file or a link has nothing to wait for: it never leaves its page, so it goes into the draft as any
    // row, whatever it applies to, and is typed on where it stands (owner request, 2026-11-03).
    [Fact]
    public void A_new_link_for_another_region_goes_into_the_draft_and_stays_on_the_page_while_it_is_typed_in()
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

            // Nothing typed yet: no row of the draft, and still the row of the page.
            Assert.Same(row, list.Items.OfType<ComponentPaneLinkRow>().Single());
            Assert.Empty(session.Content.Board.ComponentLinks);

            List<TextBox> boxes = list.GetRealizedContainers().Single().GetVisualDescendants().OfType<TextBox>().ToList();

            boxes[0].Focus();
            boxes[0].Text = "Datasheet";
            Dispatcher.UIThread.RunJobs();

            pane.CommitHeld(fromTyping: true);
            Dispatcher.UIThread.RunJobs();

            Assert.Same(row, list.Items.OfType<ComponentPaneLinkRow>().Single());
            Assert.Equal(["Datasheet||NTSC"], session.Content.Board.ComponentLinks.Select(link => $"{link.Name}|{link.Url}|{link.Region}"));

            boxes[1].Focus();
            boxes[1].Text = "https://example.org/d";
            Dispatcher.UIThread.RunJobs();

            Get<Button>(pane, "AddLinkButton").Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.Same(row, list.Items.OfType<ComponentPaneLinkRow>().Single());
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

    // ---- what stops Submit, on its page's entry ---------------------------------------------------------

    private static bool EntryIsRed(ComponentPane pane, ComponentPage page) => pane.PagesShown![page].HasError;

    // ###########################################################################################
    // *** EVERY ERROR ON A PAGE MARKS ITS ENTRY (code review, 2026-11-02). *** Only a pin past the case did
    // on the Baselines page, and nothing on the Links page, so from another page there was no sign until
    // Submit refused: a baseline with no signal name (typed in the table) and a link's "www.lemon64.com".
    // ###########################################################################################
    [Fact]
    public void A_baseline_lacking_what_it_needs_marks_its_entry_and_says_why()
    {
        BoardData board = Board();
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U2", Pin = "1", Name = "", File = "Scope baseline/x.png" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, label: "U2");

            pane.ShowPage(ComponentPage.Basics);
            Dispatcher.UIThread.RunJobs();

            Assert.True(EntryIsRed(pane, ComponentPage.Baselines));
            Assert.Contains("its signal name", pane.BaselinesForTests.Single().Problem, StringComparison.Ordinal);

            window.Close();
        });
    }

    [Fact]
    public void A_baseline_with_all_it_needs_leaves_its_entry_plain()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.False(EntryIsRed(pane, ComponentPage.Baselines));
            Assert.False(pane.BaselinesForTests.Single().HasProblem);

            window.Close();
        });
    }

    [Fact]
    public void A_link_address_Submit_would_refuse_marks_its_entry_until_it_is_put_right()
    {
        BoardData board = Board();
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Forum", Url = "www.lemon64.com" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.True(EntryIsRed(pane, ComponentPage.Links));

            pane.ShowPage(ComponentPage.Links);
            Dispatcher.UIThread.RunJobs();

            pane.GetLogicalDescendants().OfType<ItemsControl>().Single(list => list.Name == "LinksItemsControl")
                .Items.OfType<ComponentPaneLinkRow>().Single().Url = "https://www.lemon64.com";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            Assert.False(EntryIsRed(pane, ComponentPage.Links));

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

    // Owner's list, in the owner's order and words - each told to the left column with how many it holds,
    // Basics its rows: 1, as this component does not differ per region.
    [Fact]
    public void A_component_has_its_seven_pages_each_saying_how_many_it_holds()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal(
                ["Basics", "Replacement parts", "Pinout", "Pictures", "Oscilloscope baseline", "Files", "Links"],
                Enum.GetValues<ComponentPage>().Select(ComponentPaneWording.PageName));

            Assert.Equal(
                [1, 0, 2, 1, 1, 0, 0],
                Enum.GetValues<ComponentPage>().Select(page => pane.PagesShown![page].Count));

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

    // Every section stays filled behind the page shown: a page not on screen is still in the draft.
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
    public void A_board_part_and_the_new_components_form_have_no_pages()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);
            Assert.Null(pane.PagesShown);

            pane.ShowNew(session, "PAL", Board());
            Assert.Null(pane.PagesShown);

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
            Assert.Equal("Component: U80 | CPU | 6510", Get<TextBlock>(pane, "TitleText").Text);

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

            Get<ComboBox>(pane, "CategoryBox").SelectedItem = null;
            pane.CommitHeld();

            // Unfinished work is saved - it is the submission that needs one.
            Assert.Equal(string.Empty, ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!.Category);
            Assert.True(Get<TextBlock>(pane, "CategoryRequiredText").IsVisible);
            Assert.Contains("Missing", Get<ComboBox>(pane, "CategoryBox").Classes);
            Assert.True(EntryIsRed(pane, ComponentPage.Basics));

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

            Get<ComboBox>(pane, "NewCategoryBox").SelectedItem = "IC";
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
            Assert.True(EntryIsRed(pane, ComponentPage.Baselines));

            row.Pin = "40";
            pane.CommitHeld();

            Assert.False(row.PinOutside);
            Assert.False(EntryIsRed(pane, ComponentPage.Baselines));

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
            Assert.True(EntryIsRed(pane, ComponentPage.Pinout));

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
    public void Told_apart_per_region_each_region_keeps_its_own_fields_and_the_label_stays_one()
    {
        BoardData board = Board();
        board.Components.Add(new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", FriendlyName = "Filter" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "C1");
            CheckBox tick = Get<CheckBox>(pane, "BasicsRegionalCheck");
            TextBlock labelNote = Get<TextBlock>(pane, "BoardLabelOneText");

            Assert.False(labelNote.IsVisible);

            tick.IsChecked = true;
            tick.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.True(labelNote.IsVisible);
            Assert.Equal(ComponentPaneWording.LabelInEveryRegion, labelNote.Text);

            // Typed in PAL, typed in NTSC, each box as the switch goes back and forth.
            void Type(string region)
            {
                Get<TextBox>(pane, "FriendlyNameBox").Text = $"Friendly {region}";
                Get<TextBox>(pane, "TechnicalNameBox").Text = $"Technical {region}";
                Get<TextBox>(pane, "PartNumberBox").Text = $"Part {region}";
                Get<TextBox>(pane, "DescriptionBox").Text = $"Description {region}";
            }

            void AssertShows(string region)
            {
                Assert.Equal($"Friendly {region}", Get<TextBox>(pane, "FriendlyNameBox").Text);
                Assert.Equal($"Technical {region}", Get<TextBox>(pane, "TechnicalNameBox").Text);
                Assert.Equal($"Part {region}", Get<TextBox>(pane, "PartNumberBox").Text);
                Assert.Equal($"Description {region}", Get<TextBox>(pane, "DescriptionBox").Text);
            }

            Type("PAL");
            pane.UseRegion("NTSC");
            Type("NTSC");
            pane.UseRegion("PAL");
            AssertShows("PAL");
            pane.UseRegion("NTSC");
            AssertShows("NTSC");

            // The label is every region's: what is printed on the board, and what the rows are filed under.
            Get<TextBox>(pane, "BoardLabelBox").Text = "C1 NTSC";
            pane.CommitHeld();
            pane.UseRegion("PAL");

            Assert.Equal("C1 NTSC", Get<TextBox>(pane, "BoardLabelBox").Text);
            AssertShows("PAL");
            Assert.Equal(["C1 NTSC", "C1 NTSC"], session.Content.Board.Components.Where(row => row.FriendlyName.StartsWith("Friendly", StringComparison.Ordinal)).Select(row => row.BoardLabel));

            window.Close();
        });
    }

    // ###########################################################################################
    // BASICS' ORDER (owner request, 2026-11-04): Board label, Friendly name, Technical name or value and
    // Short description down the left, Category and Part number on the right, every box one column wide -
    // and the warning right under its box, nothing between them.
    // ###########################################################################################
    [Fact]
    public void Basics_stand_in_two_columns_the_short_description_as_wide_as_the_board_label()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U8");

            window.Measure(new Size(1100, 1000));
            window.Arrange(new Rect(0, 0, 1100, 1000));
            Dispatcher.UIThread.RunJobs();

            Point At(string name) => Get<Control>(pane, name).TranslatePoint(default, pane)!.Value;

            string[] left = ["BoardLabelBox", "FriendlyNameBox", "TechnicalNameBox", "DescriptionBox"];
            string[] right = ["CategoryBox", "PartNumberBox"];

            foreach (string name in left)
            {
                Assert.Equal(At("BoardLabelBox").X, At(name).X, 1);
                Assert.Equal(Get<Control>(pane, "BoardLabelBox").Bounds.Width, Get<Control>(pane, name).Bounds.Width, 1);
            }

            for (int i = 1; i < left.Length; i++)
            {
                Assert.True(At(left[i]).Y > At(left[i - 1]).Y, $"{left[i]} should stand under {left[i - 1]}");
            }

            Assert.All(right, name => Assert.True(At(name).X > At("BoardLabelBox").X + Get<Control>(pane, "BoardLabelBox").Bounds.Width, $"{name} should stand in the right column"));
            Assert.Equal(At("BoardLabelBox").Y, At("CategoryBox").Y, 1);
            Assert.Equal(At("FriendlyNameBox").Y, At("PartNumberBox").Y, 1);

            // Too long: said right under the box, the counter beside it on the same line.
            Get<TextBox>(pane, "DescriptionBox").Text = new string('x', ComponentVocabulary.ShortDescriptionLimit + 1);
            Dispatcher.UIThread.RunJobs();
            window.Measure(new Size(1100, 1000));
            window.Arrange(new Rect(0, 0, 1100, 1000));
            Dispatcher.UIThread.RunJobs();

            TextBlock warning = Get<TextBlock>(pane, "DescriptionTooLongText");
            Assert.True(warning.IsVisible);

            double boxBottom = At("DescriptionBox").Y + Get<Control>(pane, "DescriptionBox").Bounds.Height;
            Assert.InRange(warning.TranslatePoint(default, pane)!.Value.Y - boxBottom, 0, 4);
            Assert.Equal(At("DescriptionTooLongText").Y, At("DescriptionCounter").Y, 1);

            window.Close();
        });
    }

    [Fact]
    public void Which_region_a_page_shows_is_said_only_where_the_page_depends_on_it()
    {
        BoardData board = Board();
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Region = "PAL", Category = "IC" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Region = "NTSC", Category = "IC" });
        board.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "VCC" });

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            // Told apart per region, in both: the line says the region shown, on either page.
            (ComponentPane pane, Window window) = Shown(session, "U1");
            TextBlock line = Get<TextBlock>(pane, "RegionShownText");

            Assert.Equal("This data is different per region", Get<CheckBox>(pane, "BasicsRegionalCheck").Content);
            Assert.Equal("This data is different per region", Get<CheckBox>(pane, "PinsRegionalCheck").Content);
            Assert.True(line.IsVisible);
            Assert.Equal(ComponentPaneWording.RegionShownLine("PAL"), line.Text);

            pane.ShowPage(ComponentPage.Pinout);
            Assert.True(line.IsVisible);

            // The buttons switching it are CRT's own, and the line follows them.
            pane.UseRegion("NTSC");
            Assert.Equal(ComponentPaneWording.RegionShownLine("NTSC"), line.Text);

            // A component that is one and the same everywhere: the ticks, and nothing to say on those pages.
            pane.Show(session, "U8", "PAL", board);
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<CheckBox>(pane, "BasicsRegionalCheck").IsVisible);
            pane.ShowPage(ComponentPage.Basics);
            Assert.False(line.IsVisible);
            pane.ShowPage(ComponentPage.Pinout);
            Assert.False(line.IsVisible);

            // Ticked, it comes - on that page.
            pane.ShowPage(ComponentPage.Basics);
            Tick(Get<CheckBox>(pane, "BasicsRegionalCheck"), on: true);
            Dispatcher.UIThread.RunJobs();

            Assert.True(line.IsVisible);
            pane.ShowPage(ComponentPage.Pinout);
            Assert.False(line.IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void A_board_without_regions_says_no_region_and_offers_no_ticks_and_no_applies_to()
    {
        DraftSession session = this.OpenDraft(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "a.png" }],
        });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            foreach (ComponentPage page in Enum.GetValues<ComponentPage>())
            {
                pane.ShowPage(page);
                Assert.False(Get<TextBlock>(pane, "RegionShownText").IsVisible, page.ToString());
            }

            Assert.False(Get<CheckBox>(pane, "BasicsRegionalCheck").IsVisible);
            Assert.False(Get<CheckBox>(pane, "PinsRegionalCheck").IsVisible);
            Assert.False(pane.PicturesForTests[0].ShowsAppliesTo);

            window.Close();
        });
    }

    [Fact]
    public void The_region_CRT_shows_is_what_a_page_lists()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal("Scope baseline/a.png", Assert.Single(pane.BaselinesForTests).File);

            // CRT's buttons under its component list (Main.PickRegion -> TabDrafts.UseRegion).
            pane.UseRegion("NTSC");

            Assert.Equal("Scope baseline/b.png", Assert.Single(pane.BaselinesForTests).File);

            // The one shown already: nothing to do.
            ComponentPaneBaselineRow shown = Assert.Single(pane.BaselinesForTests);
            pane.UseRegion("NTSC");
            Assert.Same(shown, Assert.Single(pane.BaselinesForTests));

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
            Assert.Contains("Jumper", Assert.IsAssignableFrom<IEnumerable<string>>(Get<ComboBox>(pane, "CategoryBox").ItemsSource));

            // ...and, used, it is real data, listed once.
            Get<ComboBox>(pane, "CategoryBox").SelectedItem = "Jumper";
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

    // ---- the wheel over a drop-down list --------------------------------------------------------------

    // ###########################################################################################
    // A FOCUSED, CLOSED LIST NEVER STEPS UNDER THE WHEEL HERE (code review, 2026-11-04). Scrolling the page
    // with the pointer over the category just picked gave the component another one, a step of the draft
    // nobody meant; over the clone panel's board, each notch read another board's file. The page scrolls on,
    // as it does over any other box.
    // ###########################################################################################
    [Fact]
    public void The_wheel_over_a_closed_list_scrolls_the_page_and_changes_nothing()
    {
        DraftSession session = this.OpenDraft(new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
                new ComponentEntry { BoardLabel = "R1", Category = "Resistor" },
            ],
        });

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            window.Height = 300;
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            ComboBox box = Get<ComboBox>(pane, "CategoryBox");
            ScrollViewer page = Get<ScrollViewer>(pane, "SectionsScrollViewer");

            Assert.True(page.Extent.Height >= page.Viewport.Height + 50, $"The page should scroll a whole step: {page.Extent.Height} in {page.Viewport.Height}");

            object? picked = box.SelectedItem;
            string? step = session.UndoDescription;

            Assert.NotNull(picked);

            box.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(box.IsFocused);

            Point over = box.TranslatePoint(new Point(10, 10), window)!.Value;

            window.MouseWheel(over, new Vector(0, -1));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(picked, box.SelectedItem);
            Assert.Equal(step, session.UndoDescription);
            Assert.Equal(50, page.Offset.Y, 1);

            // Past the top of the page there is nowhere to go.
            window.MouseWheel(over, new Vector(0, 3));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(picked, box.SelectedItem);
            Assert.Equal(0, page.Offset.Y, 1);

            window.Close();
        });
    }

    // ---- TEMPORARY: pinouts made of scope baselines --------------------------------------------------

    // Every wait the user watches runs under BusyOverlay: the board is worked out off the UI thread, as a
    // large KiCad board froze the window. On this thread, the call could not come back before the gate.
    [Fact]
    public async Task The_temporary_tool_works_the_whole_board_out_off_the_UI_thread_under_the_overlay()
    {
        DraftSession session = this.OpenDraft(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
        });

        await UiTest.RunAsync(async () =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var overlay = new BusyOverlay();
            var window = new Window { Content = new Grid { Children = { pane, overlay } }, Width = 1100, Height = 900 };

            window.Show();
            pane.Show(session, "U8", "PAL", Board());
            Dispatcher.UIThread.RunJobs();

            using var gate = new ManualResetEventSlim();

            pane.ReadKiCadForPinouts = () => Task.FromResult<KiCadPinoutSource?>(new KiCadPinoutSource(
                [new KiCadBoardComponent("U8", string.Empty, "IC", string.Empty, string.Empty)],
                _ =>
                {
                    gate.Wait(TimeSpan.FromSeconds(2));
                    return [new KiCadPin("1", "VSS", "Power input", "GND")];
                }));

            Task running = pane.MigrateBoardAsync();

            Assert.False(running.IsCompleted);
            Assert.True(overlay.IsBusy);

            gate.Set();
            await running;

            Assert.False(overlay.IsBusy);
            Assert.Equal("U8", Assert.Single(session.Content.Board.ComponentPins).BoardLabel);
            Assert.Equal(DraftStepWording.PinsFromBaselinesOfBoard(1), session.UndoDescription);

            window.Close();
        });
    }

    [Fact]
    public async Task The_temporary_tool_makes_the_components_pins_as_one_step_and_the_whole_boards()
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

        await UiTest.RunAsync(async () =>
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

            await pane.MigrateBoardAsync();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["U8", "U2"], session.Content.Board.ComponentPins.Select(pin => pin.BoardLabel));
            Assert.Equal(DraftStepWording.PinsFromBaselinesOfBoard(1), session.UndoDescription);

            // Nothing more to make is said, and no step.
            await pane.MigrateBoardAsync();
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

    // A board with regions: the pages listing what the region shown shows say which, always. Files and
    // Links list every row, so nothing there hangs on it (owner request, 2026-11-03).
    [Fact]
    public void The_pages_with_no_tick_say_the_region_they_show_and_files_and_links_do_not()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            TextBlock line = Get<TextBlock>(pane, "RegionShownText");

            foreach (ComponentPage page in new[] { ComponentPage.Replacements, ComponentPage.Pictures, ComponentPage.Baselines })
            {
                pane.ShowPage(page);
                Assert.True(line.IsVisible, page.ToString());
                Assert.Equal(ComponentPaneWording.RegionShownLine("PAL"), line.Text);
            }

            foreach (ComponentPage page in new[] { ComponentPage.Files, ComponentPage.Links })
            {
                pane.ShowPage(page);
                Assert.False(line.IsVisible, page.ToString());
            }

            window.Close();
        });
    }

    // ###########################################################################################
    // THE CATEGORY IS PICKED, NEVER TYPED (owner request, 2026-11-03: "everywhere we have the 'Category',
    // then it should be a drop-down ... Only in the 'Board: Category' we should allow for an input field"),
    // and a pick is a step of the draft like any other. A row's own spelling stays listed, or showing the
    // component would change it.
    // ###########################################################################################
    [Fact]
    public void A_components_category_is_a_drop_down_of_the_boards_and_a_pick_goes_into_the_draft()
    {
        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
                new ComponentEntry { BoardLabel = "U9", Category = "ic" },
            ],
        };

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            ComboBox box = Get<ComboBox>(pane, "CategoryBox");

            Assert.Equal("IC", box.SelectedItem);
            Assert.Contains("Capacitor", box.Items.OfType<string>());

            box.SelectedItem = "Capacitor";
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Capacitor", session.Content.Board.Components.Single(row => row.BoardLabel == "U8").Category);

            // Another row's spelling is listed for it as it is spelt, and showing it changes nothing.
            int steps = session.UndoDescription?.Length ?? 0;
            string? undo = session.UndoDescription;

            pane.Show(session, "U9", "PAL", board);
            Dispatcher.UIThread.RunJobs();
            pane.CommitHeld();

            Assert.Equal("ic", box.SelectedItem);
            Assert.Equal("ic", session.Content.Board.Components.Single(row => row.BoardLabel == "U9").Category);
            Assert.Equal(undo, session.UndoDescription);

            window.Close();
        });
    }

    // With nothing to pick, the form says where a category is made - and adds nothing until there is one.
    [Fact]
    public void The_new_component_form_says_where_a_category_is_made_when_the_board_has_none()
    {
        DraftSession session = this.OpenDraft(new BoardData());

        UiTest.Run(() =>
        {
            var pane = new ComponentPane();
            var window = new Window { Content = pane, Width = 1000, Height = 800 };
            window.Show();

            pane.ShowNew(session, "PAL", null);
            Dispatcher.UIThread.RunJobs();

            Get<TextBox>(pane, "NewLabelBox").Text = "U1";
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(Get<ComboBox>(pane, "NewCategoryBox").Items);
            Assert.Equal(ComponentPaneWording.NoCategoriesToPick, Get<TextBlock>(pane, "NewCategoryNoteText").Text);
            Assert.False(Get<Button>(pane, "AddNewComponentButton").IsEnabled);

            window.Close();
        });
    }

    // ###########################################################################################
    // THE PACKAGE'S LIST OPENS WITH THE FOCUS, every case in it even with one already chosen (owner
    // request, 2026-11-03: "when the 'Package' has cursor focus, then I would like for it to show that list
    // of options"), and narrows once typed in.
    // ###########################################################################################
    [Fact]
    public void The_packages_list_opens_with_the_focus_showing_every_case_and_narrows_as_typed()
    {
        BoardData board = new()
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC", Package = "DIP-40" }],
        };

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.ShowPage(ComponentPage.Pinout);
            Dispatcher.UIThread.RunJobs();

            AutoCompleteBox box = Get<AutoCompleteBox>(pane, "PackageBox");
            Assert.False(box.IsDropDownOpen);

            box.GetVisualDescendants().OfType<TextBox>().First().Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.True(box.IsDropDownOpen);
            Assert.Equal(ComponentVocabulary.Packages.Count, ListedIn(box));

            box.Text = "TO-";
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(2, ListedIn(box));

            window.Close();
        });

        // What its list shows - the popup's list is laid out only once the box is, so read through its template.
        static int ListedIn(AutoCompleteBox box) =>
            box.GetVisualDescendants().OfType<Avalonia.Controls.Primitives.Popup>().Single().Child?.GetVisualDescendants().OfType<ListBox>().FirstOrDefault()?.ItemCount
                ?? box.GetLogicalDescendants().OfType<ListBox>().First().ItemCount;
    }

    // ###########################################################################################
    // A package picked with the mouse leaves the list shut, though the list opens on every GotFocus (code
    // review, 2026-11-04, feared otherwise; measured on Avalonia 12.1.4): the press takes the focus to the
    // list's row, and AutoCompleteBox closes its list on any LostFocus, so nothing hands the focus back to
    // the box's text. An Avalonia update changing that would open the list over the page again.
    // ###########################################################################################
    [Fact]
    public void A_package_picked_with_the_mouse_leaves_the_list_shut_until_the_box_is_entered_again()
    {
        BoardData board = new()
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC", Package = "DIP-40" }],
        };

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.ShowPage(ComponentPage.Pinout);
            Dispatcher.UIThread.RunJobs();

            AutoCompleteBox box = Get<AutoCompleteBox>(pane, "PackageBox");
            TextBox text = box.GetVisualDescendants().OfType<TextBox>().First();

            text.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(box.IsDropDownOpen);

            ListBox list = box.GetVisualDescendants().OfType<Avalonia.Controls.Primitives.Popup>().Single().Child!
                .GetVisualDescendants().OfType<ListBox>().First();
            Dispatcher.UIThread.RunJobs();

            ListBoxItem item = list.GetVisualDescendants().OfType<ListBoxItem>().First(row => row.DataContext is string name && name != "DIP-40");
            string picked = (string)item.DataContext!;
            TopLevel listWindow = TopLevel.GetTopLevel(item)!;
            Point at = item.TranslatePoint(new Point(4, 4), listWindow)!.Value;

            listWindow.MouseDown(at, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            listWindow.MouseUp(at, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(picked, box.Text);
            Assert.False(box.IsDropDownOpen);

            // Left and entered again: the list opens, every case.
            pane.GetVisualDescendants().OfType<Button>().First(button => button.IsEffectivelyVisible && button.IsEffectivelyEnabled).Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.False(box.IsDropDownOpen);

            text.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(box.IsDropDownOpen);

            window.Close();
        });
    }

    // ###########################################################################################
    // A COMPONENT'S FILES AND LINKS ARE ITS WHOLE LISTS (owner request, 2026-11-03: "Each file or link will
    // have set a region (if any) and then the complete list can be viewed and sorted. So, it should work in
    // same was as 'Board: Files' and 'Board: Links'"): every row whatever region CRT shows, "Applies to" on
    // each, moved past any row, and a row given another region stays where it is.
    // ###########################################################################################
    [Fact]
    public void A_components_files_and_links_list_every_row_and_say_what_each_applies_to()
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
            new ComponentLinkEntry { BoardLabel = "U8", Name = "Second", Url = "https://example.org/2", Region = "NTSC" },
        ]);

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            ItemsControl links = Get<ItemsControl>(pane, "LinksItemsControl");

            // CRT shows PAL, and the NTSC rows are listed with the rest.
            Assert.Equal(["One", "Two", "NTSC"], pane.FilesForTests.Select(row => row.Name));
            Assert.Equal(["First", "Second"], links.Items.OfType<ComponentPaneLinkRow>().Select(row => row.Name));
            Assert.Equal("NTSC", pane.FilesForTests[2].AppliesTo);
            Assert.All(pane.FilesForTests, row => Assert.True(row.ShowsAppliesTo));
            Assert.True(Get<TextBlock>(pane, "FilesAppliesToHeading").IsVisible);
            Assert.True(Get<TextBlock>(pane, "LinksAppliesToHeading").IsVisible);

            // Another region shown changes nothing on either list.
            pane.UseRegion("NTSC");
            Assert.Equal(["One", "Two", "NTSC"], pane.FilesForTests.Select(row => row.Name));
            pane.UseRegion("PAL");

            // Moved past a row of the other region: the order is what the draft holds, one step.
            Assert.True(pane.MoveFile(pane.FilesForTests[2], -1));
            Assert.Equal(["One", "NTSC", "Two"], session.Content.Board.ComponentLocalFiles.Select(row => row.Name));
            Assert.Equal(DraftStepWording.ComponentFiles("U8"), session.UndoDescription);
            Assert.False(pane.MoveFile(pane.FilesForTests[0], -1));

            ComponentPaneLinkRow second = Assert.Single(links.Items.OfType<ComponentPaneLinkRow>(), row => row.Name == "Second");
            Assert.True(pane.MoveLink(second, -1));
            Assert.Equal(["Second", "First"], session.Content.Board.ComponentLinks.Select(row => row.Name));

            // "Applies to": the draft has it, and the row stays - nothing to say about it.
            ComponentPaneFileRow two = pane.FilesForTests[2];
            two.AppliesTo = "NTSC";
            pane.CommitHeld();

            Assert.Same(two, pane.FilesForTests[2]);
            Assert.Equal("NTSC", session.Content.Board.ComponentLocalFiles.Single(row => row.Name == "Two").Region);
            Assert.False(Get<TextBlock>(pane, "PaneMessageText").IsVisible);
            Assert.Equal(3, session.Content.Board.ComponentLocalFiles.Count);

            window.Close();
        });
    }

    // A row of the region NOT shown is changed where it stands, and the region's own stay as they are.
    [Fact]
    public void A_link_of_the_region_not_shown_is_changed_on_the_page_as_any_other()
    {
        BoardData board = Board();
        board.ComponentLinks.AddRange(
        [
            new ComponentLinkEntry { BoardLabel = "U8", Name = "PAL", Url = "https://example.org/1", Region = "PAL" },
            new ComponentLinkEntry { BoardLabel = "U8", Name = "NTSC", Url = "https://example.org/2", Region = "NTSC" },
        ]);

        DraftSession session = this.OpenDraft(board);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            ComponentPaneLinkRow row = Assert.Single(Get<ItemsControl>(pane, "LinksItemsControl").Items.OfType<ComponentPaneLinkRow>(), link => link.Name == "NTSC");
            row.Url = "https://example.org/changed";
            pane.CommitHeld();

            Assert.Equal(
                ["PAL|https://example.org/1|PAL", "NTSC|https://example.org/changed|NTSC"],
                session.Content.Board.ComponentLinks.Select(link => $"{link.Name}|{link.Url}|{link.Region}"));

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

            // The other region: the row stays where it is, as "Board: Links"' do.
            box.SelectedItem = BoardRegions.Ntsc;
            Dispatcher.UIThread.RunJobs();
            pane.CommitHeld();

            Assert.Equal(BoardRegions.Ntsc, Assert.Single(session.Content.Board.ComponentLinks).Region);
            Assert.Equal(BoardRegions.Ntsc, AppliesToBox(pane, "LinksItemsControl", 0).SelectedItem);
            Assert.False(Get<TextBlock>(pane, "PaneMessageText").IsVisible);

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
