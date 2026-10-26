using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// ComponentPane - a component edited in place on the Contribute tab (CRT 4.0 step 2, 2026-10-25:
// "A component opens in a pane, never in a maximized window ... Every change is saved within a
// second ... An edit in one shows up in all the others at once").
//
// The pane works on the draft held in memory (DraftSession): what is typed goes in as the box is
// left, Enter is pressed or typing pauses - each change one named step of the draft's history -
// and whatever another editor changes shows in the pane at once, landing on top of what the pane
// still held.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentPaneTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData Published() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510", Category = "IC", Description = "The processor" },
            new ComponentEntry { BoardLabel = "U1", Region = "PAL", FriendlyName = "VIC", Category = "IC" },
            new ComponentEntry { BoardLabel = "U1", Region = "NTSC", FriendlyName = "VIC", Category = "IC" },
        ],
        ComponentPins = [new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", Direction = "Ground" }],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.org/6510.pdf" }],
    };

    // The draft, as the published board - every change the pane makes is then a change of the draft.
    private DraftSession OpenDraft()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPaneTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, Published());

        return DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;
    }

    private static (ComponentPane Pane, Window Window) Shown(DraftSession session, string label = "U8", string region = "PAL")
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 1400 };
        window.Show();

        pane.Show(session, label, region, Published());
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    private static ComponentEntry U8(DraftSession session) => ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!;

    private static T Get<T>(ComponentPane pane, string name) where T : Control => pane.GetControl<T>(name);

    // The buttons inside an items list's rows, as drawn.
    private static IEnumerable<Button> RowButtons(ComponentPane pane, string list) =>
        Get<ItemsControl>(pane, list)
            .GetRealizedContainers()
            .SelectMany(container => container.GetVisualDescendants().OfType<Button>());

    private static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    [Fact]
    public void The_pane_shows_the_component_as_the_draft_holds_it()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal("U8 - CPU | 6510", Get<TextBlock>(pane, "TitleText").Text);
            Assert.Equal("CPU", Get<TextBox>(pane, "FriendlyNameBox").Text);
            Assert.Equal("6510", Get<TextBox>(pane, "TechnicalNameBox").Text);
            Assert.Equal("IC", Get<AutoCompleteBox>(pane, "CategoryBox").Text);
            Assert.Equal("The processor", Get<TextBox>(pane, "DescriptionBox").Text);
            Assert.Equal("13/100", Get<TextBlock>(pane, "DescriptionCounter").Text);

            // The draft is the published board: nothing changed, nothing to put back.
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.False(Get<Button>(pane, "UndoChangesButton").IsVisible);

            Assert.True(Get<Grid>(pane, "PinsHeading").IsVisible);
            Assert.True(Get<TextBlock>(pane, "NoReplacementsText").IsVisible);
            Assert.False(Get<TextBlock>(pane, "NoLinksText").IsVisible);
            Assert.False(Get<TextBlock>(pane, "RegionText").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // A box's change goes into the draft as ONE step named after the field - and once in, the pane
    // says the draft changes the component, with the way to put it back.
    // ###########################################################################################
    [Fact]
    public void A_field_goes_into_the_draft_as_one_named_step()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "DescriptionBox").Text = "Runs the show";
            pane.CommitHeld();

            Assert.Equal("Runs the show", U8(session).Description);
            Assert.Equal("Short description of [U8]", session.UndoDescription);
            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);

            Button undoChanges = Get<Button>(pane, "UndoChangesButton");
            Assert.True(undoChanges.IsVisible);
            Assert.Equal("Undo changes to [U8]", undoChanges.Content);

            // Nothing more held: no second step.
            pane.CommitHeld();
            Assert.True(session.Undo());
            Assert.False(session.CanUndo);

            window.Close();
        });
    }

    // Enter in a one-line box takes what it says into the draft at once.
    [Fact]
    public void Enter_in_a_one_line_box_takes_it_into_the_draft()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            TextBox name = Get<TextBox>(pane, "FriendlyNameBox");
            name.Focus();
            name.Text = "Processor";
            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.Equal("Processor", U8(session).FriendlyName);

            window.Close();
        });
    }

    // Leaving a box takes what it says into the draft.
    [Fact]
    public void Leaving_a_box_takes_it_into_the_draft()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            TextBox partNumber = Get<TextBox>(pane, "PartNumberBox");
            partNumber.Focus();
            partNumber.Text = "906122";
            Get<TextBox>(pane, "TechnicalNameBox").Focus();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("906122", U8(session).PartNumber);

            window.Close();
        });
    }

    // The short description is one line of the popup: past the limit the counter says so.
    [Fact]
    public void The_counter_warns_past_the_limit()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "DescriptionBox").Text = new string('x', ComponentVocabulary.ShortDescriptionLimit + 1);
            Dispatcher.UIThread.RunJobs();

            TextBlock counter = Get<TextBlock>(pane, "DescriptionCounter");
            Assert.Equal("101/100", counter.Text);
            Assert.Contains("over", counter.Classes);

            window.Close();
        });
    }

    // ###########################################################################################
    // The case picked, its pins: every number not listed yet added in number order - the named pin
    // kept - as the package's step and then the pins' step; the case drawn with them.
    // ###########################################################################################
    [Fact]
    public void Picking_a_case_offers_its_pins_and_adds_them_in_number_order()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<AutoCompleteBox>(pane, "PackageBox").Text = "DIP-8";
            pane.CommitHeld();

            Button addCasePins = Get<Button>(pane, "AddCasePinsButton");
            Assert.True(addCasePins.IsVisible);
            Assert.Equal("Add the 7 missing pins", addCasePins.Content);
            Assert.Contains("Accept", addCasePins.Classes);
            Assert.Equal("DIP-8", U8(session).Package);

            pane.AddCasePins();

            List<ComponentPinEntry> pins = ComponentPaneEdits.PinsOf(session.Content.Board, "U8", "PAL");
            Assert.Equal(["1", "2", "3", "4", "5", "6", "7", "8"], pins.Select(pin => pin.Pin));
            Assert.Equal("VSS", pins[0].Name);
            Assert.Equal("Pins of [U8]", session.UndoDescription);

            Assert.False(addCasePins.IsVisible);
            Assert.Equal(8, pane.GetControl<PinoutDrawing>("PinoutPicture").LayoutForTests!.Pins.Count);

            window.Close();
        });
    }

    // ###########################################################################################
    // A pin's direction picked goes in; a drop-down list's empty write (a list built before its value
    // arrived) is never taken for clearing it.
    // ###########################################################################################
    [Fact]
    public void A_pins_direction_goes_into_the_draft_and_an_empty_write_clears_nothing()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            ComponentPanePinRow row = Assert.Single(Get<ItemsControl>(pane, "PinsItemsControl").Items.OfType<ComponentPanePinRow>());

            row.Direction = null;
            Assert.Equal("Ground", row.Direction);

            row.Direction = "Power";
            pane.CommitHeld();

            Assert.Equal("Power", ComponentPaneEdits.PinsOf(session.Content.Board, "U8", "PAL").Single().Direction);

            window.Close();
        });
    }

    [Fact]
    public void Removing_a_pin_takes_it_out_of_the_draft()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Button remove = RowButtons(pane, "PinsItemsControl").Single(button => Equals(button.Content, "Remove"));
            Assert.Contains("Danger", remove.Classes);
            Click(remove);

            Assert.Empty(ComponentPaneEdits.PinsOf(session.Content.Board, "U8", "PAL"));
            Assert.True(Get<TextBlock>(pane, "NoPinsText").IsVisible);

            window.Close();
        });
    }

    // A replacement part and a link added, filled and removed - each into the draft.
    [Fact]
    public void Replacement_parts_and_links_are_added_filled_and_removed()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Click(Get<Button>(pane, "AddReplacementButton"));
            ComponentPaneReplacementRow part = Get<ItemsControl>(pane, "ReplacementsItemsControl").Items.OfType<ComponentPaneReplacementRow>().Single();
            part.PartNumber = "8500";
            part.Kind = "Drop-in";

            Click(Get<Button>(pane, "AddLinkButton"));
            ComponentPaneLinkRow link = Get<ItemsControl>(pane, "LinksItemsControl").Items.OfType<ComponentPaneLinkRow>().Last();
            link.Name = "Pinout";
            link.Url = "https://example.org/pinout";

            pane.CommitHeld();

            ComponentReplacementEntry saved = Assert.Single(ComponentPaneEdits.ReplacementsOf(session.Content.Board, "U8", "PAL"));
            Assert.Equal("8500|Drop-in", $"{saved.PartNumber}|{saved.Kind}");
            Assert.Equal(["Datasheet", "Pinout"], ComponentPaneEdits.LinksOf(session.Content.Board, "U8").Select(row => row.Name));

            Dispatcher.UIThread.RunJobs();
            Click(RowButtons(pane, "ReplacementsItemsControl").Single(button => Equals(button.Content, "Remove")));
            Assert.Empty(ComponentPaneEdits.ReplacementsOf(session.Content.Board, "U8", "PAL"));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** ANOTHER EDITOR'S CHANGE LANDS ON TOP OF WHAT THE PANE HOLDS. *** What the boxes held goes in
    // first (the draft's ChangeComing), and the pane then reads the component again - neither change
    // is lost.
    // ###########################################################################################
    [Fact]
    public void Another_editors_change_lands_on_top_of_what_the_pane_holds()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "DescriptionBox").Text = "Typed, not yet in";

            session.Apply(
                "Table edit",
                content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.FriendlyName, "Microprocessor") },
                origin: new object());

            Assert.Equal("Typed, not yet in", U8(session).Description);
            Assert.Equal("Microprocessor", U8(session).FriendlyName);
            Assert.Equal("Microprocessor", Get<TextBox>(pane, "FriendlyNameBox").Text);
            Assert.Equal("Typed, not yet in", Get<TextBox>(pane, "DescriptionBox").Text);

            window.Close();
        });
    }

    // The draft's Undo - from its header - shows in the pane at once.
    [Fact]
    public void An_undo_shows_the_component_as_it_was()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "DescriptionBox").Text = "Runs the show";
            pane.CommitHeld();

            session.Undo();

            Assert.Equal("The processor", Get<TextBox>(pane, "DescriptionBox").Text);
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // An undo while typing is held: what was typed goes in first, as the latest step - so it is what
    // the undo takes back, as Ctrl+Z in a text box would - and the box shows the field as it was.
    // ###########################################################################################
    [Fact]
    public void An_undo_while_typing_is_held_takes_back_what_was_typed()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "FriendlyNameBox").Text = "Processor";
            pane.CommitHeld();
            Get<TextBox>(pane, "DescriptionBox").Text = "Typed, not yet in";

            session.Undo();

            Assert.Equal("The processor", U8(session).Description);
            Assert.Equal("The processor", Get<TextBox>(pane, "DescriptionBox").Text);
            Assert.Equal("Processor", Get<TextBox>(pane, "FriendlyNameBox").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Undo changes to [U8]": every change to the component taken back as ONE step - which the
    // draft's Undo brings back.
    // ###########################################################################################
    [Fact]
    public void Undo_changes_puts_the_component_back_as_published_in_one_step()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "FriendlyNameBox").Text = "Processor";
            Get<AutoCompleteBox>(pane, "PackageBox").Text = "DIP-40";
            pane.CommitHeld();
            pane.AddCasePins();

            pane.UndoComponentChanges();

            Assert.Equal("CPU", U8(session).FriendlyName);
            Assert.Single(ComponentPaneEdits.PinsOf(session.Content.Board, "U8", "PAL"));
            Assert.Equal("CPU", Get<TextBox>(pane, "FriendlyNameBox").Text);
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.Equal("[U8] back to its published values", session.UndoDescription);

            session.Undo();
            Assert.Equal("Processor", U8(session).FriendlyName);
            Assert.Equal(40, ComponentPaneEdits.PinsOf(session.Content.Board, "U8", "PAL").Count);

            window.Close();
        });
    }

    // ###########################################################################################
    // A picture replaced under its own name changes the component (code review, 2026-10-26): the pane
    // says so, and "Undo changes to [U8]" puts the published picture back in the draft's list.
    // ###########################################################################################
    [Fact]
    public async Task A_picture_replaced_under_its_own_name_is_changed_and_undo_changes_puts_it_back()
    {
        BoardData published = Published();
        published.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/u8.png" });

        string publishedBoardFile = Path.Combine(this.thisWorkspace.Root, "Data", "Commodore", "C64", "250407", DataTreeLayout.BoardFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(publishedBoardFile)!);
        var theirs = new BoardDocument.StoredFileRow { Name = "Component images/u8.png", Sha256 = new string('a', 64), Size = 9 };
        BoardFiles.Write(publishedBoardFile, BoardDocumentMapping.ToDocument(new BoardContent(published, [], []), name => name, [theirs]));

        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPaneTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        BoardDocument.StoredFileRow mine = StoredFiles.Put(Path.GetDirectoryName(boardFile)!, "Component images/u8.png", "the contributor's picture");
        BoardFiles.Write(boardFile, BoardDocumentMapping.ToDocument(new BoardContent(published, [], []), name => name, [mine]));
        DraftSession session = DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;

        await UiTest.RunAsync(async () =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();
            pane.Show(session, "U8", "PAL", published, publishedBoardFile);
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.True(Get<Button>(pane, "UndoChangesButton").IsVisible);

            await pane.UndoComponentChanges();

            // The file read back is said from the pool; the pane shows it on the UI thread.
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(theirs.Sha256, Assert.Single(BoardFiles.ReadDocument(boardFile)!.Files).Sha256);
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // A write of the draft's file run off the UI thread - the KiCad data window's Remove, "Bring my
    // draft up to date" (DraftSession.WriteFileDirectly) - raises Changed on the thread it ran on.
    // The pane drew from there and threw, failing the write it was told about (found 2026-10-26,
    // while "Undo changes" came to put files back). It shows the component again on the UI thread.
    // ###########################################################################################
    [Fact]
    public async Task A_write_of_the_file_run_off_the_UI_thread_is_shown_on_it()
    {
        DraftSession session = this.OpenDraft();
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPaneTests.BoardKey);
        BoardData written = Published();
        written.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "Processor", TechnicalNameOrValue = "6510", Category = "IC" };

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            bool done = await Task.Run(() => session.WriteFileDirectly(
                () =>
                {
                    BoardFiles.Write(boardFile, BoardDocumentMapping.ToDocument(new BoardContent(written, [], []), name => name, []));
                    return true;
                },
                succeeded => succeeded,
                "Written elsewhere",
                clearHistory: false));

            Dispatcher.UIThread.RunJobs();

            Assert.True(done);
            Assert.Equal("Processor", Get<TextBox>(pane, "FriendlyNameBox").Text);

            window.Close();
        });
    }

    // A component the draft no longer has - its adding undone, say - is said, not shown empty.
    [Fact]
    public void A_component_the_draft_no_longer_has_says_so()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            session.Apply(
                "Deleted elsewhere",
                content => content with { Board = ComponentBoardWriter.ApplyComponentDelete(content.Board, "U8") },
                origin: new object());

            Assert.True(Get<StackPanel>(pane, "GonePanel").IsVisible);
            Assert.False(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);
            Assert.Contains("[U8]", Get<TextBlock>(pane, "GoneText").Text, StringComparison.Ordinal);

            window.Close();
        });
    }

    // A component with a row of its own per region: which one is open is said, and a new pin is the
    // region's.
    [Fact]
    public void A_regional_component_says_its_region_and_its_new_pins_are_that_regions()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U1", "NTSC");

            Assert.True(Get<TextBlock>(pane, "RegionText").IsVisible);
            Assert.StartsWith("The NTSC version", Get<TextBlock>(pane, "RegionText").Text, StringComparison.Ordinal);

            Click(Get<Button>(pane, "AddPinButton"));
            ComponentPanePinRow pin = Get<ItemsControl>(pane, "PinsItemsControl").Items.OfType<ComponentPanePinRow>().Single();
            pin.Pin = "1";
            pin.Name = "D6";
            pane.CommitHeld();

            Assert.Equal("NTSC", session.Content.Board.ComponentPins.Single(row => row.BoardLabel == "U1").Region);

            window.Close();
        });
    }

    // ###########################################################################################
    // The pane saves the draft itself - here when asked, as before submitting or quitting - and says
    // which draft it saved.
    // ###########################################################################################
    [Fact]
    public async Task The_pane_saves_the_draft_and_says_so()
    {
        DraftSession session = this.OpenDraft();
        var saved = new List<string>();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.Saved += (_, excelDataFile) => saved.Add(excelDataFile);

            Get<TextBox>(pane, "DescriptionBox").Text = "Runs the show";

            Assert.True(await pane.SaveNowAsync());

            window.Close();
        });

        Assert.Equal([ComponentPaneTests.BoardKey], saved);

        BoardContent written = BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPaneTests.BoardKey))!;
        Assert.Equal("Runs the show", ComponentPaneEdits.RowShown(written.Board, "U8", "PAL")!.Description);
    }

    // Closing hands what the boxes hold to the draft; letting go of a draft that is going does not.
    [Fact]
    public async Task Closing_takes_what_the_boxes_hold_and_forgetting_does_not()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Get<TextBox>(pane, "FriendlyNameBox").Text = "Processor";
            pane.Close();

            Assert.Equal("Processor", U8(session).FriendlyName);
            Assert.False(pane.IsShowing);

            // Closing saves what it handed over - finished here, so nothing runs on into the next test.
            await pane.LastSave;
            Assert.False(session.HasUnsavedChanges);

            pane.Show(session, "U8", "PAL", Published());
            Get<TextBox>(pane, "FriendlyNameBox").Text = "Not kept";
            pane.Forget();

            Assert.Equal("Processor", U8(session).FriendlyName);
            Assert.False(pane.IsShowing);

            window.Close();
        });
    }

    // The two buttons labelled in code carry their colours (CLAUDE.md, button colours).
    [Fact]
    public void The_buttons_labelled_in_code_carry_their_colours()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Contains("Danger", Get<Button>(pane, "UndoChangesButton").Classes);
            Assert.Contains("Accept", Get<Button>(pane, "AddCasePinsButton").Classes);
            Assert.DoesNotContain("Accept", Get<Button>(pane, "CloseButton").Classes);
            Assert.DoesNotContain("Danger", Get<Button>(pane, "CloseButton").Classes);

            window.Close();
        });
    }
}
