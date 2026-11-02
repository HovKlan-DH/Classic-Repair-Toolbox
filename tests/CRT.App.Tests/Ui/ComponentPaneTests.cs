using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

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

            Assert.Equal("Component: U8 - CPU | 6510", Get<TextBlock>(pane, "TitleText").Text);
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

            pane.ShowPage(ComponentPage.Pinout);
            Dispatcher.UIThread.RunJobs();

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
            Assert.Equal(["Datasheet", "Pinout"], ComponentPaneEdits.LinksOf(session.Content.Board, "U8", "PAL").Select(row => row.Name));

            pane.ShowPage(ComponentPage.Replacements);
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
    public void A_regional_component_says_its_region_and_its_new_pins_are_for_every_region_until_told_apart()
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

            // Its own row per region says nothing of its pins (owner decision, 2026-11-01): the tick does.
            Assert.Equal(string.Empty, session.Content.Board.ComponentPins.Single(row => row.BoardLabel == "U1").Region);

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

    // ------------------------------------------------------------------ Scope baselines, pictures, files (slice 3b)

    // A 1 x 1 PNG - enough for a thumbnail to be decoded from.
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private string BoardFile => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPaneTests.BoardKey);

    private string DraftFolder => Path.GetDirectoryName(this.BoardFile)!;

    // `name`'s bytes in the draft's own store, and its entry.
    private BoardDocument.StoredFileRow Stored(string name, byte[] bytes)
    {
        string sha = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes));
        string path = Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, bytes);

        return new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = bytes.Length };
    }

    // The published board with U8's picture, one scope baseline and a datasheet.
    private static BoardData PublishedWithFiles()
    {
        BoardData board = Published();
        board.ComponentImages.AddRange([
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/U8.png" },
            new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/U8_1.png", TimeDiv = "50us", VoltsDiv = "1V", TriggerLevelVolts = "1.5V" },
        ]);
        board.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Component local files/6510.pdf" });
        return board;
    }

    // That board as the draft, its files stored in the draft.
    private DraftSession OpenDraftWithFiles()
    {
        BoardData board = PublishedWithFiles();

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [
                this.Stored("Component images/U8.png", ComponentPaneTests.Png),
                this.Stored("Scope baseline/U8_1.png", ComponentPaneTests.Png),
                this.Stored("Component local files/6510.pdf", [37, 80, 68, 70]),
            ]));

        return DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;
    }

    // An uncompressed 24-bit BMP of `width` x `height` - a picture type the PNG/JPEG reader does not know.
    private static byte[] Bmp(int width, int height)
    {
        int rowSize = ((width * 3) + 3) & ~3;
        int pixels = rowSize * height;
        var bmp = new byte[54 + pixels];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        BitConverter.GetBytes(bmp.Length).CopyTo(bmp, 2);
        BitConverter.GetBytes(54).CopyTo(bmp, 10);
        BitConverter.GetBytes(40).CopyTo(bmp, 14);
        BitConverter.GetBytes(width).CopyTo(bmp, 18);
        BitConverter.GetBytes(height).CopyTo(bmp, 22);
        BitConverter.GetBytes((short)1).CopyTo(bmp, 26);
        BitConverter.GetBytes((short)24).CopyTo(bmp, 28);
        BitConverter.GetBytes(pixels).CopyTo(bmp, 34);
        return bmp;
    }

    // ###########################################################################################
    // A thumbnail is never another board's (code review, 2026-10-26): kept by file NAME, a load still
    // running when the pane moved to another board stored its picture after the move - and that
    // board's file of the same name showed it at its next re-read.
    // ###########################################################################################
    [Fact]
    public async Task A_thumbnail_read_for_another_board_never_shows_for_a_file_of_the_same_name()
    {
        DraftSession first = this.OpenDraftWithFiles();

        const string OtherKey = "Commodore/C64/250425/board.json";
        string otherFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, OtherKey);
        BoardData otherBoard = Published();
        otherBoard.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/U8_1.png", TimeDiv = "50us" });
        Directory.CreateDirectory(Path.GetDirectoryName(otherFile)!);
        byte[] wide = ComponentPaneTests.Bmp(8, 2);
        string sha = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(wide));
        string stored = Path.Combine(Path.GetDirectoryName(otherFile)!, DataTreeLayout.StoredFilePath(sha, "Scope baseline/U8_1.png").Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(stored)!);
        File.WriteAllBytes(stored, wide);
        BoardFiles.Write(otherFile, BoardDocumentMapping.ToDocument(
            new BoardContent(otherBoard, [], []),
            name => name,
            [new BoardDocument.StoredFileRow { Name = "Scope baseline/U8_1.png", Sha256 = sha, Size = wide.Length }]));
        DraftSession second = DraftSession.Open(this.DraftsRoot, OtherKey)!;

        await UiTest.RunAsync(async () =>
        {
            var held = new TaskCompletionSource();

            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            // The first board's pictures decoded are held - their loads still running.
            pane.ImageDecodedForTests = _ => held.Task;
            pane.Show(first, "U8", "PAL", Published());
            Dispatcher.UIThread.RunJobs();

            pane.ImageDecodedForTests = null;
            pane.Show(second, "U8", "PAL", Published());
            Dispatcher.UIThread.RunJobs();

            for (int wait = 0; wait < 200 && pane.BaselinesForTests.Single().Thumbnail is null; wait++)
            {
                await Task.Delay(10);
                Dispatcher.UIThread.RunJobs();
            }

            // The second board's own picture, read for it.
            Avalonia.Media.Imaging.Bitmap secondBoards = pane.BaselinesForTests.Single().Thumbnail!;

            held.SetResult();
            await pane.ThumbnailsLoadedForTests;

            // A change somewhere else: the second board's rows read again.
            Assert.True(second.Apply("Elsewhere", content => content with { Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Board data", NameOrHandle = "Bo" }]) }));
            Dispatcher.UIThread.RunJobs();
            await pane.ThumbnailsLoadedForTests;

            Assert.Same(secondBoards, pane.BaselinesForTests.Single().Thumbnail);

            window.Close();
        });
    }

    // A file to pick, lying outside the draft as the contributor's own.
    private string Pickable(string fileName, byte[] bytes)
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Picked", fileName);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    private List<string> ListedOnDisk() => BoardFiles.ReadDocument(this.BoardFile)!.Files.Select(file => file.Name).ToList();

    // ###########################################################################################
    // The component's scope baselines, pictures and files are in the pane (slice 3b - the window
    // edited them until now), counted in their headings, the pictures with their thumbnails.
    // ###########################################################################################
    [Fact]
    public async Task The_pane_shows_the_components_baselines_pictures_and_files()
    {
        DraftSession session = this.OpenDraftWithFiles();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            await pane.ThumbnailsLoadedForTests;

            Assert.Equal("Oscilloscope baseline (1)", Get<TextBlock>(pane, "BaselinesPageLabel").Text);
            Assert.Equal("Pictures (1)", Get<TextBlock>(pane, "PicturesPageLabel").Text);
            Assert.Equal("Files (1)", Get<TextBlock>(pane, "FilesPageLabel").Text);

            ComponentPaneBaselineRow baseline = Assert.Single(pane.BaselinesForTests);
            Assert.Equal(("1", "VSS", "50us", "1V", "1.5V"), (baseline.Pin, baseline.Name, baseline.TimeDiv, baseline.VoltsDiv, baseline.TriggerLevel));
            Assert.NotNull(baseline.Thumbnail);
            Assert.NotNull(Assert.Single(pane.PicturesForTests).Thumbnail);
            Assert.Equal("Component local files/6510.pdf", Assert.Single(pane.FilesForTests).File);

            Assert.False(Get<TextBlock>(pane, "NoBaselinesText").IsVisible);
            Assert.False(Get<Border>(pane, "NewBaselinePanel").IsVisible);

            window.Close();
        });
    }

    // A baseline's field goes into the draft as one step, as every other section's does.
    [Fact]
    public void A_baselines_field_goes_into_the_draft_as_one_step()
    {
        DraftSession session = this.OpenDraftWithFiles();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            ComponentPaneBaselineRow baseline = pane.BaselinesForTests[0];
            baseline.TimeDiv = "20us";
            baseline.MachineState = "At the READY prompt";
            pane.CommitHeld();

            ComponentImageEntry saved = Assert.Single(ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL"));
            Assert.Equal(("20us", "At the READY prompt"), (saved.TimeDiv, saved.MachineState));
            Assert.Equal("Oscilloscope baselines of [U8]", session.UndoDescription);

            window.Close();
        });
    }

    // ###########################################################################################
    // An edit that would make a baseline no baseline - its pin cleared - is held back: every reader
    // would take the row for a picture from then on. The draft keeps it as it was, the reason said
    // under it; given its pin again, it goes in.
    // ###########################################################################################
    [Fact]
    public void A_baseline_edit_that_would_make_it_a_picture_is_held_back_with_the_reason()
    {
        DraftSession session = this.OpenDraftWithFiles();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            ComponentPaneBaselineRow baseline = pane.BaselinesForTests[0];
            baseline.Pin = " ";
            baseline.Name = "Ground";
            pane.CommitHeld();

            Assert.Equal("VSS", Assert.Single(ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL")).Name);
            Assert.True(baseline.HasProblem);
            Assert.StartsWith("An oscilloscope baseline needs its pin.", baseline.Problem, StringComparison.Ordinal);

            baseline.Pin = "1";
            pane.CommitHeld();

            Assert.Equal("Ground", Assert.Single(ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL")).Name);
            Assert.False(baseline.HasProblem);

            window.Close();
        });
    }

    // ###########################################################################################
    // A PICTURE ADDED GOES INTO THE DRAFT AND IS SAVED AT ONCE (slice 3b) - copied in under the
    // board's own folder for pictures, listed with its bytes in the draft's store, nothing left lying
    // in the folder; and never under a name the board has: that would replace that file for every row.
    // ###########################################################################################
    [Fact]
    public async Task A_picture_added_is_copied_in_under_a_free_name_and_saved_at_once()
    {
        DraftSession session = this.OpenDraftWithFiles();
        string picked = this.Pickable("U8.png", ComponentPaneTests.Png);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.PickFileOverrideForTests = kind => Task.FromResult<string?>(kind == ComponentPane.PaneFileKind.Picture ? picked : null);

            await pane.AddPictureAsync();

            Assert.Equal(["Top", "U8"], ComponentPaneEdits.PicturesOf(session.Content.Board, "U8", "PAL").Select(row => row.Name));
            Assert.Equal("Component images/U8 (2).png", pane.PicturesForTests[1].File);
            Assert.Equal("Pictures of [U8]", session.UndoDescription);

            Assert.False(session.HasUnsavedChanges);
            Assert.Contains("Component images/U8 (2).png", this.ListedOnDisk());
            Assert.False(File.Exists(Path.Combine(this.DraftFolder, "Component images", "U8 (2).png")));
            Assert.True(File.Exists(picked));

            window.Close();
        });
    }

    // A picked file CRT cannot show is no picture - said, and nothing changes.
    [Fact]
    public async Task A_picked_file_that_is_no_picture_is_refused_and_changes_nothing()
    {
        DraftSession session = this.OpenDraftWithFiles();
        string picked = this.Pickable("notes.txt", [65, 66]);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);

            await pane.AddPictureAsync();

            Assert.Single(pane.PicturesForTests);
            Assert.False(session.CanUndo);
            Assert.True(Get<TextBlock>(pane, "PaneMessageText").IsVisible);
            Assert.Equal(ComponentPaneWording.NotAnImage, Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // A NEW SCOPE BASELINE waits in its panel until it can be one - its pin, a screenshot and one
    // scope setting - starting from the component's last baseline's settings; added, its screenshot
    // is named as the published ones are and the draft saved.
    // ###########################################################################################
    [Fact]
    public async Task A_new_baseline_waits_until_it_has_its_pin_screenshot_and_setting_then_goes_in()
    {
        DraftSession session = this.OpenDraftWithFiles();
        string screenshot = this.Pickable("IMG_0042.png", ComponentPaneTests.Png);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.PickFileOverrideForTests = kind => Task.FromResult<string?>(kind == ComponentPane.PaneFileKind.Screenshot ? screenshot : null);

            pane.OpenNewBaseline();

            Assert.True(Get<Border>(pane, "NewBaselinePanel").IsVisible);
            Assert.False(Get<StackPanel>(pane, "AddBaselineButtons").IsVisible);
            Assert.Equal("1V", Get<AutoCompleteBox>(pane, "NewBaselineVoltsDivBox").Text);
            Assert.Equal("1.5V", Get<TextBox>(pane, "NewBaselineTriggerBox").Text);
            Assert.False(Get<Button>(pane, "AddNewBaselineButton").IsEnabled);
            Assert.Equal("An oscilloscope baseline needs its pin, its signal name and an image.", Get<TextBlock>(pane, "NewBaselineProblemText").Text);

            Get<TextBox>(pane, "NewBaselinePinBox").Text = "2";
            Get<TextBox>(pane, "NewBaselineNameBox").Text = "PA3";
            await pane.ChooseNewBaselineScreenshotAsync();

            Assert.True(Get<Button>(pane, "AddNewBaselineButton").IsEnabled);

            await pane.AddNewBaselineAsync();

            ComponentImageEntry added = ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL")[^1];
            Assert.Equal(("2", "PA3", "Scope baseline/U8_2_PA3.png", "1V"), (added.Pin, added.Name, added.File, added.VoltsDiv));
            Assert.Contains("Scope baseline/U8_2_PA3.png", this.ListedOnDisk());
            Assert.False(Get<Border>(pane, "NewBaselinePanel").IsVisible);
            Assert.Equal("Oscilloscope baseline (2)", Get<TextBlock>(pane, "BaselinesPageLabel").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // "CAPTURE FROM MY OSCILLOSCOPE" is offered while the Oscilloscope tab has a session, and its
    // screen becomes the new baseline's screenshot; a capture that brings nothing says so.
    // ###########################################################################################
    [Fact]
    public async Task Capture_from_the_oscilloscope_is_offered_while_connected_and_starts_a_baseline()
    {
        DraftSession session = this.OpenDraftWithFiles();
        byte[]? screen = ComponentPaneTests.Png;

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.False(Get<Button>(pane, "CaptureBaselineButton").IsVisible);

            pane.CaptureScopeImage = _ => Task.FromResult<byte[]?>(screen);
            pane.ShowScopeConnected(true);
            Assert.True(Get<Button>(pane, "CaptureBaselineButton").IsVisible);

            pane.OpenNewBaseline();
            Assert.True(Get<Button>(pane, "NewBaselineCaptureButton").IsVisible);

            screen = null;
            await pane.CaptureNewBaselineAsync();
            Assert.Equal(ComponentPaneWording.CaptureFailed, Get<TextBlock>(pane, "PaneMessageText").Text);
            Assert.False(Get<Button>(pane, "AddNewBaselineButton").IsEnabled);

            screen = ComponentPaneTests.Png;
            Get<TextBox>(pane, "NewBaselinePinBox").Text = "3";
            Get<TextBox>(pane, "NewBaselineNameBox").Text = "PA1";
            await pane.CaptureNewBaselineAsync();

            // Adding it deletes the capture, which its thumbnail may still be reading - left in the
            // user's own folder, the delete being refused on Windows.
            await pane.ThumbnailsLoadedForTests;
            await pane.AddNewBaselineAsync();

            Assert.Equal("Scope baseline/U8_3_PA1.png", ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL")[^1].File);

            pane.ShowScopeConnected(false);
            Assert.False(Get<Button>(pane, "CaptureBaselineButton").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // A capture waits for "Add oscilloscope image" in the user's OWN app data folder (code review,
    // 2026-10-26): under the temporary folder - the shared /tmp on Linux, at a path anyone can work
    // out - another account could swap the screenshot before it was added to the draft and sent.
    // ###########################################################################################
    [Fact]
    public async Task A_capture_waits_in_the_users_own_folder_not_the_shared_temporary_one()
    {
        DraftSession session = this.OpenDraftWithFiles();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.CaptureScopeImage = _ => Task.FromResult<byte[]?>(ComponentPaneTests.Png);
            pane.ShowScopeConnected(true);
            pane.OpenNewBaseline();

            await pane.CaptureNewBaselineAsync();

            string waiting = pane.NewBaselineSourceForTests!;
            string ownFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppConfig.AppFolderName);
            Assert.StartsWith(ownFolder + Path.DirectorySeparatorChar, waiting, StringComparison.Ordinal);
            Assert.False(waiting.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase));

            // Its thumbnail is read off the UI thread, and on Windows a file being read cannot be deleted.
            await pane.ThumbnailsLoadedForTests;

            window.Close();
            File.Delete(waiting);
        });
    }

    // A file added and one removed, each one step - a file of any type, under the board's folder for them.
    [Fact]
    public async Task A_file_is_added_under_the_files_folder_and_removed_as_one_step()
    {
        DraftSession session = this.OpenDraftWithFiles();
        string picked = this.Pickable("6510 errata.txt", [65, 66]);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);

            await pane.AddFileAsync();

            Assert.Equal(["Datasheet", "6510 errata"], ComponentPaneEdits.FilesOf(session.Content.Board, "U8", "PAL").Select(row => row.Name));
            Assert.Contains("Component local files/6510 errata.txt", this.ListedOnDisk());
            Assert.Equal("Files of [U8]", session.UndoDescription);

            pane.FilesForTests[1].Name = "Errata";
            pane.CommitHeld();
            Assert.Equal("Errata", ComponentPaneEdits.FilesOf(session.Content.Board, "U8", "PAL")[1].Name);

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
            Assert.Contains("Danger", Get<Button>(pane, "DeleteComponentButton").Classes);
            Assert.Contains("Accept", Get<Button>(pane, "BringBackButton").Classes);

            // No Close any more (owner request, 2026-10-30) - the left column's "Board status" goes back.
            Assert.Null(pane.FindControl<Button>("CloseButton"));

            window.Close();
        });
    }

    // ---- the component as a whole (slice 3c) --------------------------------------------------

    private static (ComponentPane Pane, Window Window) ShownNew(DraftSession session, string region = "PAL")
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 1400 };
        window.Show();

        pane.ShowNew(session, region, Published());
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    // ###########################################################################################
    // "Add component": a form for its label (and region, category) - nothing goes into the draft
    // until "Add component" there, and then as ONE step, the pane showing the new component.
    // ###########################################################################################
    [Fact]
    public void A_new_component_goes_into_the_draft_as_one_step_once_added_and_is_shown()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownNew(session);

            Assert.True(pane.IsNewComponentForm);
            Assert.Equal("New component", Get<TextBlock>(pane, "TitleText").Text);
            Assert.True(Get<ScrollViewer>(pane, "NewComponentPanel").IsVisible);
            Assert.False(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);
            Assert.False(Get<Button>(pane, "AddNewComponentButton").IsEnabled);

            Get<TextBox>(pane, "NewLabelBox").Text = " U9 ";
            Get<AutoCompleteBox>(pane, "NewCategoryBox").Text = "IC";
            Dispatcher.UIThread.RunJobs();

            // Typed, not added: the draft is as it was.
            pane.CommitHeld();
            Assert.False(session.CanUndo);
            Assert.True(Get<Button>(pane, "AddNewComponentButton").IsEnabled);

            Click(Get<Button>(pane, "AddNewComponentButton"));

            Assert.False(pane.IsNewComponentForm);
            Assert.Equal("[U9] added", session.UndoDescription);
            ComponentEntry added = ComponentPaneEdits.RowShown(session.Content.Board, "U9", "PAL")!;
            Assert.Equal("IC", added.Category);
            Assert.Equal(string.Empty, added.Region);
            Assert.Equal("U9", pane.Label);
            Assert.Equal("Component: U9", Get<TextBlock>(pane, "TitleText").Text);
            Assert.True(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);
            Assert.False(Get<ScrollViewer>(pane, "NewComponentPanel").IsVisible);
            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);

            // What is typed next is a change of the new component.
            Get<TextBox>(pane, "FriendlyNameBox").Text = "Glue logic";
            pane.CommitHeld();
            Assert.Equal("Glue logic", ComponentPaneEdits.RowShown(session.Content.Board, "U9", "PAL")!.FriendlyName);

            window.Close();
        });
    }

    // A label the board has - in any region, in any capitals - is refused with the reason; no category
    // is said and added all the same (unfinished work is fine).
    [Fact]
    public void A_label_taken_is_refused_and_no_category_is_said_and_adds_nothing()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownNew(session);

            Get<TextBox>(pane, "NewLabelBox").Text = "u1";
            Dispatcher.UIThread.RunJobs();

            Assert.False(Get<Button>(pane, "AddNewComponentButton").IsEnabled);
            Assert.True(Get<TextBlock>(pane, "NewLabelProblemText").IsVisible);
            Assert.Contains("already has [u1]", Get<TextBlock>(pane, "NewLabelProblemText").Text, StringComparison.Ordinal);
            Assert.False(pane.AddNewComponent());

            Get<TextBox>(pane, "NewLabelBox").Text = "U20";
            Dispatcher.UIThread.RunJobs();

            // Owner decision, 2026-11-01: "No components must have an empty category".
            Assert.False(Get<TextBlock>(pane, "NewLabelProblemText").IsVisible);
            Assert.True(Get<TextBlock>(pane, "NewCategoryNoteText").IsVisible);
            Assert.False(Get<Button>(pane, "AddNewComponentButton").IsEnabled);
            Assert.False(pane.AddNewComponent());

            Get<AutoCompleteBox>(pane, "NewCategoryBox").Text = "IC";
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<Button>(pane, "AddNewComponentButton").IsEnabled);
            Assert.True(pane.AddNewComponent());
            Assert.Equal("IC", ComponentPaneEdits.RowShown(session.Content.Board, "U20", "PAL")!.Category);

            window.Close();
        });
    }

    // ###########################################################################################
    // The region is asked only on a board with regional components. A component of ONE region added
    // while CRT shows the other is worked on in its own region - in the other, every section read back
    // empty, so each commit would file its pins again beside the ones already in.
    // ###########################################################################################
    [Fact]
    public void A_component_of_the_other_region_is_worked_on_in_its_own_region()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownNew(session, region: "PAL");

            Assert.True(Get<StackPanel>(pane, "NewRegionPanel").IsVisible);

            Get<TextBox>(pane, "NewLabelBox").Text = "U30";
            Get<AutoCompleteBox>(pane, "NewCategoryBox").Text = "IC";
            Get<ComboBox>(pane, "NewRegionBox").SelectedItem = "NTSC only";
            Assert.True(pane.AddNewComponent());
            Assert.Equal("NTSC", ComponentPaneEdits.RowShown(session.Content.Board, "U30", "NTSC")!.Region);

            Click(Get<Button>(pane, "AddPinButton"));
            ComponentPanePinRow pin = Get<ItemsControl>(pane, "PinsItemsControl").Items.OfType<ComponentPanePinRow>().Single();
            pin.Pin = "1";
            pin.Name = "CLK";
            pane.CommitHeld();
            pin.Name = "CLK in";
            pane.CommitHeld();

            ComponentPinEntry filed = Assert.Single(session.Content.Board.ComponentPins, row => row.BoardLabel == "U30");
            Assert.Equal("NTSC", filed.Region);
            Assert.Equal("CLK in", filed.Name);

            window.Close();
        });
    }

    // The same for a component of one region OPENED while CRT shows the other: the pane works in the
    // row's own region, so a pin changed twice is one pin.
    [Fact]
    public void A_component_of_one_region_opened_in_the_other_files_its_pins_once()
    {
        DraftSession session = this.OpenDraft();
        session.Apply(
            "Added",
            content => content with { Board = ComponentPaneEdits.WithNewComponent(content.Board, "U31", "NTSC", "IC") },
            origin: new object());

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U31", "PAL");

            Click(Get<Button>(pane, "AddPinButton"));
            ComponentPanePinRow pin = Get<ItemsControl>(pane, "PinsItemsControl").Items.OfType<ComponentPanePinRow>().Single();
            pin.Pin = "1";
            pin.Name = "CLK";
            pane.CommitHeld();
            pin.Name = "CLK in";
            pane.CommitHeld();

            ComponentPinEntry filed = Assert.Single(session.Content.Board.ComponentPins, row => row.BoardLabel == "U31");
            Assert.Equal("NTSC", filed.Region);
            Assert.Equal("CLK in", filed.Name);

            window.Close();
        });
    }

    [Fact]
    public void A_board_without_regional_components_asks_no_region()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPaneTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] });
        DraftSession session = DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownNew(session);

            Assert.False(Get<StackPanel>(pane, "NewRegionPanel").IsVisible);

            Get<TextBox>(pane, "NewLabelBox").Text = "U9";
            Get<AutoCompleteBox>(pane, "NewCategoryBox").Text = "IC";
            Assert.True(pane.AddNewComponent());
            Assert.Equal(string.Empty, ComponentPaneEdits.RowShown(session.Content.Board, "U9", "PAL")!.Region);

            window.Close();
        });
    }

    // Enter in the label adds the component, as the button does.
    [Fact]
    public void Enter_in_the_new_label_adds_the_component()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownNew(session);

            TextBox label = Get<TextBox>(pane, "NewLabelBox");
            label.Text = "U40";
            Get<AutoCompleteBox>(pane, "NewCategoryBox").Text = "IC";
            label.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter, Source = label });

            Assert.Equal("[U40] added", session.UndoDescription);
            Assert.False(pane.IsNewComponentForm);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Delete component": every row of the component, every region, as ONE step - and while that step
    // is the latest, "Bring [U1] back" undoes it.
    // ###########################################################################################
    [Fact]
    public void Delete_takes_the_component_off_in_one_step_and_bring_back_undoes_it()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U1", "PAL");

            Assert.Contains("every region", ToolTip.GetTip(Get<Button>(pane, "DeleteComponentButton")) as string, StringComparison.Ordinal);

            Click(Get<Button>(pane, "DeleteComponentButton"));

            Assert.DoesNotContain(session.Content.Board.Components, row => row.BoardLabel == "U1");
            Assert.Equal("[U1] deleted", session.UndoDescription);
            Assert.True(Get<StackPanel>(pane, "GonePanel").IsVisible);
            Assert.False(Get<WrapPanel>(pane, "ActionsPanel").IsVisible);

            Button bringBack = Get<Button>(pane, "BringBackButton");
            Assert.True(bringBack.IsVisible);
            Assert.Equal("Bring [U1] back", bringBack.Content);

            Click(bringBack);

            Assert.Equal(2, session.Content.Board.Components.Count(row => row.BoardLabel == "U1"));
            Assert.True(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);
            Assert.False(Get<StackPanel>(pane, "GonePanel").IsVisible);

            window.Close();
        });
    }

    // Gone some other way - its adding undone - there is no deletion to bring back.
    [Fact]
    public void Bring_back_is_offered_only_while_the_deletion_is_the_latest_step()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.True(pane.DeleteComponent());

            session.Apply("Something else", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U1", "PAL", ComponentField.FriendlyName, "VIC-II") }, origin: new object());

            Assert.True(Get<StackPanel>(pane, "GonePanel").IsVisible);
            Assert.False(Get<Button>(pane, "BringBackButton").IsVisible);

            window.Close();
        });
    }

    // "Show on board" can show only a component marked on a picture - and says why not.
    [Fact]
    public void Show_on_board_is_offered_for_a_component_marked_on_a_picture()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Button show = Get<Button>(pane, "ShowOnBoardButton");
            Assert.False(show.IsEnabled);
            Assert.Contains("not marked on any picture", ToolTip.GetTip(show) as string, StringComparison.Ordinal);

            session.Apply(
                "Marked",
                content => content with
                {
                    Board = new BoardData
                    {
                        Components = content.Board.Components,
                        ComponentPins = content.Board.ComponentPins,
                        ComponentLinks = content.Board.ComponentLinks,
                        ComponentHighlights = [new ComponentHighlightEntry { BoardLabel = "U8", SchematicName = "Top", X = "1", Y = "1", Width = "5", Height = "5" }],
                    },
                },
                origin: new object());

            Assert.True(show.IsEnabled);

            var asked = new List<string>();
            pane.ShowOnBoardRequested += (_, label) => asked.Add(label);

            // What the boxes hold goes in first.
            Get<TextBox>(pane, "FriendlyNameBox").Text = "Processor";
            Click(show);

            Assert.Equal(["U8"], asked);
            Assert.Equal("Processor", U8(session).FriendlyName);

            window.Close();
        });
    }

    // ---- the board's own lists (slice 4a) -------------------------------------------------------

    // The published board with its own lists: files of two categories, out of their groups' order.
    private static BoardData WithBoardParts()
    {
        BoardData board = Published();
        board.BoardLocalFiles.AddRange([
            new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf" },
            new BoardLocalFileEntry { Category = "Troubleshooting", Name = "Dead test", File = "Board local files/dead.pdf" },
            new BoardLocalFileEntry { Category = "Resources", Name = "Notes", File = "Board local files/notes.txt" },
        ]);
        board.BoardLinks.Add(new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = "https://lemon64.com" });
        board.Credits.Add(new CreditEntry { Category = "Schematic image", SubCategory = "PAL", NameOrHandle = "Anna", Contact = "RHD" });
        board.KiCadImportantSignals.Add(new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "PHI2" });
        return board;
    }

    private DraftSession OpenDraftWithBoardParts()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, ComponentPaneTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, WithBoardParts());

        return DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;
    }

    // ###########################################################################################
    // The category "Add link" offers from the row above is no typing (code review, 2026-11-01): pressed and
    // thought better of, the row went into the draft at the next commit as two errors stopping Submit -
    // and a draft made only to look became a real one. One emptied again leaves the draft.
    // ###########################################################################################
    [Fact]
    public void Add_link_pressed_and_left_alone_puts_nothing_in_the_draft()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Links);
            int links = session.Content.Board.BoardLinks.Count;

            Click(Get<Button>(pane, "AddBoardLinkButton"));
            ComponentPaneBoardLinkRow added = pane.BoardLinksForTests[^1];
            Assert.NotEqual(string.Empty, added.Category);

            pane.CommitHeld();

            Assert.Equal(links, session.Content.Board.BoardLinks.Count);
            Assert.False(session.CanUndo);
            Assert.Null(added.Committed);

            added.Name = "Forum 64";
            pane.CommitHeld();
            Assert.Equal(links + 1, session.Content.Board.BoardLinks.Count);

            added.Name = string.Empty;
            pane.CommitHeld();
            Assert.Equal(links, session.Content.Board.BoardLinks.Count);

            // Another category picked IS typing - a row of the draft, with what it lacks marked.
            added.Category = "Videos";
            pane.CommitHeld();
            Assert.Equal(links + 1, session.Content.Board.BoardLinks.Count);

            window.Close();
        });
    }

    private static (ComponentPane Pane, Window Window) ShownPart(DraftSession session, PaneBoardPart part, double width = 1100)
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = width, Height = 1400 };
        window.Show();

        pane.ShowBoardPart(session, part, WithBoardParts());
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    // ###########################################################################################
    // A board part in place of a component: its list in the board's own order (grouped by category
    // until 2026-10-30), counted in its heading - and opening and leaving it changes nothing.
    // ###########################################################################################
    [Fact]
    public void A_board_part_shows_its_list_in_the_boards_order_and_opening_it_changes_nothing()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);

            // The files alone - the links are a part of their own since 2026-10-29 - under ONE
            // headline, the pane's title with the count (owner request, 2026-10-29).
            Assert.Equal("Board: Files (3)", Get<TextBlock>(pane, "TitleText").Text);
            Assert.True(Get<ScrollViewer>(pane, "BoardPartPanel").IsVisible);
            Assert.False(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);
            Assert.False(Get<WrapPanel>(pane, "ActionsPanel").IsVisible);
            Assert.True(Get<StackPanel>(pane, "BoardFilesPart").IsVisible);
            Assert.False(Get<StackPanel>(pane, "BoardLinksPart").IsVisible);
            Assert.False(Get<StackPanel>(pane, "CreditsPart").IsVisible);
            Assert.Equal(["Service manual", "Dead test", "Notes"], pane.BoardFilesForTests.Select(row => row.Name));
            Assert.Equal(["Resources", "Troubleshooting"], pane.BoardFilesForTests[0].CategoryChoices);

            pane.CommitHeld();

            Assert.False(session.CanUndo);
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);

            window.Close();
        });
    }

    // A credit typed in goes into the draft as ONE step, and the part says the draft changes it;
    // removed, it is another step.
    [Fact]
    public void A_credit_goes_into_the_draft_as_one_step_and_the_part_says_it_changed()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Credits);

            Assert.Equal("Board: Credits (1)", Get<TextBlock>(pane, "TitleText").Text);

            // The row added holds nothing yet - no credit of the draft, not counted (code review,
            // 2026-10-30: the title said 2 while the draft and the left column said 1).
            Click(Get<Button>(pane, "AddCreditButton"));
            Assert.Equal("Board: Credits (1)", Get<TextBlock>(pane, "TitleText").Text);
            ComponentPaneCreditRow added = pane.CreditsForTests[^1];
            Assert.Equal(["Schematic image"], added.CategoryChoices);
            added.Category = "Board labelling";
            added.NameOrHandle = " Bo ";
            pane.CommitHeld();

            Assert.Equal("Board: Credits (2)", Get<TextBlock>(pane, "TitleText").Text);
            Assert.Equal("Credits", session.UndoDescription);
            Assert.Equal(["Anna", "Bo"], session.Content.Board.Credits.Select(row => row.NameOrHandle));
            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);

            Click(RowButtons(pane, "CreditsItemsControl").First(button => button.Content as string == "Remove"));

            Assert.Equal(["Bo"], session.Content.Board.Credits.Select(row => row.NameOrHandle));

            window.Close();
        });
    }

    // ###########################################################################################
    // A FILE IS DRAGGED PAST ANY ROW BY ITS HANDLE (owner requests, 2026-10-29: "a possibility to
    // reorder the list elements"; 2026-10-30: "the entire line moving up/down, and not only within its
    // own category", then "instead of having the up/down arrows, then please have a 'North/South' icon
    // and then drag the whole line up/down, just alike in the workbook for photos. It should show a
    // placeholder when dragging"): a handle on every row and no arrows; while it moves the row is a
    // dashed placeholder and the draft untouched; the drop is ONE step written in the new order, and
    // Undo takes it back. "Add me" is gone (owner decision, 2026-10-29) - only "Add credit" adds one.
    // ###########################################################################################
    [Fact]
    public void A_file_dragged_by_its_handle_moves_past_a_file_of_another_category_as_one_step()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);

            // Shown as the board has them: "Service manual" (Resources), "Dead test" (Troubleshooting),
            // "Notes" (Resources).
            ComponentPaneBoardFileRow dead = pane.BoardFilesForTests[1];
            ItemsControl list = Get<ItemsControl>(pane, "BoardFilesItemsControl");

            Assert.Equal(3, Handles(list).Count);
            // The hand says it can be grabbed (owner request, 2026-11-01 - it was the north/south arrows).
            Assert.All(Handles(list), handle => Assert.Equal(Avalonia.Input.StandardCursorType.Hand, CursorTypeOf(handle)));
            Assert.DoesNotContain(
                RowButtons(pane, "BoardFilesItemsControl"),
                button => Avalonia.Automation.AutomationProperties.GetName(button) is "Move up" or "Move down");

            // "Dead test" dragged by its handle up past "Service manual", another category.
            Point from = CentreOf(window, Handles(list)[1]);
            Point top = TopOf(window, list.ContainerFromIndex(0)!) + new Point(0, 2);

            window.MouseDown(from, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            window.MouseMove(top, RawInputModifiers.LeftMouseButton);
            Dispatcher.UIThread.RunJobs();

            Assert.True(dead.IsDropPlaceholder);
            Assert.Same(dead, pane.BoardFilesForTests[0]);
            Assert.Equal(["Service manual", "Dead test", "Notes"], session.Content.Board.BoardLocalFiles.Select(row => row.Name));

            window.MouseUp(top, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.False(dead.IsDropPlaceholder);
            Assert.False(pane.IsDraggingARowForTests);
            Assert.Equal(["Dead test", "Service manual", "Notes"], pane.BoardFilesForTests.Select(row => row.Name));
            Assert.Equal(["Dead test", "Service manual", "Notes"], session.Content.Board.BoardLocalFiles.Select(row => row.Name));
            Assert.Equal("The board's files", session.UndoDescription);

            // ORDER IS A CHANGE (owner decision, 2026-10-30) - the part says so; no row is counted.
            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);
            Assert.Equal(0, BoardDataDiffer.CountChanges(WithBoardParts(), session.Content.Board));
            Assert.Equal([BoardWorkbookSchema.SheetBoardLocalFiles], BoardRowOrder.ReorderedSheets(WithBoardParts(), session.Content.Board));

            // Past the first row: nothing changes.
            Assert.False(pane.MoveBoardFile(dead, -1));
            Assert.Equal(["Dead test", "Service manual", "Notes"], session.Content.Board.BoardLocalFiles.Select(row => row.Name));

            Assert.True(session.Undo());
            Assert.Equal(["Service manual", "Dead test", "Notes"], pane.BoardFilesForTests.Select(row => row.Name));
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);

            Assert.Null(pane.FindControl<Button>("AddMeButton"));

            window.Close();
        });
    }

    // ###########################################################################################
    // FROM THE KEYBOARD: Alt+Up and Alt+Down in a row's box move the row (the arrows' buttons went with
    // the drag, 2026-10-30) - one step, and the focus stays in the same box of the row where it now is.
    // ###########################################################################################
    [Fact]
    public void Alt_up_and_alt_down_in_a_rows_box_move_the_row_and_keep_the_focus_in_it()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);
            ItemsControl list = Get<ItemsControl>(pane, "BoardFilesItemsControl");
            ComponentPaneBoardFileRow dead = pane.BoardFilesForTests[1];

            TextBox name = NameBoxOf(list, 1);
            name.Focus();
            Dispatcher.UIThread.RunJobs();

            window.KeyPress(Key.Up, RawInputModifiers.Alt, PhysicalKey.ArrowUp, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Dead test", "Service manual", "Notes"], session.Content.Board.BoardLocalFiles.Select(row => row.Name));
            Assert.Same(NameBoxOf(list, 0), window.FocusManager?.GetFocusedElement());

            window.KeyPress(Key.Down, RawInputModifiers.Alt, PhysicalKey.ArrowDown, keySymbol: null);
            window.KeyPress(Key.Down, RawInputModifiers.Alt, PhysicalKey.ArrowDown, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Service manual", "Notes", "Dead test"], session.Content.Board.BoardLocalFiles.Select(row => row.Name));
            Assert.Same(dead, pane.BoardFilesForTests[2]);

            window.Close();
        });
    }

    private static List<Border> Handles(ItemsControl list) =>
        list.GetVisualDescendants().OfType<Border>().Where(border => border.Classes.Contains("RowDragHandle")).ToList();

    private static Avalonia.Input.StandardCursorType? CursorTypeOf(Control control) =>
        control.Cursor is null ? null : Enum.Parse<Avalonia.Input.StandardCursorType>(control.Cursor.ToString()!);

    private static TextBox NameBoxOf(ItemsControl list, int index) =>
        list.ContainerFromIndex(index)!.GetVisualDescendants().OfType<TextBox>().First(box => box.FindAncestorOfType<AutoCompleteBox>() is null);

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    private static Point TopOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, 0), window)!.Value;

    // ###########################################################################################
    // A ROW STAYS WHERE IT WAS TYPED, IN THE PANE AND IN THE DRAFT (2026-10-30): the lists are no longer
    // grouped by category, so a row of a category typed last stays last - opened again, the part shows
    // the draft's own order, and the arrows judge the neighbours shown.
    // ###########################################################################################
    [Fact]
    public void A_row_stays_where_it_was_typed_and_opens_again_in_the_drafts_order()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Links);

            // Lemon64 (Forums), then a Videos link, then a Forums link typed in last.
            Click(Get<Button>(pane, "AddBoardLinkButton"));
            pane.BoardLinksForTests[^1].Category = "Videos";
            pane.BoardLinksForTests[^1].Name = "Repair video";
            pane.BoardLinksForTests[^1].Url = "https://example.com/video";
            pane.CommitHeld();

            Click(Get<Button>(pane, "AddBoardLinkButton"));
            pane.BoardLinksForTests[^1].Category = "Forums";
            pane.BoardLinksForTests[^1].Name = "Forum 64";
            pane.BoardLinksForTests[^1].Url = "https://example.com/forum";
            pane.CommitHeld();

            Assert.Equal(["Lemon64", "Repair video", "Forum 64"], pane.BoardLinksForTests.Select(row => row.Name));
            Assert.Equal(["Lemon64", "Repair video", "Forum 64"], session.Content.Board.BoardLinks.Select(row => row.Name));

            // Opened again: as the draft has them.
            pane.ShowBoardPart(session, PaneBoardPart.Links, WithBoardParts());
            Assert.Equal(["Lemon64", "Repair video", "Forum 64"], pane.BoardLinksForTests.Select(row => row.Name));

            window.Close();
        });
    }

    // ###########################################################################################
    // AN IMPORTANT SIGNAL MOVES UP OR DOWN (owner request, 2026-10-30: "I need a possibility to reorder
    // the 'Important signals' entries") - one step, in the order the Schematics tab's panel lists them.
    // ###########################################################################################
    [Fact]
    public void An_important_signal_moves_as_one_step()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);
            pane.UseNetChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false), ("RESET", "RESET", false)]));

            Assert.True(pane.AddSignalForNet("RESET"));
            Assert.Equal(["PHI2", "RESET"], session.Content.Board.KiCadImportantSignals.Select(row => row.KiCadNetName));

            ComponentPaneSignalRow reset = pane.SignalsForTests[1];
            ItemsControl list = Get<ItemsControl>(pane, "SignalsItemsControl");
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(2, Handles(list).Count);

            // Dragged by its handle above PHI2.
            Point from = CentreOf(window, Handles(list)[1]);
            Point top = TopOf(window, list.ContainerFromIndex(0)!) + new Point(0, 2);

            window.MouseDown(from, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            window.MouseMove(top, RawInputModifiers.LeftMouseButton);
            Dispatcher.UIThread.RunJobs();
            Assert.True(reset.IsDropPlaceholder);
            window.MouseUp(top, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["RESET", "PHI2"], session.Content.Board.KiCadImportantSignals.Select(row => row.KiCadNetName));
            Assert.Equal(["RESET", "PHI2"], pane.SignalsForTests.Select(row => row.KiCadNetName));
            Assert.Equal(DraftStepWording.ImportantSignals, session.UndoDescription);

            Assert.True(session.Undo());
            Assert.Equal(["PHI2", "RESET"], session.Content.Board.KiCadImportantSignals.Select(row => row.KiCadNetName));

            window.Close();
        });
    }

    // ###########################################################################################
    // ANOTHER VIEW PICKED WITH NO TICK ON LETS GO OF THE OLD VIEW'S TRACES (code review, 2026-10-30):
    // they stayed, and ticking again showed the old view's copper - or its old problem - over the
    // new view. Ticked again, the traces are read for the view picked.
    // ###########################################################################################
    [Fact]
    public async Task Another_view_picked_with_no_tick_on_reads_the_traces_again_once_ticked()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            var traces = new PreviewTraces([], [new PreviewTrace(new Avalonia.Point(0, 0), new Avalonia.Point(1, 0), 1)]);
            int reads = 0;
            pane.ReadPreviewTraces = (_, _) =>
            {
                reads++;
                (PreviewTraces Traces, string? Problem) answer = reads == 1
                    ? (PreviewTraces.None, PictureToolWording.NotABoardView("Top"))
                    : (traces, null);
                return Task.FromResult(answer);
            };

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            top.ShowOtherSideTraces = true;
            await pane.PreviewsLoadedForTests;
            Assert.Equal(PictureToolWording.NotABoardView("Top"), top.TracesNote);

            top.ShowOtherSideTraces = false;
            pane.CommitCadName(top, "PCB bottom");
            Assert.Equal(1, reads);
            Assert.Null(top.Traces);

            top.ShowOtherSideTraces = true;
            await pane.PreviewsLoadedForTests;

            Assert.Equal(2, reads);
            Assert.Same(traces.OtherSide, top.ShownOtherSideTraces);
            Assert.False(top.HasTracesNote);

            window.Close();
        });
    }

    // A credit is dragged by its handle as a file is - its row a placeholder while it moves.
    [Fact]
    public void A_credit_dragged_by_its_handle_moves_as_one_step()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Credits);

            Click(Get<Button>(pane, "AddCreditButton"));
            pane.CreditsForTests[^1].Category = "Schematic image";
            pane.CreditsForTests[^1].NameOrHandle = "Bo";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            ItemsControl list = Get<ItemsControl>(pane, "CreditsItemsControl");
            Point from = CentreOf(window, Handles(list)[1]);
            Point top = TopOf(window, list.ContainerFromIndex(0)!) + new Point(0, 2);

            window.MouseDown(from, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            window.MouseMove(top, RawInputModifiers.LeftMouseButton);
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.CreditsForTests[0].IsDropPlaceholder);
            window.MouseUp(top, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Bo", "Anna"], session.Content.Board.Credits.Select(row => row.NameOrHandle));
            Assert.Equal("Credits", session.UndoDescription);

            window.Close();
        });
    }

    // The boxes of the board part shown marked as missing, and the "Required" helpers shown under them.
    private static List<Control> MissingBoxes(ComponentPane pane, string list) =>
        Get<ItemsControl>(pane, list).GetVisualDescendants().OfType<Control>()
            .Where(control => control is TextBox or AutoCompleteBox && control.Classes.Contains("Missing") && control.IsEffectivelyVisible)
            .ToList();

    private static int RequiredHelpers(ComponentPane pane, string list) =>
        Get<ItemsControl>(pane, list).GetVisualDescendants().OfType<TextBlock>()
            .Count(block => block.Classes.Contains("Required") && block.IsEffectivelyVisible && block.Text == "Required before submitting");

    // ###########################################################################################
    // *** EVERY OTHER NET IS LISTED LAST (owner request, 2026-11-01: "as the last data in here I would like
    // to see all other nets not already used ... below the 'Suggested' ones, so all nets will be visible
    // in here"). *** A net neither a row nor a suggestion has, with its "Add" - one step, the net becoming
    // a row and leaving the list, the rows still there the very same objects (a row made again under a
    // press threw the click away) - and its "Show on board".
    // ###########################################################################################
    [Fact]
    public void Every_other_net_is_listed_last_and_can_be_added()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);

            pane.UseNetChoices(KiCadNetNames.Choices([("/RESET", "RESET", false), ("N$1", "N$1", false), ("N$2", "N$2", false)]));
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<StackPanel>(pane, "OtherNetsPanel").IsVisible);
            Assert.Equal(ComponentPaneWording.OtherNetsHeading(2), Get<TextBlock>(pane, "OtherNetsHeading").Text);
            Assert.Equal(["N$1", "N$2"], pane.OtherNetsForTests.Select(net => net.KiCadNetName));
            Assert.Contains(pane.SignalSuggestionsForTests, suggestion => suggestion.KiCadNetName == "RESET");

            ComponentPaneOtherNet kept = pane.OtherNetsForTests[1];

            Assert.True(pane.AddSignalForNet("N$1"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["N$2"], pane.OtherNetsForTests.Select(net => net.KiCadNetName));
            Assert.Same(kept, pane.OtherNetsForTests[0]);
            Assert.Contains(session.Content.Board.KiCadImportantSignals, row => row.KiCadNetName == "N$1");
            Assert.Equal(DraftStepWording.ImportantSignals, session.UndoDescription);

            string? shown = null;
            pane.ShowNetRequested += (_, net) => shown = net;
            pane.ShowNetOnBoard(pane.OtherNetsForTests[0].KiCadNetName);
            Assert.Equal("N$2", shown);

            // Without KiCad nets there is nothing to list.
            pane.UseNetChoices(KiCadNetChoices.None);
            Dispatcher.UIThread.RunJobs();
            Assert.False(Get<StackPanel>(pane, "OtherNetsPanel").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A BOARD'S HUNDREDS OF NETS ARE MADE ONLY AS THEY ARE SCROLLED TO *** - hundreds of rows of
    // buttons, realised at once, were what "will it feel sluggish?" asks about.
    // ###########################################################################################
    [Fact]
    public void Hundreds_of_other_nets_are_not_all_realised()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);

            pane.UseNetChoices(KiCadNetNames.Choices(Enumerable.Range(1, 450).Select(number => ((string?)$"N${number}", (string?)$"N${number}", false))));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(450, pane.OtherNetsForTests.Count);

            int realised = Get<ItemsControl>(pane, "OtherNetsItemsControl").GetRealizedContainers().Count();
            Assert.InRange(realised, 1, 100);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** UNDER A PICTURE'S THUMBNAIL: ITS KiCad COMPONENTS AND ITS MARKS (owner request, 2026-11-01: "if
    // it has KiCad data and the image has been mapped to a KiCad file, then per image I want to see status
    // of how many components is available from KiCad and how many has been marked"). *** Only once it
    // shows a view and the counts are known; a card picking another view follows.
    // ###########################################################################################
    [Fact]
    public void A_picture_showing_a_KiCad_view_says_its_components_and_marks()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];

            // Nothing said while it is not known how many a view shows.
            Assert.False(top.ShowsKiCadCounts);

            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["PCB top"] = 120, ["PCB bottom"] = 1 };
            pane.UseKiCadViews(["PCB top", "PCB bottom"], viewComponents: counts);
            Dispatcher.UIThread.RunJobs();

            top.CadName = "PCB top";
            Assert.True(top.ShowsKiCadCounts);
            Assert.Equal("- 120 components available from KiCad data", top.ComponentsAvailableText);
            Assert.Equal(top.MarkedComponents.ToString(System.Globalization.CultureInfo.InvariantCulture), top.MarkedComponentsNumber);

            top.CadName = "PCB bottom";
            Assert.Equal("- 1 component available from KiCad data", top.ComponentsAvailableText);

            // A picture showing no view says nothing.
            top.CadName = string.Empty;
            Assert.False(top.ShowsKiCadCounts);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE MARKS TOLD APART, ONCE ANY WAS MADE BY HAND (owner request, 2026-11-01: "If any components
    // have been manually marked, then I think it should show like this"). *** A mark exactly where KiCad
    // puts it is automatic; anything else is manual. With none by hand, only the total is said. On a
    // schematic sheet there is no automatic line: nothing marks one from KiCad. Worked out from the
    // rectangles, so no board file changes: 1000 x 500 pixels over a 100 x 50 world is 10 pixels a unit.
    // ###########################################################################################
    [Fact]
    public void A_cards_marks_are_told_apart_once_any_was_made_by_hand()
    {
        UiTest.Run(() =>
        {
            var footprints = new Dictionary<string, KiCadViewFootprints>(StringComparer.OrdinalIgnoreCase)
            {
                ["PCB top"] = new KiCadViewFootprints(new Rect(0, 0, 100, 50), new Dictionary<string, Rect>(StringComparer.OrdinalIgnoreCase) { ["U1"] = new Rect(10, 10, 10, 10) }),
            };

            ComponentPaneBoardPictureRow Card(Dictionary<string, List<Rect>> marks, KiCadCalibrationValues? alignment = null)
            {
                ComponentPaneBoardPictureRow card = ComponentPaneBoardPictureRow.From(
                    new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png", CadName = "PCB top" },
                    [],
                    boardHasKiCad: true,
                    aligned: alignment is not null,
                    markedComponents: marks.Count,
                    marksByLabel: marks,
                    alignment: alignment);

                card.ViewComponents = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["PCB top"] = 359 };
                card.ViewFootprints = footprints;
                card.BoardSides = ["PCB top"];
                card.ImageSize = new Size(1000, 500);
                return card;
            }

            var unmoved = new KiCadCalibrationValues(1, 1, 0, 0, false, false);
            var onFootprint = new List<Rect> { new(100, 100, 100, 100) };
            var byHand = new List<Rect> { new(1, 1, 5, 5) };

            // Marked from KiCad alone: only the total.
            ComponentPaneBoardPictureRow automatic = Card(new() { ["U1"] = onFootprint }, unmoved);
            Assert.True(automatic.ShowsKiCadCounts);
            Assert.False(automatic.ShowsManuallyMarked);
            Assert.False(automatic.ShowsAutomaticallyMarked);
            Assert.Equal("1", automatic.MarkedComponentsNumber);
            Assert.Equal(" component marked", automatic.MarkedComponentsWords);

            // One by hand: every line.
            ComponentPaneBoardPictureRow both = Card(new() { ["U1"] = onFootprint, ["U2"] = byHand }, unmoved);
            Assert.Equal("- 359 components available from KiCad data", both.ComponentsAvailableText);
            Assert.True(both.ShowsAutomaticallyMarked);
            Assert.Equal("- 1 component automatically marked from KiCad data", both.ComponentsAutomaticallyMarkedText);
            Assert.True(both.ShowsManuallyMarked);
            Assert.Equal("- 1 component manually marked", both.ComponentsManuallyMarkedText);
            Assert.Equal("2", both.MarkedComponentsNumber);
            Assert.Equal(" components marked", both.MarkedComponentsWords);

            // A mark moved off its footprint is by hand; so is every mark on a picture not aligned.
            Assert.True(Card(new() { ["U1"] = byHand }, unmoved).ShowsManuallyMarked);
            Assert.Equal("- 1 component manually marked", Card(new() { ["U1"] = onFootprint }).ComponentsManuallyMarkedText);

            // A schematic sheet: no line about automatic marks.
            ComponentPaneBoardPictureRow sheet = Card(new() { ["U2"] = byHand }, unmoved);
            sheet.BoardSides = ["PCB bottom"];
            Assert.True(sheet.ShowsManuallyMarked);
            Assert.False(sheet.ShowsAutomaticallyMarked);
        });
    }

    // What a new component needs next is said where it is added (owner request, 2026-11-01: "highlight
    // that an added component must be manually marked in at least one of the schematic images").
    [Fact]
    public void The_new_component_form_says_it_must_be_marked_by_hand()
    {
        DraftSession session = this.OpenDraftWithFiles();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.Equal(ComponentPaneWording.NewComponentMarkNote, Get<TextBlock>(pane, "NewComponentMarkNoteText").Text);
            Assert.Contains("\"Mark components manually\"", ComponentPaneWording.NewComponentMarkNote, StringComparison.Ordinal);
            Assert.Contains("orphan", ComponentPaneWording.NewComponentMarkNote, StringComparison.Ordinal);

            window.Close();
        });
    }

    // ###########################################################################################
    // An other net comes with a proposed name in a box (owner request, 2026-11-01: "a proposed net name
    // in an input field"), changed before it is added; emptied, the proposed name is used.
    // ###########################################################################################
    [Fact]
    public void An_other_net_is_added_under_the_name_in_its_box()
    {
        var net = new ComponentPaneOtherNet("Net-(U1-Pad3)");

        Assert.Equal(ImportantSignalSuggestions.StartingNameOf("Net-(U1-Pad3)"), net.SuggestedName);
        Assert.Equal(net.SuggestedName, net.DisplayName);
        Assert.Equal(net.SuggestedName, net.NameToAdd);

        net.DisplayName = "  Pin 3 of the CPU ";
        Assert.Equal("Pin 3 of the CPU", net.NameToAdd);

        net.DisplayName = "   ";
        Assert.Equal(net.SuggestedName, net.NameToAdd);
    }

    // ###########################################################################################
    // *** AN ADDRESS SUBMIT WOULD REFUSE IS MARKED ON ITS BOX (code review, 2026-11-01). *** "Open link"
    // opens "www.lemon64.com" - LinkTarget adds the https:// - with nothing on the row, while Board status
    // said "1 error to fix" and Submit sent nothing. Framed as an empty one is, the helper saying what
    // to write; "Open link" stays, to look at where it goes. A whole address loses the mark.
    // ###########################################################################################
    [Fact]
    public void A_board_links_address_submit_would_refuse_is_marked_on_its_box()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Links);
            ComponentPaneBoardLinkRow link = pane.BoardLinksForTests[0];

            link.Url = "www.lemon64.com";
            Dispatcher.UIThread.RunJobs();

            Assert.True(link.UrlNotWeb);
            Assert.False(link.UrlMissing);
            Assert.True(link.CanOpen);
            Assert.Single(MissingBoxes(pane, "BoardLinksItemsControl"));
            Assert.Equal(0, RequiredHelpers(pane, "BoardLinksItemsControl"));
            Assert.Equal(1, NotWebHelpers(pane, "BoardLinksItemsControl"));

            link.Url = "https://www.lemon64.com";
            Dispatcher.UIThread.RunJobs();

            Assert.False(link.UrlNotWeb);
            Assert.Empty(MissingBoxes(pane, "BoardLinksItemsControl"));
            Assert.Equal(0, NotWebHelpers(pane, "BoardLinksItemsControl"));

            window.Close();
        });
    }

    // The same on a component's own links, which the check refuses as it does the board's.
    [Fact]
    public void A_component_links_address_submit_would_refuse_is_marked_on_its_box()
    {
        var row = new ComponentPaneLinkRow { Name = "Datasheet", Url = "https://example.org/6510.pdf" };
        Assert.False(row.UrlNotWeb);

        var raised = new List<string?>();
        row.PropertyChanged += (_, e) => raised.Add(e.PropertyName);

        row.Url = "example.org/6510.pdf";

        Assert.True(row.UrlNotWeb);
        Assert.Contains(nameof(ComponentPaneLinkRow.UrlNotWeb), raised);
        Assert.Equal(ComponentPaneWording.LinkNotWeb, row.NotWebText);
    }

    private static int NotWebHelpers(ComponentPane pane, string list) =>
        Get<ItemsControl>(pane, list).GetVisualDescendants().OfType<TextBlock>()
            .Count(block => block.Classes.Contains("Required") && block.IsEffectivelyVisible && block.Text == ComponentPaneWording.LinkNotWeb);

    // ###########################################################################################
    // *** A FIELD A ROW NEEDS, LEFT EMPTY, IS MARKED ON ITS BOX, AND THE ROW GOES INTO THE DRAFT AS IT
    // IS (owner request, 2026-11-01: "highlight the problematic cell with a helper what is missing.
    // Currently this is noted at the bottom of the table ... These errors must be flagged in 'Board
    // status' and then it should not be possible to submit the draft board"). *** A new link's empty
    // name and address are framed red with "Required before submitting" under each - at once; the row
    // goes into the draft with them empty, an error of the draft's (BoardRowRequired), and each box
    // filled in loses its mark. It was held back in the pane until then, a line under it.
    // ###########################################################################################
    [Fact]
    public void A_links_empty_fields_are_marked_on_their_boxes_and_the_link_goes_into_the_draft_as_it_is()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Links);
            Click(Get<Button>(pane, "AddBoardLinkButton"));
            ComponentPaneBoardLinkRow added = pane.BoardLinksForTests[^1];
            Dispatcher.UIThread.RunJobs();

            // Marked at once - the category comes from the row above, so the name and the address.
            Assert.False(added.CategoryMissing);
            Assert.True(added.NameMissing);
            Assert.True(added.UrlMissing);
            Assert.Equal(2, MissingBoxes(pane, "BoardLinksItemsControl").Count);
            Assert.Equal(2, RequiredHelpers(pane, "BoardLinksItemsControl"));

            // In the draft as it is - one step, counted, an error there.
            added.Name = "Forum 64";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Lemon64", "Forum 64"], session.Content.Board.BoardLinks.Select(row => row.Name));
            Assert.Equal(DraftStepWording.BoardLinks, session.UndoDescription);
            Assert.Equal(ComponentPaneWording.PartTitle(PaneBoardPart.Links, 2), Get<TextBlock>(pane, "TitleText").Text);
            Assert.Equal([BoardWorkbookSchema.ColUrl], BoardRowRequired.EmptyColumns(session.Content.Board.BoardLinks[1]));
            Assert.Single(MissingBoxes(pane, "BoardLinksItemsControl"));
            Assert.Equal(1, RequiredHelpers(pane, "BoardLinksItemsControl"));

            added.Url = "https://example.com/forum";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("https://example.com/forum", session.Content.Board.BoardLinks[1].Url);
            Assert.Empty(MissingBoxes(pane, "BoardLinksItemsControl"));
            Assert.Equal(0, RequiredHelpers(pane, "BoardLinksItemsControl"));

            // Its name emptied: into the draft as it is, marked again.
            added.Name = "  ";
            pane.CommitHeld();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Lemon64", ""], session.Content.Board.BoardLinks.Select(row => row.Name.Trim()));
            Assert.True(added.NameMissing);
            Assert.Single(MissingBoxes(pane, "BoardLinksItemsControl"));

            window.Close();
        });
    }

    // A credit needs "What" and "Name or handle" - which part of it and the contact may stay empty; an
    // important signal needs its name; a file every field. Each goes into the draft as it is, marked; a
    // row with nothing typed in it yet is no row.
    [Fact]
    public void A_credit_a_signal_and_a_file_go_into_the_draft_as_they_are_with_their_empty_fields_marked()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Credits);

            Click(Get<Button>(pane, "AddCreditButton"));
            ComponentPaneCreditRow credit = pane.CreditsForTests[^1];
            Assert.True(credit.CategoryMissing);
            Assert.True(credit.NameOrHandleMissing);

            // Nothing typed in it yet: no row of the draft.
            pane.CommitHeld();
            Assert.Equal(["Anna"], session.Content.Board.Credits.Select(row => row.NameOrHandle));
            Assert.Equal(ComponentPaneWording.PartTitle(PaneBoardPart.Credits, 1), Get<TextBlock>(pane, "TitleText").Text);

            credit.NameOrHandle = "Bo";
            pane.CommitHeld();
            Assert.Equal(["Anna", "Bo"], session.Content.Board.Credits.Select(row => row.NameOrHandle));
            Assert.True(credit.CategoryMissing);
            Assert.False(credit.NameOrHandleMissing);

            credit.Category = "Board data";
            pane.CommitHeld();
            Assert.Equal(["Schematic image", "Board data"], session.Content.Board.Credits.Select(row => row.Category));
            Assert.False(credit.CategoryMissing);

            // An important signal's name emptied.
            pane.ShowBoardPart(session, PaneBoardPart.ImportantSignals, WithBoardParts());
            pane.UseNetChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false)]));
            ComponentPaneSignalRow signal = pane.SignalsForTests[0];

            signal.DisplayName = string.Empty;
            pane.CommitHeld();
            Assert.Equal([string.Empty], session.Content.Board.KiCadImportantSignals.Select(row => row.DisplayName));
            Assert.True(signal.DisplayNameMissing);
            Assert.False(signal.NetMissing);

            // A file's category emptied.
            pane.ShowBoardPart(session, PaneBoardPart.Files, WithBoardParts());
            pane.BoardFilesForTests[0].Category = " ";
            pane.CommitHeld();
            Assert.Equal(string.Empty, session.Content.Board.BoardLocalFiles[0].Category.Trim());
            Assert.True(pane.BoardFilesForTests[0].CategoryMissing);

            // Every one of them an error of the draft's, the table's and the server's own rule.
            Assert.Equal(
                new BoardRowsIncomplete(Files: 1, Links: 0, Credits: 0, Signals: 1, Errors: 2),
                BoardRowRequired.Of(session.Content.Board));

            window.Close();
        });
    }

    // ###########################################################################################
    // "Open" on a file of the board (owner request, 2026-10-29: "a possibility to open the file so I
    // can view it"): the file found as every file of the draft is, handed to the launcher; one not on
    // this machine is said, and nothing is opened.
    // ###########################################################################################
    [Fact]
    public void Open_on_a_board_file_opens_the_file_found_and_says_when_it_is_not_there()
    {
        BoardData board = WithBoardParts();
        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [this.Stored("Board local files/manual.pdf", [37, 80, 68, 70])]));
        DraftSession session = DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);
            var opened = new List<string>();
            pane.OpenFileOverrideForTests = (path, root) => opened.Add(Path.GetFileName(path));

            // "Open file" (owner request, 2026-10-30 - "Open" until then).
            List<Button> open = RowButtons(pane, "BoardFilesItemsControl").Where(button => button.Content as string == "Open file").ToList();
            Assert.Equal(3, open.Count);

            // "Service manual" - its bytes are in the draft's store.
            Click(open[0]);
            Assert.Single(opened);
            Assert.False(Get<TextBlock>(pane, "PaneMessageText").IsVisible);

            // "Notes", last as the board has it - cited, but on no disk of this machine.
            Click(open[2]);
            Assert.Single(opened);
            Assert.Equal(ComponentPaneWording.FileNotFound("Notes"), Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Open" on a link (owner request, 2026-10-29: "open a link in a browser"): the address by
    // LinkTarget's rule - one typed without a scheme gets https:// - and the button off while the
    // address is nothing a browser can be sent to.
    // ###########################################################################################
    [Fact]
    public void Open_on_a_link_sends_its_address_to_the_browser()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Links);
            var opened = new List<string>();
            pane.OpenLinkOverrideForTests = opened.Add;

            ComponentPaneBoardLinkRow link = pane.BoardLinksForTests[0];
            Assert.True(link.CanOpen);

            // "Open link" (owner request, 2026-10-30 - "Open" until then).
            Dispatcher.UIThread.RunJobs();
            Assert.Single(RowButtons(pane, "BoardLinksItemsControl"), button => button.Content as string == "Open link");

            pane.OpenBoardLink(link);
            Assert.Equal(["https://lemon64.com"], opened);

            link.Url = "www.zimmers.net";
            pane.OpenBoardLink(link);
            Assert.Equal("https://www.zimmers.net", opened[^1]);

            link.Url = "see the forum";
            Assert.False(link.CanOpen);
            pane.OpenBoardLink(link);
            Assert.Equal(2, opened.Count);
            Assert.Equal(ComponentPaneWording.LinkNotOpened("Lemon64"), Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // A file of the board is copied in under the board's files folder, joins the last category listed,
    // and the draft is saved at once - as a component's file is.
    [Fact]
    public async Task A_board_file_added_is_copied_in_and_saved_at_once()
    {
        DraftSession session = this.OpenDraftWithBoardParts();
        string picked = this.Pickable("Voltages.png", ComponentPaneTests.Png);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Files);
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);

            await pane.AddBoardFileAsync();

            // The last row listed is "Notes" (Resources) since the lists keep the board's order
            // (2026-10-30) - it was "Dead test" (Troubleshooting), last of the grouped list.
            BoardLocalFileEntry added = session.Content.Board.BoardLocalFiles.Single(row => row.Name == "Voltages");
            Assert.Equal("Board local files/Voltages.png", added.File);
            Assert.Equal("Resources", added.Category);
            Assert.Equal("The board's files", session.UndoDescription);
            Assert.Contains("Board local files/Voltages.png", this.ListedOnDisk());
            Assert.False(session.HasUnsavedChanges);

            window.Close();
        });
    }

    // ###########################################################################################
    // A NEW SIGNAL PICKS ITS NET (owner request, 2026-10-29: the KiCad net is the KiCad data's, shown
    // as text, never typed): the picker offers the board's nets no row has, "Add signal" is on only
    // for one of them, and the row comes named as it would be suggested - ONE step. A row naming a net
    // the KiCad data lacks (from the table, say) still says the Schematics tab will not light it up.
    // Without KiCad data there is nothing to pick, and the part says why.
    // ###########################################################################################
    [Fact]
    public void A_new_signal_picks_one_of_the_boards_nets_and_a_net_the_KiCad_data_lacks_is_said()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);
            AutoCompleteBox picker = Get<AutoCompleteBox>(pane, "NewSignalNetBox");
            Button add = Get<Button>(pane, "AddSignalButton");

            Assert.True(Get<TextBlock>(pane, "KiCadNetsLineText").IsVisible);
            Assert.Equal("Reading the board's KiCad nets...", Get<TextBlock>(pane, "KiCadNetsLineText").Text);
            Assert.False(pane.SignalsForTests[0].HasProblem);
            Assert.False(add.IsEnabled);

            pane.UseNetChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false), ("/RESET", "RESET", false)]));

            Assert.False(Get<TextBlock>(pane, "KiCadNetsLineText").IsVisible);
            Assert.False(pane.SignalsForTests[0].HasProblem);

            // PHI2 is the "Clock" row's already.
            Assert.True(picker.IsEnabled);
            Assert.Equal(["RESET"], picker.ItemsSource!.Cast<string>());
            Assert.False(add.IsEnabled);

            picker.Text = "NRESET";
            Assert.False(add.IsEnabled);
            Assert.False(pane.AddSignalForNet("NRESET"));
            Assert.False(pane.AddSignalForNet("PHI2"));
            Assert.False(session.CanUndo);

            picker.Text = "/RESET";
            Assert.True(add.IsEnabled);
            Click(add);

            ComponentPaneSignalRow added = pane.SignalsForTests[^1];
            Assert.Equal(("RESET", "RESET"), (added.DisplayName, added.KiCadNetName));
            Assert.Equal("Important KiCad signals", session.UndoDescription);
            Assert.Contains(session.Content.Board.KiCadImportantSignals, row => row.DisplayName == "RESET" && row.KiCadNetName == "RESET");
            Assert.Equal(string.Empty, picker.Text);
            Assert.Empty(picker.ItemsSource!.Cast<string>());
            Assert.Equal("KiCad: Important signals (2)", Get<TextBlock>(pane, "TitleText").Text);

            // A row naming a net the KiCad data does not have says so.
            pane.SignalsForTests[0].KiCadNetName = "NRESET";
            pane.CommitHeld();
            Assert.True(pane.SignalsForTests[0].HasProblem);
            Assert.Contains("[NRESET]", pane.SignalsForTests[0].Problem, StringComparison.Ordinal);

            pane.UseNetChoices(KiCadNetChoices.None);
            Assert.Equal(ComponentPaneWording.NoKiCadNets, Get<TextBlock>(pane, "KiCadNetsLineText").Text);
            Assert.False(picker.IsEnabled);
            Assert.False(add.IsEnabled);

            window.Close();
        });
    }

    // ###########################################################################################
    // THE NEW SIGNAL'S PICKER IS LEFT ALONE WHILE A NET IS PICKED (code review, 2026-10-30): typing in
    // it started the pane's typing pause, whose commit of the signals replaced the picker's list - the
    // drop-down cleared and refilled under the net being picked. Its list is replaced only when the
    // nets it offers change.
    // ###########################################################################################
    [Fact]
    public void The_new_signals_picker_keeps_its_list_while_a_net_is_typed()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);
            AutoCompleteBox picker = Get<AutoCompleteBox>(pane, "NewSignalNetBox");
            pane.UseNetChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false), ("/RESET", "RESET", false), ("/IRQ", "IRQ", false)]));
            Dispatcher.UIThread.RunJobs();

            System.Collections.IEnumerable? offered = picker.ItemsSource;
            Assert.Equal(["IRQ", "RESET"], offered!.Cast<string>().Order(StringComparer.Ordinal));

            // Typed into the picker's own box: no pause, so nothing is committed under it. (A row's box
            // left first - its first layout may have started one.)
            TextBox rowBox = Get<ItemsControl>(pane, "SignalsItemsControl").GetVisualDescendants().OfType<TextBox>().First();
            TextBox typedIn = picker.GetVisualDescendants().OfType<TextBox>().First();
            Assert.True(rowBox.Focus());
            Assert.True(typedIn.Focus());
            Dispatcher.UIThread.RunJobs();
            Assert.False(pane.IsTypingPauseRunningForTests);

            // A TextBox says its text changed on the dispatcher, after the change.
            typedIn.Text = "RE";
            Dispatcher.UIThread.RunJobs();
            Assert.False(pane.IsTypingPauseRunningForTests);

            // A row's box typed in still starts one.
            rowBox.Text += " ";
            Dispatcher.UIThread.RunJobs();
            Assert.True(pane.IsTypingPauseRunningForTests);

            // A commit that leaves the nets to pick from as they were keeps the very list.
            pane.SignalsForTests[0].DisplayName = "Clock signal";
            pane.CommitHeld();
            Assert.Equal("Clock signal", session.Content.Board.KiCadImportantSignals[0].DisplayName);
            Assert.Same(offered, picker.ItemsSource);

            // A net a row takes is no longer offered - then the list is replaced.
            Assert.True(pane.AddSignalForNet("/IRQ"));
            Assert.NotSame(offered, picker.ItemsSource);
            Assert.Equal(["RESET"], picker.ItemsSource!.Cast<string>());

            window.Close();
        });
    }

    // ###########################################################################################
    // Suggestions and click-to-light-up (the plan: "pick nets from a searchable KiCad list, with
    // suggestions and click-to-light-up"): once the nets are read, the ones whose names say what they
    // are - but none a row has - are suggested; "Add" takes one, "Add all suggestions" the rest as ONE
    // step; "Show on board" asks the host to light a net up.
    // ###########################################################################################
    [Fact]
    public void The_KiCad_nets_suggest_signals_and_Show_on_board_asks_for_the_net()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);
            StackPanel suggestions = Get<StackPanel>(pane, "SignalSuggestionsPanel");

            // Nothing while the nets are being read.
            Assert.False(suggestions.IsVisible);

            pane.UseNetChoices(KiCadNetNames.Choices([
                ("PHI2", "PHI2", false),
                ("/CPU/~{RESET}", "~{RESET}", false),
                ("+5V", "+5V", false),
                ("GND", "GND", false),
                ("MA0", "MA0", false),
            ]));

            // PHI2 is the "Clock" row's already.
            Assert.True(suggestions.IsVisible);
            Assert.Equal(
                ["+5VDC:+5V", "Ground:GND", "RESET:~{RESET}"],
                pane.SignalSuggestionsForTests.Select(suggestion => $"{suggestion.DisplayName}:{suggestion.KiCadNetName}"));
            Assert.Equal("Suggested from the KiCad nets (3)", Get<TextBlock>(pane, "SignalSuggestionsHeading").Text);

            // The name is a box (owner request, 2026-10-30) - changed before it is added.
            pane.SignalSuggestionsForTests[2].DisplayName = " System reset ";
            Assert.Equal("RESET", pane.SignalSuggestionsForTests[2].SuggestedName);

            Assert.True(pane.AddSuggestions([pane.SignalSuggestionsForTests[2]]));
            Assert.Equal("Important KiCad signals", session.UndoDescription);
            Assert.Contains(session.Content.Board.KiCadImportantSignals, row => row.DisplayName == "System reset" && row.KiCadNetName == "~{RESET}");
            Assert.Equal(2, pane.SignalSuggestionsForTests.Count);

            // A box emptied adds the suggested name; a name typed survives the boxes being committed.
            pane.SignalSuggestionsForTests[0].DisplayName = "  ";
            pane.SignalSuggestionsForTests[1].DisplayName = "0V";
            pane.CommitHeld();
            Assert.Equal("0V", pane.SignalSuggestionsForTests[1].DisplayName);

            Click(Get<Button>(pane, "AddAllSuggestionsButton"));
            Assert.Empty(pane.SignalSuggestionsForTests);
            Assert.False(suggestions.IsVisible);
            Assert.Equal(4, session.Content.Board.KiCadImportantSignals.Count);
            Assert.Contains(session.Content.Board.KiCadImportantSignals, row => row.DisplayName == "+5VDC" && row.KiCadNetName == "+5V");
            Assert.Contains(session.Content.Board.KiCadImportantSignals, row => row.DisplayName == "0V" && row.KiCadNetName == "GND");

            // All of them were one step.
            Assert.True(session.Undo());
            Assert.Equal(2, session.Content.Board.KiCadImportantSignals.Count);

            var asked = new List<string>();
            pane.ShowNetRequested += (_, net) => asked.Add(net);

            Assert.True(pane.SignalsForTests[0].CanShowOnBoard);
            pane.ShowNetOnBoard(" PHI2 ");
            pane.ShowNetOnBoard(" ");
            Assert.Equal(["PHI2"], asked);

            window.Close();
        });
    }

    // A signal with no net, or one the KiCad data does not have, has nothing to light up.
    [Fact]
    public void A_signal_without_a_net_the_KiCad_data_has_cannot_be_shown_on_the_board()
    {
        var row = new ComponentPaneSignalRow { DisplayName = "Reset" };
        row.UseChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false)]));
        Assert.False(row.CanShowOnBoard);

        row.KiCadNetName = "NRESET";
        row.UseChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false)]));
        Assert.False(row.CanShowOnBoard);

        // A board without KiCad data has no net to light up at all (code review, 2026-10-26: the
        // button switched to the Schematics tab and lit nothing).
        row.KiCadNetName = "PHI2";
        row.UseChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false)]));
        Assert.True(row.CanShowOnBoard);
        row.UseChoices(KiCadNetChoices.None);
        Assert.False(row.CanShowOnBoard);
    }

    // ###########################################################################################
    // The suggestions are kept in place while they stay the same (code review, 2026-10-26): every box
    // left commits the signals, and a list cleared and refilled then threw away the button a press had
    // just gone down on - the click was lost (the trap CLAUDE.md records for the drafts list).
    // ###########################################################################################
    [Fact]
    public void A_suggestion_still_suggested_keeps_its_row_when_the_signals_change()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);
            pane.UseNetChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false), ("+5V", "+5V", false), ("GND", "GND", false)]));

            ComponentPaneSignalSuggestion first = pane.SignalSuggestionsForTests[0];

            pane.SignalsForTests[0].DisplayName = "Clock; PHI2";
            pane.CommitHeld();

            Assert.Equal("Important KiCad signals", session.UndoDescription);
            Assert.Same(first, pane.SignalSuggestionsForTests[0]);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Add all suggestions" adds what is suggested once what the rows hold is in (code review,
    // 2026-10-26): it took the list first, and a net just given to a row went in twice. Since
    // 2026-10-29 a row's net is picked, not typed - the net picked first is suggested no more.
    // ###########################################################################################
    [Fact]
    public void Add_all_suggestions_does_not_add_a_net_just_picked_for_a_row_again()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);
            pane.UseNetChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false), ("+5V", "+5V", false), ("GND", "GND", false)]));

            Assert.True(pane.AddSignalForNet("+5V"));
            pane.SignalsForTests[^1].DisplayName = "5 volts";

            Click(Get<Button>(pane, "AddAllSuggestionsButton"));

            Assert.Single(session.Content.Board.KiCadImportantSignals, row => row.KiCadNetName == "+5V");
            Assert.Single(session.Content.Board.KiCadImportantSignals, row => row.KiCadNetName == "GND");

            window.Close();
        });
    }

    // Another editor's change, or an undo, shows in the part at once - its rows read again.
    [Fact]
    public void An_undo_reads_the_board_part_again()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Links);

            Assert.Equal("Board: Links (1)", Get<TextBlock>(pane, "TitleText").Text);
            Assert.True(Get<StackPanel>(pane, "BoardLinksPart").IsVisible);
            Assert.False(Get<StackPanel>(pane, "BoardFilesPart").IsVisible);

            pane.BoardLinksForTests[0].Name = "Lemon64 forum";
            pane.CommitHeld();
            Assert.Equal("The board's links", session.UndoDescription);

            Assert.True(session.Undo());

            Assert.Equal("Lemon64", pane.BoardLinksForTests[0].Name);
            Assert.False(Get<Border>(pane, "ChangedMark").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // How the board's baselines were measured is its "Scope measurements" files (owner decision,
    // 2026-10-26): linked above the component's baselines, a click opening the one found - or, with
    // none, a line saying where to add them.
    // ###########################################################################################
    [Fact]
    public void The_boards_scope_measurement_files_are_linked_above_the_baselines()
    {
        BoardData board = Published();
        board.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Scope measurements", Name = "Introduction", File = "Scope baseline/intro.txt" });

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [this.Stored("Scope baseline/intro.txt", [72, 105])]));
        DraftSession session = DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            var opened = new List<string>();
            pane.OpenFileOverrideForTests = (path, root) => opened.Add(Path.GetFileName(path));

            Assert.Equal("How this board's baselines were measured:", Get<TextBlock>(pane, "MeasuringNotesIntro").Text);
            Assert.True(Get<ItemsControl>(pane, "MeasuringNotesItemsControl").IsVisible);

            pane.ShowPage(ComponentPage.Baselines);
            Dispatcher.UIThread.RunJobs();

            Click(RowButtons(pane, "MeasuringNotesItemsControl").Single());

            Assert.Single(opened);

            window.Close();
        });
    }

    // A note of one region is linked only there (code review, 2026-11-02) - as the "Resources" tab lists it.
    [Fact]
    public void A_scope_measurement_file_of_the_other_region_is_not_linked()
    {
        BoardData board = Published();
        board.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Scope measurements", Name = "NTSC's", File = "Scope baseline/ntsc.txt", Region = "NTSC" });

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [this.Stored("Scope baseline/ntsc.txt", [72, 105])]));
        DraftSession session = DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.False(Get<ItemsControl>(pane, "MeasuringNotesItemsControl").IsVisible);

            // Written down, for the other region - never "not written down yet", which asks for a second
            // copy of notes the board has (code review, 2026-11-02).
            string line = Get<TextBlock>(pane, "MeasuringNotesIntro").Text ?? string.Empty;
            Assert.DoesNotContain("not written down", line, StringComparison.Ordinal);
            Assert.Contains("NTSC", line, StringComparison.Ordinal);

            pane.UseRegion("NTSC");
            Dispatcher.UIThread.RunJobs();

            Assert.True(Get<ItemsControl>(pane, "MeasuringNotesItemsControl").IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void Without_scope_measurement_files_the_pane_says_where_to_add_them()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            Assert.False(Get<ItemsControl>(pane, "MeasuringNotesItemsControl").IsVisible);
            Assert.Contains("\"Scope measurements\"", Get<TextBlock>(pane, "MeasuringNotesIntro").Text, StringComparison.Ordinal);

            window.Close();
        });
    }

    // ---- copy from another board (slice 5) -------------------------------------------------------

    // The PLA of another board, with its package, pins and a replacement part.
    private static ComponentCopySource OtherPla() =>
        ComponentCopy.Sources(
            [(
                "Commodore 64 - 250425",
                new BoardData
                {
                    Components = [new ComponentEntry { BoardLabel = "U7", FriendlyName = "PLA", PartNumber = "906114-01", Category = "IC", Package = "DIP-28" }],
                    ComponentPins = [new ComponentPinEntry { BoardLabel = "U7", Pin = "28", Name = "VCC", Direction = "Power" }],
                    ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U7", PartNumber = "PLAnkton", Kind = "Drop-in" }],
                }
            )],
            "906114-01",
            "PAL")[0];

    // ###########################################################################################
    // The same part on another board, found by the part number: the panel offers what it has, ticked,
    // and "Copy" puts it into this component as ONE step named after the board.
    // ###########################################################################################
    [Fact]
    public async Task Copy_from_another_board_takes_what_is_ticked_in_one_step()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            var asked = new List<string>();
            pane.FindOnOtherBoards = (partNumber, region) =>
            {
                asked.Add($"{partNumber}|{region}");
                return Task.FromResult<IReadOnlyList<ComponentCopySource>>([OtherPla()]);
            };

            Get<TextBox>(pane, "PartNumberBox").Text = "906114-01";
            await pane.OpenCopyPanelAsync();

            Assert.Equal(["906114-01|PAL"], asked);
            Assert.True(pane.IsCopyPanelOpen);
            List<CheckBox> parts = Get<StackPanel>(pane, "CopyPartsPanel").Children.OfType<CheckBox>().ToList();
            Assert.Equal(["Package (DIP-28)", "Pins (1)", "Replacement parts (1)"], parts.Select(box => box.Content as string));

            parts[2].IsChecked = false;
            Assert.True(pane.CopyPicked());

            Assert.Equal("[U8] copied from Commodore 64 - 250425", session.UndoDescription);
            Assert.Equal("DIP-28", U8(session).Package);
            Assert.Equal(["28:VCC"], ComponentPaneEdits.PinsOf(session.Content.Board, "U8", "PAL").Select(pin => $"{pin.Pin}:{pin.Name}"));
            Assert.Empty(ComponentPaneEdits.ReplacementsOf(session.Content.Board, "U8", "PAL"));
            Assert.False(pane.IsCopyPanelOpen);
            Assert.Equal("DIP-28", Get<AutoCompleteBox>(pane, "PackageBox").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Also on other boards": offered once the draft changes this component and it has a part number;
    // the panel lists the same part's boards, every one ticked, and what the draft changed - and hands
    // what is ticked to the host, saying what came of it.
    // ###########################################################################################
    [Fact]
    public async Task Also_on_other_boards_hands_the_ticked_boards_and_changes_to_the_host()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            Button also = Get<Button>(pane, "AlsoOnOtherBoardsButton");
            Assert.False(also.IsVisible);

            var handed = new List<(string Label, List<string> Boards, List<ComponentCopyPart> Parts)>();
            var searched = new List<string>();
            pane.FindAlsoTargets = (partNumber, region, files) =>
            {
                searched.Add($"{partNumber}|{region}|{files.Count}");
                return Task.FromResult(new AlsoTargets([OtherPla(), OtherPla() with { BoardName = "Commodore 64 - 250469" }], []));
            };
            pane.ApplyToOtherBoards = (here, targets, parts, files) =>
            {
                Assert.Empty(files);
                handed.Add((here!.Row.BoardLabel, targets.Select(target => target.BoardName).ToList(), parts.ToList()));
                return Task.FromResult("Changed on 1 board, in its own draft: Commodore 64 - 250425.");
            };

            Get<TextBox>(pane, "PartNumberBox").Text = "906114-01";
            Get<TextBox>(pane, "DescriptionBox").Text = "Socket it";
            pane.CommitHeld();
            Assert.True(also.IsVisible);

            await pane.OpenAlsoPanelAsync();

            // One search of the other boards, for the part and its files at once.
            Assert.Equal(["906114-01|PAL|0"], searched);
            Assert.True(pane.IsAlsoPanelOpen);
            List<CheckBox> boards = Get<StackPanel>(pane, "AlsoBoardsPanel").Children.OfType<CheckBox>().ToList();
            Assert.Equal(2, boards.Count);
            Assert.All(boards, box => Assert.True(box.IsChecked));
            Assert.Equal(["Short description"], Get<StackPanel>(pane, "AlsoPartsPanel").Children.OfType<CheckBox>().Select(box => box.Content as string));

            boards[1].IsChecked = false;
            await pane.ChangeOtherBoardsAsync();

            (string label, List<string> handedBoards, List<ComponentCopyPart> handedParts) = Assert.Single(handed);
            Assert.Equal("U8", label);
            Assert.Equal(["Commodore 64 - 250425"], handedBoards);
            Assert.Equal([ComponentCopyPart.Description], handedParts);
            Assert.False(pane.IsAlsoPanelOpen);
            Assert.Equal("Changed on 1 board, in its own draft: Commodore 64 - 250425.", Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // Without a part number there is nothing to match by; with one nobody else has, it says so.
    [Fact]
    public async Task Copy_from_another_board_says_why_when_there_is_nothing_to_offer()
    {
        DraftSession session = this.OpenDraft();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            pane.FindOnOtherBoards = (_, _) => Task.FromResult<IReadOnlyList<ComponentCopySource>>([]);

            await pane.OpenCopyPanelAsync();
            Assert.Equal(ComponentPaneWording.CopyNeedsPartNumber, Get<TextBlock>(pane, "PaneMessageText").Text);

            Get<TextBox>(pane, "PartNumberBox").Text = "XYZ";
            await pane.OpenCopyPanelAsync();
            Assert.Equal(ComponentPaneWording.CopyFoundNothing("XYZ"), Get<TextBlock>(pane, "PaneMessageText").Text);
            Assert.False(pane.IsCopyPanelOpen);

            window.Close();
        });
    }

    // ---- also on other boards, for files ------------------------------------------------------

    // A published board's own board file, listing `files` - what the draft's files are compared with.
    private string PublishedBoardFile(BoardData board, params (string Name, byte[] Bytes)[] files)
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Published", Guid.NewGuid().ToString("N"), DataTreeLayout.BoardFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        BoardFiles.Write(path, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            files.Select(file => new BoardDocument.StoredFileRow
            {
                Name = file.Name,
                Sha256 = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(file.Bytes)),
                Size = file.Bytes.Length,
            }).ToList()));

        return path;
    }

    // ###########################################################################################
    // "Also on other boards" for a file (the plan: "This datasheet is also used by 11 other boards -
    // change it there too?"): a datasheet replaced is offered - with no part number at all - to the
    // boards the host finds with the same file, each ticked, and handed over as files alone.
    // ###########################################################################################
    [Fact]
    public async Task A_datasheet_replaced_is_offered_to_the_boards_with_the_same_file()
    {
        DraftSession session = this.OpenDraftWithFiles();
        BoardData published = PublishedWithFiles();
        string publishedFile = this.PublishedBoardFile(
            published,
            ("Component images/U8.png", ComponentPaneTests.Png),
            ("Scope baseline/U8_1.png", ComponentPaneTests.Png),
            ("Component local files/6510.pdf", [37, 80, 68, 70]));
        string picked = this.Pickable("6510 rev B.pdf", [37, 80, 68, 70, 66]);

        await UiTest.RunAsync(async () =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();
            pane.Show(session, "U8", "PAL", published, publishedFile);
            Dispatcher.UIThread.RunJobs();

            Button also = Get<Button>(pane, "AlsoOnOtherBoardsButton");
            Assert.False(also.IsVisible);

            var asked = new List<ChangedFile>();
            pane.FindAlsoTargets = (partNumber, _, files) =>
            {
                Assert.Equal(string.Empty, partNumber);
                asked.AddRange(files);
                return Task.FromResult(new AlsoTargets(
                    [],
                    [new ChangedFileElsewhere(files[0], "Commodore 64 - 250425", "Commodore/C64/250425/board.json", "Component local files/6510.pdf")]));
            };

            var handed = new List<ChangedFileElsewhere>();
            pane.ApplyToOtherBoards = (here, targets, parts, files) =>
            {
                Assert.Empty(targets);
                Assert.Empty(parts);
                handed.AddRange(files);
                return Task.FromResult("Changed on 1 board, in its own draft: Commodore 64 - 250425.");
            };

            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);
            await pane.ChangeFileAsync(pane.FilesForTests[0]);

            // No part number: the file replaced is enough to look for.
            Assert.True(also.IsVisible);

            await pane.OpenAlsoPanelAsync();

            ChangedFile datasheet = Assert.Single(asked);
            Assert.Equal(("Datasheet", "Component local files/6510.pdf"), (datasheet.RowName, datasheet.PublishedName));
            Assert.False(Get<StackPanel>(pane, "AlsoPartsSection").IsVisible);
            Assert.True(Get<StackPanel>(pane, "AlsoFilesSection").IsVisible);

            StackPanel filesPanel = Get<StackPanel>(pane, "AlsoFilesPanel");
            Assert.Equal("Datasheet (6510.pdf)", Assert.Single(filesPanel.Children.OfType<TextBlock>()).Text);
            Assert.True(Assert.Single(filesPanel.Children.OfType<CheckBox>()).IsChecked);

            await pane.ChangeOtherBoardsAsync();

            Assert.Equal("Commodore 64 - 250425", Assert.Single(handed).BoardName);
            Assert.False(pane.IsAlsoPanelOpen);
            Assert.Equal("Changed on 1 board, in its own draft: Commodore 64 - 250425.", Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // Nothing found for a file replaced on a component WITH a part number, whose other parts are as
    // published: no search by part number ran, so none is said to have found nothing (code review,
    // 2026-10-26: "No other board ... has a component with the part number [6510]" - never looked for).
    // ###########################################################################################
    [Fact]
    public async Task Nothing_found_for_a_file_alone_says_nothing_of_the_part_number()
    {
        DraftSession session = this.OpenDraftWithFiles();
        BoardData published = PublishedWithFiles();
        published.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510", Category = "IC", Description = "The processor", PartNumber = "6510" };
        string publishedFile = this.PublishedBoardFile(published, ("Component local files/6510.pdf", [37, 80, 68, 70]));
        string picked = this.Pickable("6510 rev B.pdf", [37, 80, 68, 70, 66]);

        session.Apply("Part number", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "6510") });

        await UiTest.RunAsync(async () =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();
            pane.Show(session, "U8", "PAL", published, publishedFile);
            Dispatcher.UIThread.RunJobs();

            var partsAsked = new List<string>();
            pane.FindAlsoTargets = (partNumber, _, _) =>
            {
                partsAsked.Add(partNumber);
                return Task.FromResult(AlsoTargets.None);
            };
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);

            await pane.ChangeFileAsync(pane.FilesForTests[0]);
            await pane.OpenAlsoPanelAsync();

            // Nothing of the part changed: it is not looked for.
            Assert.Equal([string.Empty], partsAsked);
            Assert.Equal("No other board on this machine has the files you replaced.", Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // A part changed AND a file replaced, with no other board having the part but some having the
    // file: the panel offers the file - and says the part was not found (code review, 2026-10-26: the
    // part's section was hidden without a word, and the notes were expected to go across too).
    // ###########################################################################################
    [Fact]
    public async Task A_part_no_other_board_has_is_said_while_its_files_are_offered()
    {
        DraftSession session = this.OpenDraftWithFiles();
        BoardData published = PublishedWithFiles();
        published.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510", Category = "IC", Description = "The processor", PartNumber = "6510" };
        string publishedFile = this.PublishedBoardFile(published, ("Component local files/6510.pdf", [37, 80, 68, 70]));
        string picked = this.Pickable("6510 rev B.pdf", [37, 80, 68, 70, 66]);

        session.Apply("Part number", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "6510") });

        await UiTest.RunAsync(async () =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();
            pane.Show(session, "U8", "PAL", published, publishedFile);
            Dispatcher.UIThread.RunJobs();

            pane.FindAlsoTargets = (_, _, files) => Task.FromResult(new AlsoTargets(
                [],
                [new ChangedFileElsewhere(files[0], "Commodore 64 - 250425", "Commodore/C64/250425/board.json", "Component local files/6510.pdf")]));
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);

            Get<TextBox>(pane, "DescriptionBox").Text = "Runs hot";
            pane.CommitHeld();
            await pane.ChangeFileAsync(pane.FilesForTests[0]);
            await pane.OpenAlsoPanelAsync();

            Assert.True(pane.IsAlsoPanelOpen);
            Assert.False(Get<StackPanel>(pane, "AlsoPartsSection").IsVisible);
            Assert.Equal(ComponentPaneWording.CopyFoundNothing("6510"), Get<TextBlock>(pane, "PaneMessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // The board's own files under Files and links: nothing to offer until one is replaced, then the
    // button, the boards with the same file ticked, and the files alone handed to the host.
    // ###########################################################################################
    [Fact]
    public async Task A_board_file_replaced_is_offered_to_the_boards_with_the_same_file()
    {
        BoardData board = WithBoardParts();
        byte[] manual = [37, 80, 68, 70, 77];

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [this.Stored("Board local files/manual.pdf", manual)]));

        DraftSession session = DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;
        string publishedFile = this.PublishedBoardFile(WithBoardParts(), ("Board local files/manual.pdf", manual));
        string picked = this.Pickable("manual 2nd edition.pdf", [37, 80, 68, 70, 50]);

        await UiTest.RunAsync(async () =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();
            pane.ShowBoardPart(session, PaneBoardPart.Files, WithBoardParts(), publishedFile);
            Dispatcher.UIThread.RunJobs();

            Button also = Get<Button>(pane, "BoardFilesAlsoButton");
            Assert.False(also.IsVisible);

            await pane.OpenBoardAlsoPanelAsync();
            Assert.Equal(ComponentPaneWording.AlsoBoardFilesNothingChanged, Get<TextBlock>(pane, "PaneMessageText").Text);

            pane.FindAlsoTargets = (_, _, files) => Task.FromResult(new AlsoTargets(
                [],
                [new ChangedFileElsewhere(files[0], "Commodore 64 - 250425", "Commodore/C64/250425/board.json", "Board local files/C64 manual.pdf")]));

            var handed = new List<(ComponentCopySource? Here, ChangedFileElsewhere File)>();
            pane.ApplyToOtherBoards = (here, _, _, files) =>
            {
                handed.AddRange(files.Select(file => (here, file)));
                return Task.FromResult("Changed on 1 board, in its own draft: Commodore 64 - 250425.");
            };

            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);
            await pane.ChangeBoardFileAsync(pane.BoardFilesForTests.Single(row => row.Name == "Service manual"));

            Assert.True(also.IsVisible);

            await pane.OpenBoardAlsoPanelAsync();

            Assert.True(pane.IsBoardAlsoPanelOpen);
            StackPanel filesPanel = Get<StackPanel>(pane, "BoardAlsoFilesPanel");
            Assert.Equal("Service manual (manual.pdf)", Assert.Single(filesPanel.Children.OfType<TextBlock>()).Text);
            Assert.Equal("Commodore 64 - 250425 - as C64 manual.pdf", Assert.Single(filesPanel.Children.OfType<CheckBox>()).Content);

            await pane.ChangeBoardFilesOnOtherBoardsAsync();

            (ComponentCopySource? here, ChangedFileElsewhere file) = Assert.Single(handed);
            Assert.Null(here);
            Assert.Equal("Board local files/C64 manual.pdf", file.NameThere);
            Assert.False(pane.IsBoardAlsoPanelOpen);

            window.Close();
        });
    }

    // ---- the pins from the board's KiCad data ---------------------------------------------------

    // U8's pins as the board's KiCad data has them - pin 1 as the draft has it, two more.
    private static readonly KiCadPin[] U8KiCadPins =
    [
        new("1", "GND", "Ground", "GND"),
        new("2", "RDY", "Input", "RDY"),
        new("40", "VCC", "Power", "+5V"),
    ];

    // ###########################################################################################
    // "Fill from KiCad" (the plan: "On a KiCad board, Fill from KiCad adds each pin's net and its
    // name from the symbol"): offered once the host hands the pins over, each pin showing its net on
    // this board; the pins missing added as ONE step - pin 1's name as typed is kept - and a second
    // fill finds nothing to do.
    // ###########################################################################################
    [Fact]
    public void Fill_from_KiCad_adds_the_pins_missing_as_one_step_and_shows_each_pins_net()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            Button fill = Get<Button>(pane, "FillFromKiCadButton");

            Assert.False(fill.IsVisible);
            Assert.False(pane.FillFromKiCad());

            // Pins of another component are not taken.
            pane.UseKiCadPins(session, "U1", U8KiCadPins);
            Assert.False(fill.IsVisible);

            pane.UseKiCadPins(session, "u8", U8KiCadPins);
            Assert.True(fill.IsVisible);
            Assert.True(Get<TextBlock>(pane, "PinNetHeading").IsVisible);
            Assert.Equal("GND", Assert.Single(pane.PinsForTests).Net);

            Assert.True(pane.FillFromKiCad());

            Assert.Equal("Pins of [U8] from KiCad", session.UndoDescription);
            Assert.Equal(
                ["1:VSS:Ground", "2:RDY:Input", "40:VCC:Power"],
                ComponentPaneEdits.PinsOf(session.Content.Board, "U8", "PAL").Select(pin => $"{pin.Pin}:{pin.Name}:{pin.Direction}"));
            Assert.Equal(["GND", "RDY", "+5V"], pane.PinsForTests.Select(row => row.Net));
            Assert.Equal("Filled from KiCad: 2 pins added.", Get<TextBlock>(pane, "PaneMessageText").Text);

            Assert.False(pane.FillFromKiCad());
            Assert.Equal(ComponentPaneWording.NothingToFillFromKiCad, Get<TextBlock>(pane, "PaneMessageText").Text);

            // A pin typed by hand shows its net once it is in.
            pane.PinsForTests[1].Pin = "40";
            pane.PinsForTests[2].Pin = "2";
            pane.CommitHeld();
            Assert.Equal(["GND", "+5V", "RDY"], pane.PinsForTests.Select(row => row.Net));

            window.Close();
        });
    }

    // Another component shown lets go of the KiCad pins - they were the last one's.
    [Fact]
    public void Another_component_shown_lets_go_of_the_KiCad_pins()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);
            var shown = new List<string>();
            pane.ComponentShown += (_, label) => shown.Add(label);

            pane.UseKiCadPins(session, "U8", U8KiCadPins);
            pane.Show(session, "U1", "PAL", Published());

            Assert.Equal(["U1"], shown);
            Assert.Empty(pane.KiCadPinsForTests);
            Assert.False(Get<Button>(pane, "FillFromKiCadButton").IsVisible);

            window.Close();
        });
    }

    // ---- the pictures and their colours (slice 4b) ----------------------------------------------

    // Two pictures - one with its colours set, one blank (CRT's defaults) - and a component marked on
    // the first; the first picture's bytes stored, so its preview can be read. `withKiCad`: the board
    // lists a KiCad file and the first picture names its KiCad view (slice 6's "Align KiCad").
    // `alignBottom`: KiCad is aligned on the second picture, which has no marks. `alignTop`: on the
    // first - with `withKiCad`, the picture whose other side's traces can be drawn on its preview.
    // `topImage`: the first picture's bytes, the 1 x 1 PNG unless given.
    private DraftSession OpenDraftWithPictures(bool withKiCad = false, bool alignBottom = false, bool alignTop = false, byte[]? topImage = null)
    {
        BoardData board = Published();
        board.Schematics.AddRange([
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = withKiCad ? "PCB top" : "", SchematicHighlightColor = "Red", SchematicHighlightOpacity = "30%", ThumbnailHighlightColor = "Red", ThumbnailHighlightOpacity = "100%" },
            new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "Images/bottom.png" },
        ]);
        board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "0", Y = "0", Width = "1", Height = "1" });

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(
                board,
                alignBottom ? [new KiCadCalibrationEntry { SchematicName = "Bottom", CadName = "PCB bottom", OffsetX = 1, OffsetY = 2, ScaleX = 1, ScaleY = 1 }]
                : alignTop ? [new KiCadCalibrationEntry { SchematicName = "Top", CadName = "PCB top", OffsetX = 0, OffsetY = 0, ScaleX = 1, ScaleY = 1 }]
                : [],
                []),
            name => name,
            withKiCad
                ? [this.Stored("Images/top.png", topImage ?? ComponentPaneTests.Png), this.Stored("KiCad data/board.kicad_pcb", "(kicad_pcb)"u8.ToArray())]
                : [this.Stored("Images/top.png", topImage ?? ComponentPaneTests.Png)]));

        return DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;
    }

    // ###########################################################################################
    // The previews on the cards (owner request, 2026-10-29): as wide as dragged at any card's handle,
    // the same on every card - told to the host, which keeps it per board - and never dragged past
    // the range.
    // ###########################################################################################
    [Fact]
    public void The_previews_width_is_the_same_on_every_card_and_a_drag_is_told()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            var chosen = new List<double>();
            pane.PreviewWidthChosen += (_, width) => chosen.Add(width);

            Assert.All(pane.BoardPicturesForTests, row => Assert.Equal(ContributeWidths.DefaultPreviewWidth, row.PreviewWidth));

            pane.DragPreviewWidthForTests(+160);

            Assert.Equal([480.0], chosen);
            Assert.All(pane.BoardPicturesForTests, row => Assert.Equal(480, row.PreviewWidth));
            Assert.Equal(300, pane.BoardPicturesForTests[0].PreviewHeight);

            pane.DragPreviewWidthForTests(-5000);
            Assert.Equal(ContributeWidths.PreviewRange.Min, chosen[^1]);

            // Cards made again (an image moved) are as wide.
            pane.MovePicture(pane.BoardPicturesForTests[0], +1);
            Assert.All(pane.BoardPicturesForTests, row => Assert.Equal(ContributeWidths.PreviewRange.Min, row.PreviewWidth));

            window.Close();
        });
    }

    // ###########################################################################################
    // The traces on the preview (owner requests, 2026-10-29): this side and the other side, each its
    // own tick, both drawn when both are ticked - read ONCE for both, through the host, which reads
    // the board as "Mark all automatically" does - and each drawn only while ticked. Cards made again
    // keep the ticks and read them again, once.
    // ###########################################################################################
    [Fact]
    public async Task Each_sides_traces_are_drawn_while_ticked_and_both_are_read_once()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            var asked = new List<string>();
            var traces = new PreviewTraces(
                [new PreviewTrace(new Avalonia.Point(0, 0), new Avalonia.Point(1, 0), 1)],
                [new PreviewTrace(new Avalonia.Point(0, 0), new Avalonia.Point(0, 1), 1)]);
            pane.ReadPreviewTraces = (picture, _) =>
            {
                asked.Add(picture);
                return Task.FromResult((traces, (string?)null));
            };

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            Assert.True(top.CanShowTraces);
            Assert.Null(top.ShownThisSideTraces);
            Assert.Null(top.ShownOtherSideTraces);

            top.ShowOtherSideTraces = true;
            await pane.PreviewsLoadedForTests;
            Assert.Equal(["Top"], asked);
            Assert.Null(top.ShownThisSideTraces);
            Assert.Same(traces.OtherSide, top.ShownOtherSideTraces);

            // Both ticked: both drawn, nothing read again.
            top.ShowThisSideTraces = true;
            await pane.PreviewsLoadedForTests;
            Assert.Equal(["Top"], asked);
            Assert.Same(traces.ThisSide, top.ShownThisSideTraces);
            Assert.Same(traces.OtherSide, top.ShownOtherSideTraces);
            Assert.False(top.HasTracesNote);

            pane.MovePicture(top, +1);
            await pane.PreviewsLoadedForTests;

            ComponentPaneBoardPictureRow again = pane.BoardPicturesForTests.Single(row => row.Name == "Top");
            Assert.NotSame(top, again);
            Assert.True(again.ShowThisSideTraces && again.ShowOtherSideTraces);
            Assert.Equal(["Top", "Top"], asked);
            Assert.Same(traces.ThisSide, again.ShownThisSideTraces);

            again.ShowOtherSideTraces = false;
            Assert.Null(again.ShownOtherSideTraces);
            Assert.Same(traces.ThisSide, again.ShownThisSideTraces);

            window.Close();
        });
    }

    // ###########################################################################################
    // EACH SIDE'S TRACES HAVE A COLOUR AND AN OPACITY OF THEIR OWN (owner request, 2026-10-30): picked
    // on the card, they go into the draft as the pictures' colours step - this side's colour and both
    // opacities in their own columns - and the previews draw them (TraceColours, the Schematics tab's
    // rule). A picker writing back what it shows writes nothing: the derived values of blank columns
    // are no change.
    // ###########################################################################################
    [Fact]
    public void Each_sides_trace_colour_and_opacity_are_picked_on_the_card_and_go_into_the_draft()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow bottom = pane.BoardPicturesForTests[1];

            // Blank: as before - the highlight's colour, both sides a quarter above its opacity.
            Assert.Equal(Avalonia.Media.Colors.DeepSkyBlue, bottom.ThisSideTraceColor);
            Assert.Equal(45, bottom.ThisSideTraceOpacity, 3);
            Assert.Equal(45, bottom.OppositeTraceOpacity, 3);

            // Written back as shown: no change, no step.
            bottom.ThisSideTraceColor = bottom.ThisSideTraceColor;
            bottom.ThisSideTraceOpacity = bottom.ThisSideTraceOpacity;
            bottom.OppositeTraceOpacity = bottom.OppositeTraceOpacity;
            pane.CommitHeld();
            Assert.False(session.CanUndo);

            bottom.ThisSideTraceColor = Avalonia.Media.Colors.Lime;
            bottom.ThisSideTraceOpacity = 70;
            bottom.OppositeTraceOpacity = 15;
            pane.CommitHeld();

            BoardSchematicEntry row = session.Content.Board.Schematics[1];
            Assert.Equal("Lime", row.TraceHighlightColor);
            Assert.Equal("70%", row.TraceHighlightOpacity);
            Assert.Equal("15%", row.OppositeTraceHighlightOpacity);
            Assert.Equal(string.Empty, row.SchematicHighlightColor);

            Assert.Equal(0.70, bottom.ThisSideTraceFraction, 3);
            Assert.Equal(0.15, bottom.OppositeTraceFraction, 3);
            Assert.Equal("70%", bottom.ThisSideTraceOpacityText);
            Assert.Equal("15%", bottom.OppositeTraceOpacityText);

            // Its own colour now: the highlight's no longer moves it.
            bottom.HighlightColor = Avalonia.Media.Colors.Gold;
            Assert.Equal(Avalonia.Media.Colors.Lime, bottom.ThisSideTraceColor);

            window.Close();
        });
    }

    // This side's traces are drawn in the Schematic image colour while their own is not picked, as the
    // Schematics tab draws a lit net's own copper - its blank DeepSkyBlue, the copper's, not the
    // rectangles' IndianRed.
    [Fact]
    public void This_sides_traces_take_the_selected_components_colour()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            Assert.Equal(Avalonia.Media.Colors.Red, pane.BoardPicturesForTests[0].ThisSideTraceColor);
            Assert.Equal(Avalonia.Media.Colors.DeepSkyBlue, pane.BoardPicturesForTests[1].ThisSideTraceColor);

            pane.BoardPicturesForTests[1].HighlightColor = Avalonia.Media.Colors.Gold;
            Assert.Equal(Avalonia.Media.Colors.Gold, pane.BoardPicturesForTests[1].ThisSideTraceColor);

            window.Close();
        });
    }

    // ###########################################################################################
    // An image with no KiCad data has its KiCad options off (owner request, 2026-10-29: "if the
    // schematic image has no KiCad data, then the KiCad (traces) options should be disabled"): the
    // trace colour and the ticks, the reason in their tooltips - on a board with no KiCad files, and
    // for an image with no KiCad view picked. With a view but not aligned, the colour is on and the
    // ticks are not.
    // ###########################################################################################
    [Fact]
    public void An_image_with_no_KiCad_data_has_its_KiCad_options_off()
    {
        DraftSession withoutKiCad = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(withoutKiCad, PaneBoardPart.Pictures);

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            Assert.False(top.HasKiCadData);
            Assert.False(top.CanShowTraces);
            Assert.Equal(ComponentPaneWording.TraceColourNeedsKiCadTip, top.TraceColourTip);
            Assert.Equal(ComponentPaneWording.ShowTracesNeedsAlignTip, top.ThisSideTracesTip);

            window.Close();
        });

        DraftSession notAligned = this.OpenDraftWithPictures(withKiCad: true);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(notAligned, PaneBoardPart.Pictures);

            // "Top" names its KiCad view; "Bottom" names none.
            Assert.True(pane.BoardPicturesForTests[0].HasKiCadData);
            Assert.Equal(ComponentPaneWording.TraceColourTip, pane.BoardPicturesForTests[0].TraceColourTip);
            Assert.False(pane.BoardPicturesForTests[0].CanShowTraces);
            Assert.False(pane.BoardPicturesForTests[1].HasKiCadData);

            window.Close();
        });
    }

    // Why an image's traces cannot be drawn is said under the ticks, in the contributor's words; a
    // side ticked with no copper says that.
    [Fact]
    public async Task Why_the_traces_cannot_be_drawn_is_said_under_the_ticks()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            pane.ReadPreviewTraces = (_, _) => Task.FromResult((PreviewTraces.None, (string?)PictureToolWording.NotABoardView("Top")));

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            top.ShowOtherSideTraces = true;
            await pane.PreviewsLoadedForTests;

            Assert.Equal(PictureToolWording.NotABoardView("Top"), top.TracesNote);
            Assert.True(top.HasTracesNote);
            Assert.Empty(top.ShownOtherSideTraces!);

            top.ShowOtherSideTraces = false;
            Assert.False(top.HasTracesNote);

            window.Close();
        });
    }

    // ###########################################################################################
    // A READ OF THE TRACES THAT FAILS SAYS SO (code review, 2026-10-30): the read's exception ended the
    // load with "Reading the board's KiCad traces..." under the ticks for good, nothing in the log.
    // ###########################################################################################
    [Fact]
    public async Task A_read_of_the_traces_that_fails_says_so_instead_of_reading_for_ever()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;
            pane.ReadPreviewTraces = (_, _) => Task.FromException<(PreviewTraces Traces, string? Problem)>(new IOException("The KiCad file is held"));

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            top.ShowOtherSideTraces = true;
            await pane.PreviewsLoadedForTests;

            Assert.Equal(PictureToolWording.KiCadNotRead, top.TracesNote);
            Assert.False(top.TracesReading);

            window.Close();
        });
    }

    // A side ticked with no copper says so - and says nothing once only the side with copper is ticked.
    [Fact]
    public async Task A_side_ticked_with_no_traces_says_so_under_the_ticks()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            pane.ReadPreviewTraces = (_, _) => Task.FromResult((
                new PreviewTraces([new PreviewTrace(new Avalonia.Point(0, 0), new Avalonia.Point(1, 1), 1)], []),
                (string?)null));

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            top.ShowOtherSideTraces = true;
            await pane.PreviewsLoadedForTests;

            Assert.Equal(ComponentPaneWording.TracesNote(false, 0, true, 0), top.TracesNote);

            top.ShowThisSideTraces = true;
            top.ShowOtherSideTraces = false;
            Assert.False(top.HasTracesNote);
            Assert.Single(top.ShownThisSideTraces!);

            window.Close();
        });
    }

    // ###########################################################################################
    // A READ OF THE TRACES ANSWERS ONLY WHILE IT IS THE LATEST ASKED (code review, 2026-10-30): another
    // KiCad view picked while one ran read nothing (the running read was kept, of the old view), and a
    // rename while one ran left "Reading..." under the ticks for good. Each now reads again, and an
    // older read ending last never wins.
    // ###########################################################################################
    [Fact]
    public async Task A_view_picked_or_a_rename_while_the_traces_are_read_reads_them_again_and_the_newer_read_wins()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            var reads = new List<(string Picture, TaskCompletionSource<(PreviewTraces Traces, string? Problem)> Answer)>();
            pane.ReadPreviewTraces = (picture, _) =>
            {
                var answer = new TaskCompletionSource<(PreviewTraces Traces, string? Problem)>();
                reads.Add((picture, answer));
                return answer.Task;
            };

            PreviewTraces Traces(double y) => new([], [new PreviewTrace(new Avalonia.Point(0, y), new Avalonia.Point(1, y), 1)]);

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            top.ShowOtherSideTraces = true;
            Assert.Single(reads);
            Assert.Equal(ComponentPaneWording.ReadingTraces, top.TracesNote);

            // Another view picked while the first read runs: read again. The first ends LAST, and the
            // newer read's traces stay.
            pane.CommitCadName(top, "PCB bottom");
            Assert.Equal(2, reads.Count);

            PreviewTraces newer = Traces(2);
            reads[1].Answer.SetResult((newer, null));
            reads[0].Answer.SetResult((Traces(1), null));
            await pane.PreviewsLoadedForTests;

            Assert.Same(newer, top.Traces);
            Assert.False(top.HasTracesNote);

            // Renamed while a read runs: read again under the new name - never "Reading..." for good.
            pane.CommitCadName(top, "PCB top");
            Assert.Equal(3, reads.Count);

            top.NameDraft = "Top side";
            Assert.True(pane.CommitPictureName(top));
            Assert.Equal(4, reads.Count);
            Assert.Equal("Top side", reads[3].Picture);

            reads[2].Answer.SetResult((PreviewTraces.None, PictureToolWording.PictureNotFound("Top")));
            PreviewTraces renamed = Traces(3);
            reads[3].Answer.SetResult((renamed, null));
            await pane.PreviewsLoadedForTests;

            Assert.Same(renamed, top.Traces);
            Assert.NotEqual(ComponentPaneWording.ReadingTraces, top.TracesNote);
            Assert.False(top.HasTracesNote);

            window.Close();
        });
    }

    // ###########################################################################################
    // TICKS AN IMAGE CAN NO LONGER SHOW ARE LET GO OF (code review, 2026-10-30): a card made again put
    // its remembered ticks back even when its KiCad view had gone (an undo, the table) - ticked, greyed
    // out, and no way to untick them.
    // ###########################################################################################
    [Fact]
    public async Task Ticks_an_image_can_no_longer_show_are_let_go_of_when_its_card_is_made_again()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);
        var top = new SchematicRowRef("Top", "Images/top.png");

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;
            pane.ReadPreviewTraces = (_, _) => Task.FromResult((PreviewTraces.None, (string?)null));

            pane.BoardPicturesForTests[0].ShowThisSideTraces = true;
            await pane.PreviewsLoadedForTests;

            // Another editor takes the image's KiCad view away: the cards are made again.
            Assert.True(session.Apply("Elsewhere", content => content with { Board = SchematicRows.WithCadName(content.Board, top, "").Board }, origin: new object()));
            Dispatcher.UIThread.RunJobs();

            ComponentPaneBoardPictureRow again = pane.BoardPicturesForTests[0];
            Assert.False(again.CanShowTraces);
            Assert.False(again.ShowThisSideTraces);

            // And the tick is gone, not waiting: the view back, the card shows nothing ticked.
            Assert.True(session.Apply("Elsewhere", content => content with { Board = SchematicRows.WithCadName(content.Board, top, "PCB top").Board }, origin: new object()));
            Dispatcher.UIThread.RunJobs();

            Assert.True(pane.BoardPicturesForTests[0].CanShowTraces);
            Assert.False(pane.BoardPicturesForTests[0].ShowThisSideTraces);

            window.Close();
        });
    }

    // ###########################################################################################
    // TWO IMAGES OF ONE NAME KEEP THEIR OWN TICKS (code review, 2026-10-30): the ticks were kept by name
    // alone, so cards made again ticked both - and read both as the first of the name. They are kept
    // by name AND file, as a card is told apart, and the read is asked with the card's file.
    // ###########################################################################################
    [Fact]
    public async Task Two_images_of_one_name_keep_their_own_ticks_and_each_reads_its_own_traces()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true, alignTop: true);
        Assert.True(session.Apply("Same name", content => content with
        {
            Board = content.Board.WithSchematics(content.Board.Schematics
                .Select(row => row.SchematicName == "Bottom"
                    ? new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = row.SchematicImageFile, CadName = "PCB bottom" }
                    : row)
                .ToList()),
        }));

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            var asked = new List<(string Picture, string File)>();
            pane.ReadPreviewTraces = (picture, file) =>
            {
                asked.Add((picture, file));
                return Task.FromResult((PreviewTraces.None, (string?)null));
            };

            ComponentPaneBoardPictureRow first = pane.BoardPicturesForTests[0];
            Assert.True(pane.BoardPicturesForTests[1].CanShowTraces);

            first.ShowOtherSideTraces = true;
            await pane.PreviewsLoadedForTests;
            Assert.Equal([("Top", "Images/top.png")], asked);

            // The cards made again: only that image ticked, read again as itself.
            pane.MovePicture(first, +1);
            await pane.PreviewsLoadedForTests;

            Assert.True(pane.BoardPicturesForTests.Single(row => row.File == "Images/top.png").ShowOtherSideTraces);
            Assert.False(pane.BoardPicturesForTests.Single(row => row.File == "Images/bottom.png").ShowsAnyTraces);
            Assert.Equal([("Top", "Images/top.png"), ("Top", "Images/top.png")], asked);

            window.Close();
        });
    }

    // ###########################################################################################
    // PREVIEWS DRAGGED WIDER ARE READ WIDER (code review, 2026-10-30): read at 640 whatever their width,
    // a preview dragged to 900 on a 150 % screen was drawn soft. Read for the width in steps
    // (PicturePreviewGeometry.WholeWidth), again only once a drag ends past one; read sharper for the
    // zoom, a card decodes one image at a time, the widest asked meanwhile next - one never shown is
    // never handed to the card; and reads done are let go of.
    // ###########################################################################################
    [Fact]
    public async Task Previews_dragged_wider_are_read_again_once_the_drag_ends_past_a_step()
    {
        DraftSession session = this.OpenDraftWithPictures(topImage: WidePng());

        await UiTest.RunAsync(async () =>
        {
            // Wide enough for previews of 920 beside the settings (ContributeWidths.PreviewShown).
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures, width: 1600);
            await pane.PreviewsLoadedForTests;

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            Avalonia.Media.Imaging.Bitmap read = top.Thumbnail!;
            Assert.Equal(640, top.ThumbnailDecodedFor);

            // Wider, inside the step: nothing read again.
            pane.DragPreviewWidthForTests(+200);
            await pane.PreviewsLoadedForTests;
            Assert.Same(read, top.Thumbnail);

            // Past it: read again for the wider step, once the drag ended - and the narrower read of
            // the picture let go of, one read held per picture (code review, 2026-10-30).
            int held = pane.PicturesHeldForTests;
            pane.DragPreviewWidthForTests(+400);
            Assert.Equal(1024, top.ThumbnailAskedFor);
            await pane.PreviewsLoadedForTests;
            Assert.Equal(1024, top.ThumbnailDecodedFor);
            Assert.NotSame(read, top.Thumbnail);
            Assert.Equal(held, pane.PicturesHeldForTests);

            // Narrower again: what was read stays.
            pane.DragPreviewWidthForTests(-600);
            await pane.PreviewsLoadedForTests;
            Assert.Equal(1024, top.ThumbnailDecodedFor);

            // Every read done is let go of.
            Assert.True(pane.PreviewLoadsHeldForTests <= 2, $"{pane.PreviewLoadsHeldForTests} reads held");

            window.Close();
        });
    }

    // ###########################################################################################
    // A PICTURE IS NEVER READ WIDER THAN IT IS (code review, 2026-10-30): a 600 px scan was decoded
    // to 1536 for previews dragged wide - and decoded again at every step, for nothing.
    // ###########################################################################################
    [Fact]
    public async Task A_picture_is_never_read_wider_than_it_is()
    {
        DraftSession session = this.OpenDraftWithPictures(topImage: PngOfSize(600, 300));

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            Assert.Equal(600, top.Thumbnail!.PixelSize.Width);

            var decoded = new List<int>();
            pane.PreviewDecodeForTests = width =>
            {
                lock (decoded)
                {
                    decoded.Add(width);
                }

                return true;
            };

            // Dragged past two steps: nothing read again - it is shown at its own width already.
            pane.DragPreviewWidthForTests(+600);
            await pane.PreviewsLoadedForTests;

            Assert.Empty(decoded);
            Assert.Equal(600, top.Thumbnail!.PixelSize.Width);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A WIDER READ HELD SERVES NARROWER PREVIEWS (code review, 2026-11-01). *** One read is held per
    // picture, and it stayed the wide one once the previews were dragged narrow - but the look-up asked
    // only for the narrow width, so every card made again (an undo, an add, a move) decoded its picture
    // from disk, and declined to keep what it read.
    // ###########################################################################################
    [Fact]
    public async Task A_wider_read_held_serves_narrower_previews_with_no_second_read()
    {
        DraftSession session = this.OpenDraftWithPictures(topImage: WidePng());

        await UiTest.RunAsync(async () =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1600, Height = 1400 };
            window.Show();
            pane.ShowBoardPart(session, PaneBoardPart.Pictures, WithBoardParts());
            Dispatcher.UIThread.RunJobs();
            await pane.PreviewsLoadedForTests;

            pane.DragPreviewWidthForTests(+600);
            await pane.PreviewsLoadedForTests;
            int wide = pane.BoardPicturesForTests[0].ThumbnailDecodedFor;
            Assert.True(wide > 640, $"The previews were read for {wide}");

            pane.DragPreviewWidthForTests(-600);
            await pane.PreviewsLoadedForTests;

            int decodes = 0;
            pane.PreviewDecodeForTests = _ =>
            {
                Interlocked.Increment(ref decodes);
                return true;
            };

            // The cards made again by a step of the draft from elsewhere (an undo in the header, say),
            // the previews narrow: the wide read held is what they show.
            ComponentPaneBoardPictureRow before = pane.BoardPicturesForTests[0];
            session.Apply("From elsewhere", content => content with
            {
                Board = BoardPartEdits.WithCredits(content.Board, [.. content.Board.Credits, new CreditEntry { Category = "Thanks", NameOrHandle = "Anna" }]),
            });
            Dispatcher.UIThread.RunJobs();
            await pane.PreviewsLoadedForTests;

            Assert.NotSame(before, pane.BoardPicturesForTests[0]);
            Assert.Equal(0, decodes);
            Assert.True(pane.BoardPicturesForTests[0].Thumbnail!.PixelSize.Width >= wide);

            window.Close();
        });
    }

    // ###########################################################################################
    // A CARD BEING READ IS NOT READ AGAIN, AND A WIDER READ THAT FAILS BLANKS NOTHING (code review,
    // 2026-10-30): a card still loading was decoded twice when the previews' width was set again; and
    // a wider read that failed blanked a preview that was showing, which a narrower read could then
    // not bring back. It keeps what it shows, and the wider read may be asked again.
    // ###########################################################################################
    [Fact]
    public async Task A_card_being_read_is_not_read_again_and_a_wider_read_that_fails_blanks_nothing()
    {
        DraftSession session = this.OpenDraftWithPictures(topImage: WidePng());

        await UiTest.RunAsync(async () =>
        {
            using var hold = new ManualResetEventSlim(false);
            int decodes = 0;
            bool failWider = false;

            var pane = new ComponentPane { AutoSave = false };
            pane.PreviewDecodeForTests = width =>
            {
                Interlocked.Increment(ref decodes);
                hold.Wait(TimeSpan.FromSeconds(10));
                return !(failWider && width > 640);
            };

            // Wide enough for previews of 900 beside the settings (ContributeWidths.PreviewShown).
            var window = new Window { Content = pane, Width = 1600, Height = 1400 };
            window.Show();
            pane.ShowBoardPart(session, PaneBoardPart.Pictures, WithBoardParts());
            Dispatcher.UIThread.RunJobs();

            // The previews' width set again while the card is still being read: no second read.
            pane.DragPreviewWidthForTests(+1);
            hold.Set();
            await pane.PreviewsLoadedForTests;
            Assert.Equal(1, decodes);

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            Avalonia.Media.Imaging.Bitmap shown = top.Thumbnail!;

            // Wider, and the read fails: the card keeps its preview.
            failWider = true;
            pane.DragPreviewWidthForTests(+600);
            await pane.PreviewsLoadedForTests;
            Assert.Same(shown, top.Thumbnail);
            Assert.Equal(640, top.ThumbnailDecodedFor);

            // Asked again once it can be read: it is.
            failWider = false;
            pane.DragPreviewWidthForTests(+1);
            await pane.PreviewsLoadedForTests;
            Assert.Equal(1024, top.ThumbnailDecodedFor);
            Assert.NotSame(shown, top.Thumbnail);

            window.Close();
        });
    }

    // ###########################################################################################
    // ONE READ HELD PER PICTURE, WHICHEVER ENDS FIRST (code review, 2026-10-30): a narrower read that
    // ended after a wider one was kept beside it - both images held until the pane closed. The wider
    // read stays, the narrower one is let go of, and the card shows the wider one.
    // ###########################################################################################
    [Fact]
    public async Task A_narrower_read_ending_after_a_wider_one_is_not_kept_beside_it()
    {
        DraftSession session = this.OpenDraftWithPictures(topImage: WidePng());

        await UiTest.RunAsync(async () =>
        {
            using var holdNarrow = new ManualResetEventSlim(false);

            var pane = new ComponentPane { AutoSave = false };
            pane.PreviewDecodeForTests = width =>
            {
                if (width <= 640)
                {
                    holdNarrow.Wait(TimeSpan.FromSeconds(10));
                }

                return true;
            };

            var window = new Window { Content = pane, Width = 1600, Height = 1400 };
            window.Show();
            pane.ShowBoardPart(session, PaneBoardPart.Pictures, WithBoardParts());
            Dispatcher.UIThread.RunJobs();

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];

            // Dragged wider while the 640 read waits: the 1024 read ends first.
            pane.DragPreviewWidthForTests(+600);

            for (int i = 0; i < 400 && top.ThumbnailDecodedFor != 1024; i++)
            {
                await Task.Delay(5);
                Dispatcher.UIThread.RunJobs();
            }

            Assert.Equal(1024, top.ThumbnailDecodedFor);
            int held = pane.PicturesHeldForTests;

            holdNarrow.Set();
            await pane.PreviewsLoadedForTests;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(1024, top.ThumbnailDecodedFor);
            Assert.Equal(held, pane.PicturesHeldForTests);

            window.Close();
        });
    }

    // A picture whose header says 4000 x 2000 - the bytes after it are never decoded by the headless
    // platform, which makes an image of the width asked.
    private static byte[] WidePng() => PngOfSize(4000, 2000);

    private static byte[] PngOfSize(int width, int height)
    {
        byte[] png = (byte[])ComponentPaneTests.Png.Clone();
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(png.AsSpan(16), width);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(png.AsSpan(20), height);
        return png;
    }

    [Fact]
    public async Task A_zoomed_card_decodes_one_sharper_image_at_a_time_the_widest_asked_meanwhile_next()
    {
        DraftSession session = this.OpenDraftWithPictures(topImage: WidePng());

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            Assert.Equal(4000, top.ImageSize.Width);

            var shown = new List<int>();
            top.PropertyChanged += (_, change) =>
            {
                if (change.PropertyName == nameof(ComponentPaneBoardPictureRow.SharpThumbnail) && top.SharpThumbnail is { } sharp)
                {
                    shown.Add(sharp.PixelSize.Width);
                }
            };

            // Two zoom steps before the first decode ends: the second waits, and the first - never
            // the one wanted by then - is not handed to the card.
            top.WantSharper!(1200);
            top.WantSharper!(2900);
            await pane.PreviewsLoadedForTests;

            Assert.Equal([3072], shown);
            Assert.False(top.SharpDecoding);

            window.Close();
        });
    }

    // ###########################################################################################
    // Each picture a card: its colours as the board keeps them (blank ones as CRT draws them), its
    // preview read with the components marked on it. A colour picked goes into the draft as ONE step;
    // a picker writing back the colour it shows writes nothing - a blank colour stays blank.
    // ###########################################################################################
    [Fact]
    public async Task A_pictures_colour_picked_goes_into_the_draft_and_a_colour_shown_writes_nothing()
    {
        DraftSession session = this.OpenDraftWithPictures();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            Assert.Equal("Board: Schematic images (2)", Get<TextBlock>(pane, "TitleText").Text);
            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            ComponentPaneBoardPictureRow bottom = pane.BoardPicturesForTests[1];
            Assert.Equal((Avalonia.Media.Colors.Red, 30.0), (top.HighlightColor, top.HighlightOpacity));
            Assert.Equal((Avalonia.Media.Colors.IndianRed, 20.0), (bottom.HighlightColor, bottom.HighlightOpacity));
            Assert.Single(top.Highlights);
            Assert.NotNull(top.Thumbnail);
            Assert.Equal(new Avalonia.Size(1, 1), top.ImageSize);

            // The pickers writing back what they show: nothing changes.
            bottom.HighlightColor = Avalonia.Media.Colors.IndianRed;
            bottom.HighlightOpacity = 20;
            pane.CommitHeld();
            Assert.False(session.CanUndo);

            top.HighlightColor = Avalonia.Media.Colors.DodgerBlue;
            top.HighlightOpacity = 45;
            pane.CommitHeld();

            Assert.Equal("Colors of the schematic images", session.UndoDescription);
            BoardSchematicEntry written = session.Content.Board.Schematics[0];
            Assert.Equal(("DodgerBlue", "45%"), (written.SchematicHighlightColor, written.SchematicHighlightOpacity));
            Assert.Equal(string.Empty, session.Content.Board.Schematics[1].SchematicHighlightColor);
            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // A picture is read once while the part is open (code review, 2026-10-26): every re-read of the
    // part - an undo, a change made elsewhere - decoded every picture again, and the bitmaps piled up
    // until a collection happened to run.
    // ###########################################################################################
    [Fact]
    public async Task A_pictures_preview_is_read_once_and_kept_when_the_part_is_read_again()
    {
        DraftSession session = this.OpenDraftWithPictures();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;
            Avalonia.Media.Imaging.Bitmap first = pane.BoardPicturesForTests[0].Thumbnail!;

            Assert.True(session.Apply("Elsewhere", content => content with { Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Board data", NameOrHandle = "Bo" }]) }));
            Dispatcher.UIThread.RunJobs();
            await pane.PreviewsLoadedForTests;

            Assert.Same(first, pane.BoardPicturesForTests[0].Thumbnail);

            window.Close();
        });
    }

    // ###########################################################################################
    // A picture of a type the PNG/JPEG reader does not know has its own size read all the same (code
    // review, 2026-10-26): it was taken as the size it was decoded to, and its marks drawn far too big.
    // ###########################################################################################
    [Fact]
    public async Task A_BMP_pictures_own_size_is_what_its_marks_are_placed_in()
    {
        BoardData board = Published();
        board.Schematics.Add(new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.bmp" });
        board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "0", Y = "0", Width = "2", Height = "1" });

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [this.Stored("Images/top.bmp", ComponentPaneTests.Bmp(4000, 2))]));
        DraftSession session = DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            Assert.Equal(new Avalonia.Size(4000, 2), pane.BoardPicturesForTests.Single().ImageSize);

            window.Close();
        });
    }

    // "Use for all pictures": every picture in these colours, as one step.
    [Fact]
    public void Use_for_all_pictures_gives_every_picture_the_colours_in_one_step()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            pane.UseColoursForAll(pane.BoardPicturesForTests[0]);

            Assert.Equal("Colors of the schematic images", session.UndoDescription);
            Assert.All(session.Content.Board.Schematics, row => Assert.Equal(new PictureColours("Red", "30%", "", "Red", "100%"), row.ReadColours()));
            Assert.True(session.Undo());
            Assert.False(session.CanUndo);

            window.Close();
        });
    }

    // ---- the pictures themselves, on their cards (step 3 - the "Schematic images" window went) ----

    // The name box on a picture's card.
    private static TextBox NameBoxOf(ComponentPane pane, ComponentPaneBoardPictureRow row) =>
        pane.GetVisualDescendants().OfType<TextBox>().First(box => ReferenceEquals(box.DataContext, row) && box.Text == row.NameDraft);

    // ###########################################################################################
    // "Add pictures" (and pictures dropped on the tab): each copied in at the board's top under a name
    // of its own, named after its file and never a name a picture has - ONE step for all, the draft
    // saved at once; what is no picture is left out.
    // ###########################################################################################
    [Fact]
    public async Task Pictures_added_are_copied_in_named_after_their_files_as_one_step_and_saved_at_once()
    {
        DraftSession session = this.OpenDraftWithPictures();
        string top = this.Pickable("Top.png", ComponentPaneTests.Png);
        string under = this.Pickable("Under side.png", ComponentPaneTests.Png);
        string notes = this.Pickable("notes.txt", "x"u8.ToArray());

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            Assert.Contains("Accept", Get<Button>(pane, "AddPicturesButton").Classes);

            Assert.True(await pane.AddPicturesAsync([top, under, notes]));

            Assert.Equal(
                [("Top", "Images/top.png"), ("Bottom", "Images/bottom.png"), ("Top (2)", "Top.png"), ("Under side", "Under side.png")],
                session.Content.Board.Schematics.Select(row => (row.SchematicName, row.SchematicImageFile)));
            Assert.Equal("2 schematic images added", session.UndoDescription);
            Assert.Contains("Top.png", this.ListedOnDisk());
            Assert.Contains("Under side.png", this.ListedOnDisk());
            Assert.False(session.HasUnsavedChanges);

            Assert.Equal(4, pane.BoardPicturesForTests.Count);
            Assert.Equal("Board: Schematic images (4)", Get<TextBlock>(pane, "TitleText").Text);
            Assert.Equal(ComponentPaneWording.PicturesAdded(2), pane.PicturesNoteForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A PICTURE ADDED SHOWS ON ITS CARD AT ONCE (owner report, 2026-11-01: a new board's pictures
    // were grey boxes until the part was opened again). *** Its copy lies in the draft folder until the
    // save moves it into the draft's store; a preview read in between found it there and lost it to the
    // save. So no preview is decoded before that save is done, and every new card gets its picture.
    // ###########################################################################################
    [Fact]
    public async Task Pictures_added_are_read_for_their_cards_only_once_the_draft_is_saved()
    {
        DraftSession session = this.OpenDraftWithPictures();

        // Bytes of their own - the board's pictures' would be found decoded already.
        string top = this.Pickable("Top.png", PngOfSize(40, 20));
        string under = this.Pickable("Under side.png", PngOfSize(50, 30));

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            await pane.PreviewsLoadedForTests;

            var unsavedAtDecode = new System.Collections.Concurrent.ConcurrentBag<bool>();
            pane.PreviewDecodeForTests = _ =>
            {
                unsavedAtDecode.Add(session.HasUnsavedChanges);
                return true;
            };

            Assert.True(await pane.AddPicturesAsync([top, under]));
            await pane.PreviewsLoadedForTests;

            Assert.NotEmpty(unsavedAtDecode);
            Assert.DoesNotContain(true, unsavedAtDecode);
            Assert.All(
                pane.BoardPicturesForTests.Where(row => row.File is "Top.png" or "Under side.png"),
                row => Assert.NotNull(row.Thumbnail));

            window.Close();
        });
    }

    // Nothing but what is no picture - a text file, a hidden file: nothing added, and said why.
    [Fact]
    public async Task Nothing_is_added_when_no_picture_was_picked_and_it_says_why()
    {
        DraftSession session = this.OpenDraftWithPictures();
        string notes = this.Pickable("notes.txt", "x"u8.ToArray());
        string hidden = this.Pickable(".thumb.png", ComponentPaneTests.Png);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            Assert.False(await pane.AddPicturesAsync([notes, hidden]));

            Assert.False(session.CanUndo);
            Assert.Equal(2, session.Content.Board.Schematics.Count);
            Assert.Equal(ComponentPaneWording.NoPicturesToAdd, pane.PicturesNoteForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // A name typed renames the picture on Enter - ONE step, the components marked on it carried to
    // the new name - and the card is the SAME card (made again, it would throw away a click pressed on
    // another card as the box lost focus). Escape puts the name back.
    // ###########################################################################################
    [Fact]
    public void A_pictures_name_typed_renames_it_on_Enter_and_its_marks_go_with_it()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            TextBox box = NameBoxOf(pane, top);

            box.Text = "Something";
            box.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape });
            Assert.Equal("Top", box.Text);
            Assert.False(session.CanUndo);

            box.Text = " Front ";
            box.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });

            Assert.Equal("Schematic image [Top] renamed [Front]", session.UndoDescription);
            Assert.Equal(["Front", "Bottom"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Equal("Front", Assert.Single(session.Content.Board.ComponentHighlights).SchematicName);
            Assert.Same(top, pane.BoardPicturesForTests[0]);
            Assert.Equal("Front", top.Name);

            window.Close();
        });
    }

    // ###########################################################################################
    // A blank name, or one another picture has (capitals aside), is refused with the reason - and the
    // typed name STAYS in its box to be fixed (code review, 2026-10-28: it was put back to the old
    // name, and closing the pane then threw it away unseen). The draft is as it was.
    // ###########################################################################################
    [Fact]
    public void A_blank_name_or_one_another_picture_has_is_refused_with_the_reason_and_kept_in_its_box()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];

            top.NameDraft = "bottom";
            Assert.False(pane.CommitPictureName(top));
            Assert.Equal(SchematicRename.NameTaken("bottom"), pane.PicturesNoteForTests);
            Assert.Equal("bottom", top.NameDraft);

            top.NameDraft = "   ";
            Assert.False(pane.CommitPictureName(top));
            Assert.Equal(SchematicRename.NeedsAName, pane.PicturesNoteForTests);

            Assert.False(session.CanUndo);
            Assert.Equal(["Top", "Bottom"], session.Content.Board.Schematics.Select(row => row.SchematicName));

            // Escape puts the picture's name back.
            NameBoxOf(pane, top).RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape });
            Assert.Equal("Top", top.NameDraft);

            window.Close();
        });
    }

    // The pane stays open over a name that cannot be taken - the reason under the heading - whatever
    // is asked for in its place ("Board status" since the Close button went, 2026-10-30), and lets go
    // once the name is fixed.
    [Fact]
    public void The_pane_stays_open_over_a_picture_name_that_cannot_be_taken()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            pane.BoardPicturesForTests[0].NameDraft = "Bottom";

            Assert.True(pane.KeepsRefusedPictureName());
            Assert.Equal(SchematicRename.NameTaken("Bottom"), pane.PicturesNoteForTests);

            pane.BoardPicturesForTests[0].NameDraft = "Front";

            Assert.False(pane.KeepsRefusedPictureName());
            Assert.Equal("Front", session.Content.Board.Schematics[0].SchematicName);

            window.Close();
        });
    }

    // ###########################################################################################
    // AN IMAGE WITH NO KiCad DATA SHOWS NO KiCad CONTROLS (owner request, 2026-10-30: "do not show the
    // UI for KiCad relevant data ... less disabled elements to relate to"): the traces' colors and
    // opacities, the traces preview, Align KiCad and Mark all automatically from KiCad are hidden, not
    // greyed out - and shown on an image with KiCad data, the preview's ticks and Mark all then greyed
    // out only until KiCad is aligned on it.
    // ###########################################################################################
    [Fact]
    public void An_image_with_no_KiCad_data_shows_no_KiCad_controls()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            Dispatcher.UIThread.RunJobs();

            // Top names a KiCad view on a board with KiCad files; Bottom names none. A card's controls
            // are the ones its row is the DataContext of.
            List<Control> controls = Get<ItemsControl>(pane, "BoardPicturesItemsControl").GetVisualDescendants().OfType<Control>().ToList();

            Control Shown(ComponentPaneBoardPictureRow card, Func<Control, bool> which) =>
                controls.First(control => ReferenceEquals(control.DataContext, card) && which(control));

            bool Visible(ComponentPaneBoardPictureRow card, string text) =>
                Shown(card, control => control is TextBlock block && block.Text == text).IsEffectivelyVisible;

            bool ButtonVisible(ComponentPaneBoardPictureRow card, string content) =>
                Shown(card, control => control is Button button && button.Content as string == content).IsEffectivelyVisible;

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests.Single(row => row.Name == "Top");
            ComponentPaneBoardPictureRow bottom = pane.BoardPicturesForTests.Single(row => row.Name == "Bottom");

            foreach (string text in new[] { "Traces on this side", "Traces on the other side", "Traces preview" })
            {
                Assert.True(Visible(top, text), $"[{text}] should show on an image with KiCad data");
                Assert.False(Visible(bottom, text), $"[{text}] should not show on an image without KiCad data");
            }

            Assert.True(ButtonVisible(top, PictureToolWording.AlignButton));
            Assert.True(ButtonVisible(top, PictureToolWording.MarkAutomaticallyButton));
            Assert.False(ButtonVisible(bottom, PictureToolWording.AlignButton));
            Assert.False(ButtonVisible(bottom, PictureToolWording.MarkAutomaticallyButton));

            // The rest is everybody's.
            Assert.True(Visible(bottom, "Schematic image"));
            Assert.True(Visible(bottom, "Thumbnail"));
            Assert.True(ButtonVisible(bottom, PictureToolWording.MarkButton));

            window.Close();
        });
    }

    // ###########################################################################################
    // THE SETTINGS BESIDE THE PREVIEWS KEEP THEIR ROOM (owner request, 2026-10-30: they could be dragged
    // to 0px): a drag stops where the settings keep ContributeWidths.SettingsMinWidth, and the cards
    // made narrower narrow the previews shown, the width chosen kept.
    // ###########################################################################################
    [Fact]
    public void A_drag_of_the_previews_leaves_the_settings_their_room()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            pane.UseCardsWidth(900);
            pane.DragPreviewWidthForTests(+2000);

            double most = 900 - ContributeWidths.CardChrome - ContributeWidths.SettingsMinWidth;
            Assert.Equal(most, pane.PreviewWidth, 3);
            Assert.All(pane.BoardPicturesForTests, row => Assert.Equal(most, row.PreviewWidth, 3));

            // Narrower cards narrow what is shown; the width chosen stays for when they widen again.
            pane.UseCardsWidth(700);
            Assert.Equal(700 - ContributeWidths.CardChrome - ContributeWidths.SettingsMinWidth, pane.BoardPicturesForTests[0].PreviewWidth, 3);
            Assert.Equal(most, pane.PreviewWidth, 3);

            pane.UseCardsWidth(900);
            Assert.Equal(most, pane.BoardPicturesForTests[0].PreviewWidth, 3);

            window.Close();
        });
    }

    // ###########################################################################################
    // A picture's name is taken on Enter or when its box is left - never on a pause in typing (code
    // review, 2026-10-28: a pause after "Top" saved it as a rename, or refused it and put the old
    // name back under the caret). The pause - the colours' too - takes everything but the name.
    // ###########################################################################################
    [Fact]
    public void A_picture_name_is_never_taken_on_a_pause_in_typing()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow bottom = pane.BoardPicturesForTests[1];

            // "Top" typed over "Bottom" on the way to "Top side" - taken, or refused, on a pause.
            bottom.NameDraft = "Top";
            pane.CommitHeld(fromTyping: true);

            Assert.False(session.CanUndo);
            Assert.Equal("Top", bottom.NameDraft);
            Assert.Null(pane.PicturesNoteForTests);

            // The rest typed, and the box left: renamed.
            bottom.NameDraft = "Top side";
            pane.CommitHeld();
            Assert.Equal(["Top", "Top side"], session.Content.Board.Schematics.Select(row => row.SchematicName));

            window.Close();
        });
    }

    // Pictures picked that cannot be copied in - gone since they were picked - get no card, and are said.
    [Fact]
    public async Task Pictures_that_could_not_be_copied_get_no_card_and_are_said()
    {
        DraftSession session = this.OpenDraftWithPictures();
        string side = this.Pickable("Side.png", ComponentPaneTests.Png);
        string gone = Path.Combine(this.thisWorkspace.Root, "Picked", "Gone.png");

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            Assert.True(await pane.AddPicturesAsync([side, gone]));

            Assert.Equal(["Top", "Bottom", "Side"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Equal(ComponentPaneWording.PicturesAdded(1, notCopied: 1), pane.PicturesNoteForTests);
            Assert.Equal("Schematic image added", session.UndoDescription);

            window.Close();
        });
    }

    // ###########################################################################################
    // The Pictures part asks for the board's KiCad views when it opens, and again when the draft's
    // KiCad files change while it is open - a project imported (code review, 2026-10-28: the views
    // stayed as they were until the part was opened again). A read answers only its own ask.
    // ###########################################################################################
    [Fact]
    public void The_Pictures_part_asks_for_the_KiCad_views_again_when_the_KiCad_files_change()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();
            var asks = new List<int>();
            pane.KiCadViewsWanted += (_, ask) => asks.Add(ask);

            pane.ShowBoardPart(session, PaneBoardPart.Pictures, WithBoardParts());
            Dispatcher.UIThread.RunJobs();
            Assert.Equal([1], asks);

            // A KiCad project imported: its file arrives in the draft and a write lists it.
            Directory.CreateDirectory(Path.Combine(this.DraftFolder, "KiCad data"));
            File.WriteAllText(Path.Combine(this.DraftFolder, "KiCad data", "board.kicad_pcb"), "(kicad_pcb)");
            Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save(evenUnchanged: true));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal([1, 2], asks);

            // The read asked first ending last is not taken; the latest is.
            pane.UseKiCadViews(["PCB top", "PCB bottom"], 2);
            pane.UseKiCadViews(["Old view"], 1);
            Assert.Contains(pane.BoardPicturesForTests[1].CadNameChoices, item => item.CadName == "PCB bottom");
            Assert.DoesNotContain(pane.BoardPicturesForTests[1].CadNameChoices, item => item.CadName == "Old view");

            window.Close();
        });
    }

    // What the pane holds is handed over as the draft is asked for it - a name typed and not left yet
    // goes in with the rest (the pane closing, another board, Submit).
    [Fact]
    public void A_name_typed_and_not_left_goes_into_the_draft_when_the_pane_hands_over()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            pane.BoardPicturesForTests[1].NameDraft = "Underside";
            pane.CommitHeld();

            Assert.Equal(["Top", "Underside"], session.Content.Board.Schematics.Select(row => row.SchematicName));

            window.Close();
        });
    }

    // ###########################################################################################
    // "Associate KiCad file": the board's views once the host has read them - a view another picture has
    // left OUT of the list (owner request, 2026-10-10) - picked as ONE step, every list changed IN
    // PLACE (a list replaced lets go of its selection), and Align KiCad offered with it.
    // ###########################################################################################
    [Fact]
    public void A_KiCad_view_picked_is_the_pictures_CAD_name_and_no_other_pictures_choice()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            ComponentPaneBoardPictureRow bottom = pane.BoardPicturesForTests[1];
            Assert.False(bottom.ShowsCadNames);

            pane.UseKiCadViews(["PCB top", "PCB bottom", "Sheet 1"]);

            Assert.True(bottom.ShowsCadNames);
            Assert.DoesNotContain(bottom.CadNameChoices, item => item.CadName == "PCB top");
            Assert.Contains(top.CadNameChoices, item => item.CadName == "PCB top");
            Assert.Equal("PCB top", top.SelectedCadNameChoice?.CadName);
            Assert.False(bottom.CanAlignKiCad);
            CadNameChoiceItem topsOwn = top.SelectedCadNameChoice!;

            // Picked in the drop-down: written once the binding has finished.
            bottom.SelectedCadNameChoice = bottom.CadNameChoices.Single(item => item.CadName == "PCB bottom");
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("KiCad view of [Bottom]", session.UndoDescription);
            Assert.Equal("PCB bottom", session.Content.Board.Schematics[1].CadName);
            Assert.True(bottom.CanAlignKiCad);
            Assert.DoesNotContain(top.CadNameChoices, item => item.CadName == "PCB bottom");
            Assert.Same(topsOwn, top.SelectedCadNameChoice);
            Assert.Same(bottom, pane.BoardPicturesForTests[1]);

            window.Close();
        });
    }

    // ###########################################################################################
    // Up and down move a picture one place - the board's own order, ONE step - offered only where
    // they lead somewhere.
    // ###########################################################################################
    [Fact]
    public void Up_and_down_move_a_picture_one_place_as_one_step()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            Assert.Equal((false, true), (pane.BoardPicturesForTests[0].CanMoveUp, pane.BoardPicturesForTests[0].CanMoveDown));
            Assert.Equal((true, false), (pane.BoardPicturesForTests[1].CanMoveUp, pane.BoardPicturesForTests[1].CanMoveDown));

            pane.MovePicture(pane.BoardPicturesForTests[1], -1);

            Assert.Equal("Order of the schematic images", session.UndoDescription);
            Assert.Equal(["Bottom", "Top"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Equal(["Bottom", "Top"], pane.BoardPicturesForTests.Select(row => row.Name));
            Assert.False(pane.BoardPicturesForTests[0].CanMoveUp);

            // Already at the top: no step.
            pane.MovePicture(pane.BoardPicturesForTests[0], -1);
            Assert.True(session.Undo());
            Assert.False(session.CanUndo);

            window.Close();
        });
    }

    // ###########################################################################################
    // Remove asks first when components are marked on the picture - they go with it - and "no" keeps
    // everything; "yes" takes the picture and its marks off as ONE step, which Undo brings back.
    // ###########################################################################################
    [Fact]
    public async Task Removing_a_picture_with_marks_asks_first_and_takes_them_along()
    {
        DraftSession session = this.OpenDraftWithPictures();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            var asked = new List<(string Name, SchematicRemovalLoss Loss)>();
            bool answer = false;
            pane.ConfirmPictureRemovalOverrideForTests = (name, loss) =>
            {
                asked.Add((name, loss));
                return Task.FromResult(answer);
            };

            await pane.RemovePictureAsync(pane.BoardPicturesForTests[0]);

            Assert.Equal([("Top", new SchematicRemovalLoss(1, false))], asked);
            Assert.False(session.CanUndo);
            Assert.Equal(2, pane.BoardPicturesForTests.Count);

            answer = true;
            await pane.RemovePictureAsync(pane.BoardPicturesForTests[0]);

            Assert.Equal("Schematic image [Top] deleted", session.UndoDescription);
            Assert.Equal(["Bottom"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Empty(session.Content.Board.ComponentHighlights);
            Assert.Equal(["Bottom"], pane.BoardPicturesForTests.Select(row => row.Name));
            Assert.Equal(ComponentPaneWording.PictureRemoved("Top"), pane.PicturesNoteForTests);

            Assert.True(session.Undo());
            Assert.Equal(["Top", "Bottom"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Single(session.Content.Board.ComponentHighlights);

            window.Close();
        });
    }

    // A KiCad alignment on the picture is asked about too; a picture with nothing on it goes unasked.
    [Fact]
    public async Task Removing_asks_about_a_KiCad_alignment_and_nothing_about_a_bare_picture()
    {
        DraftSession session = this.OpenDraftWithPictures(alignBottom: true);

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            var asked = new List<SchematicRemovalLoss>();
            pane.ConfirmPictureRemovalOverrideForTests = (_, loss) =>
            {
                asked.Add(loss);
                return Task.FromResult(true);
            };

            await pane.RemovePictureAsync(pane.BoardPicturesForTests[1]);

            Assert.Equal([new SchematicRemovalLoss(0, true)], asked);
            Assert.Equal(["Top"], session.Content.Board.Schematics.Select(row => row.SchematicName));
            Assert.Empty(session.Content.Calibrations);

            window.Close();
        });
    }

    [Fact]
    public async Task A_picture_with_nothing_on_it_is_removed_without_a_question()
    {
        DraftSession session = this.OpenDraftWithPictures();

        await UiTest.RunAsync(async () =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            int asked = 0;
            pane.ConfirmPictureRemovalOverrideForTests = (_, _) =>
            {
                asked++;
                return Task.FromResult(false);
            };

            await pane.RemovePictureAsync(pane.BoardPicturesForTests[1]);

            Assert.Equal(0, asked);
            Assert.Equal("Schematic image [Bottom] deleted", session.UndoDescription);
            Assert.Equal(["Top"], session.Content.Board.Schematics.Select(row => row.SchematicName));

            window.Close();
        });
    }

    // ###########################################################################################
    // "Mark components" and "Align KiCad" on a picture (slice 6 - contributor mode went): what the
    // pane holds goes into the draft first, then the host is asked to open the tool on that picture.
    // Aligning needs the board's KiCad data and a CAD name for the picture - said when it cannot.
    // ###########################################################################################
    [Fact]
    public void A_pictures_tools_are_asked_of_the_host_after_the_pane_hands_over_what_it_holds()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            var asked = new List<(PictureTool Tool, string Picture)>();
            pane.PictureToolRequested += (_, request) => asked.Add(request);

            ComponentPaneBoardPictureRow top = pane.BoardPicturesForTests[0];
            Assert.False(top.CanAlignKiCad);
            Assert.Equal(PictureToolWording.AlignNeedsKiCadTip, top.AlignTip);

            top.HighlightColor = Avalonia.Media.Colors.DodgerBlue;
            pane.RequestPictureTool(top, PictureTool.MarkComponents);

            Assert.Equal([(PictureTool.MarkComponents, "Top")], asked);
            Assert.Equal("Colors of the schematic images", session.UndoDescription);

            window.Close();
        });
    }

    [Fact]
    public void Align_KiCad_can_be_pressed_for_a_picture_with_a_CAD_name_on_a_board_with_KiCad_data()
    {
        DraftSession session = this.OpenDraftWithPictures(withKiCad: true);

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);

            Assert.True(pane.BoardPicturesForTests[0].CanAlignKiCad);
            Assert.Equal(PictureToolWording.AlignTip, pane.BoardPicturesForTests[0].AlignTip);

            // No CAD name: no KiCad view to line up with.
            Assert.False(pane.BoardPicturesForTests[1].CanAlignKiCad);

            window.Close();
        });
    }

    // From a board part to a component: the component shown, the part's rows no longer committed.
    [Fact]
    public void A_component_shown_after_a_board_part_is_shown_as_a_component()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Credits);

            pane.Show(session, "U8", "PAL", WithBoardParts());

            Assert.Null(pane.BoardPart);
            Assert.False(Get<ScrollViewer>(pane, "BoardPartPanel").IsVisible);
            Assert.True(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);
            Assert.Equal("CPU", Get<TextBox>(pane, "FriendlyNameBox").Text);

            pane.CommitHeld();
            Assert.False(session.CanUndo);

            window.Close();
        });
    }

    // "Preview popup" asks for the component's popup - what the boxes hold in the draft first.
    [Fact]
    public void Preview_popup_asks_for_the_component_with_what_the_boxes_hold_in_the_draft()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session);

            var asked = new List<string>();
            pane.PreviewRequested += (_, label) => asked.Add(label);

            Get<TextBox>(pane, "DescriptionBox").Text = "Runs the show";
            Click(Get<Button>(pane, "PreviewPopupButton"));

            Assert.Equal(["U8"], asked);
            Assert.Equal("Runs the show", U8(session).Description);

            window.Close();
        });
    }
}
