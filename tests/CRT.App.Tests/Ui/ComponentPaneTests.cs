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

    // The published board with U8's picture, one scope baseline and a datasheet, stored in the draft.
    private DraftSession OpenDraftWithFiles()
    {
        BoardData board = Published();
        board.ComponentImages.AddRange([
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/U8.png" },
            new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/U8_1.png", TimeDiv = "50us", VoltsDiv = "1V", TriggerLevelVolts = "1.5V" },
        ]);
        board.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Component local files/6510.pdf" });

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

            Assert.Equal("Scope baselines (1)", Get<TextBlock>(pane, "BaselinesHeading").Text);
            Assert.Equal("Pictures (1)", Get<TextBlock>(pane, "PicturesHeading").Text);
            Assert.Equal("Files (1)", Get<TextBlock>(pane, "FilesHeading").Text);

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
            Assert.Equal("Scope baselines of [U8]", session.UndoDescription);

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
            Assert.StartsWith("A scope baseline needs its pin.", baseline.Problem, StringComparison.Ordinal);

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
            Assert.Equal("A scope baseline needs its pin and a screenshot.", Get<TextBlock>(pane, "NewBaselineProblemText").Text);

            Get<TextBox>(pane, "NewBaselinePinBox").Text = "2";
            Get<TextBox>(pane, "NewBaselineNameBox").Text = "PA3";
            await pane.ChooseNewBaselineScreenshotAsync();

            Assert.True(Get<Button>(pane, "AddNewBaselineButton").IsEnabled);

            await pane.AddNewBaselineAsync();

            ComponentImageEntry added = ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL")[^1];
            Assert.Equal(("2", "PA3", "Scope baseline/U8_2_PA3.png", "1V"), (added.Pin, added.Name, added.File, added.VoltsDiv));
            Assert.Contains("Scope baseline/U8_2_PA3.png", this.ListedOnDisk());
            Assert.False(Get<Border>(pane, "NewBaselinePanel").IsVisible);
            Assert.Equal("Scope baselines (2)", Get<TextBlock>(pane, "BaselinesHeading").Text);

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
            await pane.CaptureNewBaselineAsync();
            await pane.AddNewBaselineAsync();

            Assert.Equal("Scope baseline/U8_3.png", ComponentPaneEdits.BaselinesOf(session.Content.Board, "U8", "PAL")[^1].File);

            pane.ShowScopeConnected(false);
            Assert.False(Get<Button>(pane, "CaptureBaselineButton").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // A capture waits for "Add this baseline" in the user's OWN app data folder (code review,
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

            Assert.Equal(["Datasheet", "6510 errata"], ComponentPaneEdits.FilesOf(session.Content.Board, "U8").Select(row => row.Name));
            Assert.Contains("Component local files/6510 errata.txt", this.ListedOnDisk());
            Assert.Equal("Files of [U8]", session.UndoDescription);

            pane.FilesForTests[1].Name = "Errata";
            pane.CommitHeld();
            Assert.Equal("Errata", ComponentPaneEdits.FilesOf(session.Content.Board, "U8")[1].Name);

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
            Assert.DoesNotContain("Accept", Get<Button>(pane, "CloseButton").Classes);
            Assert.DoesNotContain("Danger", Get<Button>(pane, "CloseButton").Classes);

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
            Assert.Equal("U9", Get<TextBlock>(pane, "TitleText").Text);
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
    public void A_label_taken_is_refused_and_no_category_is_said_but_allowed()
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

            Assert.False(Get<TextBlock>(pane, "NewLabelProblemText").IsVisible);
            Assert.True(Get<TextBlock>(pane, "NewCategoryNoteText").IsVisible);
            Assert.True(Get<Button>(pane, "AddNewComponentButton").IsEnabled);
            Assert.True(pane.AddNewComponent());
            Assert.Equal(string.Empty, ComponentPaneEdits.RowShown(session.Content.Board, "U20", "PAL")!.Category);

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

    private static (ComponentPane Pane, Window Window) ShownPart(DraftSession session, PaneBoardPart part)
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 1400 };
        window.Show();

        pane.ShowBoardPart(session, part, WithBoardParts());
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    // ###########################################################################################
    // A board part in place of a component: its list grouped by category, counted in its heading -
    // and opening and leaving it changes nothing, though the board had the list in another order.
    // ###########################################################################################
    [Fact]
    public void A_board_part_shows_its_list_grouped_and_opening_it_changes_nothing()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.FilesAndLinks);

            Assert.Equal("Files and links", Get<TextBlock>(pane, "TitleText").Text);
            Assert.True(Get<ScrollViewer>(pane, "BoardPartPanel").IsVisible);
            Assert.False(Get<ScrollViewer>(pane, "SectionsScrollViewer").IsVisible);
            Assert.False(Get<WrapPanel>(pane, "ActionsPanel").IsVisible);
            Assert.True(Get<StackPanel>(pane, "FilesAndLinksPart").IsVisible);
            Assert.False(Get<StackPanel>(pane, "CreditsPart").IsVisible);
            Assert.Equal("Files (3)", Get<TextBlock>(pane, "BoardFilesHeading").Text);
            Assert.Equal(["Service manual", "Notes", "Dead test"], pane.BoardFilesForTests.Select(row => row.Name));
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

            Assert.Equal("Credits (1)", Get<TextBlock>(pane, "CreditsHeading").Text);

            Click(Get<Button>(pane, "AddCreditButton"));
            ComponentPaneCreditRow added = pane.CreditsForTests[^1];
            Assert.Equal(["Schematic image"], added.CategoryChoices);
            added.Category = "Board labelling";
            added.NameOrHandle = " Bo ";
            pane.CommitHeld();

            Assert.Equal("Credits", session.UndoDescription);
            Assert.Equal(["Anna", "Bo"], session.Content.Board.Credits.Select(row => row.NameOrHandle));
            Assert.True(Get<Border>(pane, "ChangedMark").IsVisible);

            Click(RowButtons(pane, "CreditsItemsControl").First(button => button.Content as string == "Remove"));

            Assert.Equal(["Bo"], session.Content.Board.Credits.Select(row => row.NameOrHandle));

            window.Close();
        });
    }

    // ###########################################################################################
    // "Add me" (2026-10-26): a credit by the name and contact CRT remembers - or, with none remembered,
    // an empty row that says to type one; once the row is in the draft with a name, it is remembered.
    // ###########################################################################################
    [Fact]
    public void Add_me_credits_the_remembered_name_and_remembers_one_typed()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Credits);
            var remembered = new List<(string, string)>();
            pane.CreditIdentity = () => ("Anna", "RHD");
            pane.RememberCreditIdentity = (name, contact) => remembered.Add((name, contact));

            pane.AddMe();

            ComponentPaneCreditRow mine = pane.CreditsForTests[^1];
            Assert.True(mine.IsMine);
            Assert.Equal(("Anna", "RHD"), (mine.NameOrHandle, mine.Contact));
            Assert.False(Get<TextBlock>(pane, "CreditMeHintText").IsVisible);

            // Nothing remembered: an empty row, and the line saying to type a name.
            pane.CreditIdentity = () => (string.Empty, string.Empty);
            pane.AddMe();
            Assert.True(Get<TextBlock>(pane, "CreditMeHintText").IsVisible);

            ComponentPaneCreditRow typed = pane.CreditsForTests[^1];
            pane.CreditsForTests[^2].NameOrHandle = string.Empty;
            typed.Category = "Board data";
            typed.NameOrHandle = " Bo ";
            typed.Contact = "https://example.org";
            pane.CommitHeld();

            Assert.Equal(("Bo", "https://example.org"), remembered[^1]);
            Assert.False(Get<TextBlock>(pane, "CreditMeHintText").IsVisible);
            Assert.Contains(session.Content.Board.Credits, credit => credit.NameOrHandle == "Bo" && credit.Category == "Board data");

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
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.FilesAndLinks);
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);

            await pane.AddBoardFileAsync();

            BoardLocalFileEntry added = session.Content.Board.BoardLocalFiles.Single(row => row.Name == "Voltages");
            Assert.Equal("Board local files/Voltages.png", added.File);
            Assert.Equal("Troubleshooting", added.Category);
            Assert.Equal("The board's files", session.UndoDescription);
            Assert.Contains("Board local files/Voltages.png", this.ListedOnDisk());
            Assert.False(session.HasUnsavedChanges);

            window.Close();
        });
    }

    // ###########################################################################################
    // An important signal picks its net from the board's KiCad nets once they are read; one naming
    // a net the KiCad data does not have says the Schematics tab will not light it up. Without KiCad
    // data, any net may be typed and the part says why there is nothing to pick.
    // ###########################################################################################
    [Fact]
    public void An_important_signal_naming_a_net_the_KiCad_data_lacks_says_so()
    {
        DraftSession session = this.OpenDraftWithBoardParts();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.ImportantSignals);

            Assert.True(Get<TextBlock>(pane, "KiCadNetsLineText").IsVisible);
            Assert.Equal("Reading the board's KiCad nets...", Get<TextBlock>(pane, "KiCadNetsLineText").Text);
            Assert.False(pane.SignalsForTests[0].HasProblem);

            pane.UseNetChoices(KiCadNetNames.Choices([("PHI2", "PHI2", false), ("/RESET", "RESET", false)]));

            Assert.False(Get<TextBlock>(pane, "KiCadNetsLineText").IsVisible);
            Assert.Equal(["PHI2", "RESET"], pane.SignalsForTests[0].NetChoices);
            Assert.False(pane.SignalsForTests[0].HasProblem);

            Click(Get<Button>(pane, "AddSignalButton"));
            ComponentPaneSignalRow added = pane.SignalsForTests[^1];
            added.DisplayName = "Reset";
            added.KiCadNetName = "NRESET";
            pane.CommitHeld();

            Assert.Equal("Important signals", session.UndoDescription);
            Assert.True(added.HasProblem);
            Assert.Contains("[NRESET]", added.Problem, StringComparison.Ordinal);

            added.KiCadNetName = "/RESET";
            pane.CommitHeld();
            Assert.False(added.HasProblem);

            pane.UseNetChoices(KiCadNetChoices.None);
            Assert.Equal(ComponentPaneWording.NoKiCadNets, Get<TextBlock>(pane, "KiCadNetsLineText").Text);

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
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.FilesAndLinks);

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

            Click(RowButtons(pane, "MeasuringNotesItemsControl").Single());

            Assert.Single(opened);

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
            pane.FindOnOtherBoards = (_, _) => Task.FromResult<IReadOnlyList<ComponentCopySource>>([OtherPla(), OtherPla() with { BoardName = "Commodore 64 - 250469" }]);
            pane.ApplyToOtherBoards = (here, targets, parts) =>
            {
                handed.Add((here.Row.BoardLabel, targets.Select(target => target.BoardName).ToList(), parts.ToList()));
                return Task.FromResult("Changed on 1 board, in its own draft: Commodore 64 - 250425.");
            };

            Get<TextBox>(pane, "PartNumberBox").Text = "906114-01";
            Get<TextBox>(pane, "RepairNotesBox").Text = "Socket it";
            pane.CommitHeld();
            Assert.True(also.IsVisible);

            await pane.OpenAlsoPanelAsync();

            Assert.True(pane.IsAlsoPanelOpen);
            List<CheckBox> boards = Get<StackPanel>(pane, "AlsoBoardsPanel").Children.OfType<CheckBox>().ToList();
            Assert.Equal(2, boards.Count);
            Assert.All(boards, box => Assert.True(box.IsChecked));
            Assert.Equal(["Repair notes"], Get<StackPanel>(pane, "AlsoPartsPanel").Children.OfType<CheckBox>().Select(box => box.Content as string));

            boards[1].IsChecked = false;
            await pane.ChangeOtherBoardsAsync();

            (string label, List<string> handedBoards, List<ComponentCopyPart> handedParts) = Assert.Single(handed);
            Assert.Equal("U8", label);
            Assert.Equal(["Commodore 64 - 250425"], handedBoards);
            Assert.Equal([ComponentCopyPart.RepairNotes], handedParts);
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

    // ---- the pictures and their colours (slice 4b) ----------------------------------------------

    // Two pictures - one with its colours set, one blank (CRT's defaults) - and a component marked on
    // the first; the first picture's bytes stored, so its preview can be read. `withKiCad`: the board
    // lists a KiCad file and the first picture names its KiCad view (slice 6's "Align KiCad").
    private DraftSession OpenDraftWithPictures(bool withKiCad = false)
    {
        BoardData board = Published();
        board.Schematics.AddRange([
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = withKiCad ? "PCB top" : "", SchematicHighlightColor = "Red", SchematicHighlightOpacity = "30%", ThumbnailHighlightColor = "Red", ThumbnailHighlightOpacity = "100%" },
            new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "Images/bottom.png" },
        ]);
        board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "0", Y = "0", Width = "1", Height = "1" });

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            withKiCad
                ? [this.Stored("Images/top.png", ComponentPaneTests.Png), this.Stored("KiCad data/board.kicad_pcb", "(kicad_pcb)"u8.ToArray())]
                : [this.Stored("Images/top.png", ComponentPaneTests.Png)]));

        return DraftSession.Open(this.DraftsRoot, ComponentPaneTests.BoardKey)!;
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

            Assert.Equal("Pictures (2)", Get<TextBlock>(pane, "BoardPicturesHeading").Text);
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

            Assert.Equal("Colours of the pictures", session.UndoDescription);
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

            Assert.Equal("Colours of the pictures", session.UndoDescription);
            Assert.All(session.Content.Board.Schematics, row => Assert.Equal(new PictureColours("Red", "30%", "", "Red", "100%"), row.ReadColours()));
            Assert.True(session.Undo());
            Assert.False(session.CanUndo);

            window.Close();
        });
    }

    // Adding, removing and renaming pictures is the "Schematic images" window's - asked of the host.
    [Fact]
    public void The_pictures_part_asks_the_host_for_the_Schematic_images_window()
    {
        DraftSession session = this.OpenDraftWithPictures();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = ShownPart(session, PaneBoardPart.Pictures);
            int asked = 0;
            pane.ManagePicturesRequested += (_, _) => asked++;

            Click(Get<Button>(pane, "ManagePicturesButton"));

            Assert.Equal(1, asked);

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
            Assert.Equal("Colours of the pictures", session.UndoDescription);

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
