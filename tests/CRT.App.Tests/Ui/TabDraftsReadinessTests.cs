using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "BOARD STATUS" (TabDrafts.Readiness.cs; owner request, 2026-11-01: "each point in its own panel with a
// headline, then an explanation what this is about, a button for it and lastly also a status").
//
// A panel per part of the board, new or published, each from what the draft held in memory says, with
// a button to the part; the KiCad data's facts arrive a moment later into the SAME panels. Under
// Components what each lacks, a click away. Worked out again when the page comes back. What a panel
// says is BoardStatusPage's (BoardStatusPageTests) - these cover how the tab shows it.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsReadinessTests : IDisposable
{
    private const string NewBoard = "Commodore/C64/999001/board.json";
    private const string OldBoard = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsReadinessTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Entry(string board, string excelDataFile, bool draftOnly) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = board,
        ExcelDataFile = excelDataFile,
        IsDraftOnly = draftOnly,
    };

    // A draft of `board` listing `files` (their bytes in the draft's own store).
    private static void WriteDraft(string excelDataFile, BoardData board, params string[] files)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        string folder = Path.GetDirectoryName(boardFile)!;
        Directory.CreateDirectory(folder);

        var listed = new List<BoardDocument.StoredFileRow>();

        foreach (string name in files)
        {
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(name);
            string sha = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes));
            string stored = Path.Combine(folder, DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(stored)!);
            File.WriteAllBytes(stored, bytes);
            listed.Add(new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = bytes.Length });
        }

        BoardFiles.Write(boardFile, BoardDocumentMapping.ToDocument(new BoardContent(board, [], []), name => name, listed));
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile), new DraftMarker { BoardKey = excelDataFile });
    }

    private static TabDrafts Tab(params HardwareBoardEntry[] entries)
    {
        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [.. entries],
            PublishedBoardOverrideForTests = _ => null,
        };

        tab.RefreshDrafts();
        return tab;
    }

    // A board just made: a picture, one component, nothing marked.
    private static BoardData Started() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" }],
        Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "CPU" }],
    };

    // A panel's lines, as it says them.
    private static List<string> LinesOf(TabDrafts tab, BoardStatusPart part) =>
        tab.StatusPanelForTests(part).Lines.Select(line => line.Text).ToList();

    // The buttons of the components listed under the Components panel.
    private static List<Button> ComponentButtons(TabDrafts tab) =>
        tab.GetControl<StackPanel>("StatusPanelsHost")
            .GetLogicalDescendants()
            .OfType<Button>()
            .Where(button => button.Tag is string)
            .ToList();

    // Long enough for a busy machine, and failing HERE when it never comes - half a second, given up
    // silently, failed later as something else (2026-11-01).
    private static async Task Until(Func<bool> condition)
    {
        for (int attempt = 0; attempt < 3000 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(condition(), "What this test waits for never came.");
    }

    // ###########################################################################################
    // A new board without KiCad data: its panels from its draft, in the page's order, each with its
    // button - plain, every one of them opens its part (CLAUDE.md, button colors), none with "...". The
    // KiCad project is optional, so its panel is neither done nor missing, and neither is Important
    // KiCad signals, which need it.
    // ###########################################################################################
    [Fact]
    public async Task A_new_board_shows_its_panels_from_its_draft()
    {
        WriteDraft(NewBoard, Started());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            Assert.True(tab.GetControl<StackPanel>("StatusPanelsHost").IsVisible);

            // Its picture's file is not there: said on the images' panel, and no panel of problems.
            Assert.Equal(
                [
                    BoardStatusPart.Name, BoardStatusPart.Regions, BoardStatusPart.KiCad, BoardStatusPart.Signals, BoardStatusPart.Pictures,
                    BoardStatusPart.Categories, BoardStatusPart.Components, BoardStatusPart.Files, BoardStatusPart.Links, BoardStatusPart.Credits,
                ],
                tab.StatusPanelsForTests.Select(panel => panel.Part));

            Assert.Contains(LinesOf(tab, BoardStatusPart.Pictures), line => line.StartsWith("Error - ", StringComparison.Ordinal));
            Assert.True(tab.StatusTableButtonForTests(BoardStatusPart.Pictures).IsVisible);
            Assert.False(tab.StatusTableButtonForTests(BoardStatusPart.Files).IsVisible);

            Assert.Equal(BoardStatusState.Done, tab.StatusPanelForTests(BoardStatusPart.Name).State);
            Assert.Equal(BoardStatusState.Optional, tab.StatusPanelForTests(BoardStatusPart.KiCad).State);
            Assert.Equal(BoardStatusState.Error, tab.StatusPanelForTests(BoardStatusPart.Pictures).State);
            Assert.Equal(BoardStatusState.Optional, tab.StatusPanelForTests(BoardStatusPart.Signals).State);

            // U1 is marked on no image: a warning, said by the group of buttons.
            Assert.Equal(BoardStatusState.ToDo, tab.StatusPanelForTests(BoardStatusPart.Components).State);
            Assert.Equal(BoardStatusState.ToDo, tab.StatusPanelForTests(BoardStatusPart.Credits).State);

            foreach (BoardStatusPanel panel in tab.StatusPanelsForTests)
            {
                Button button = tab.StatusButtonForTests(panel.Part);

                Assert.True(tab.StatusBorderForTests(panel.Part).IsVisible);
                Assert.Equal(panel.Action, button.Content);
                Assert.DoesNotContain("Accept", button.Classes);
                Assert.DoesNotContain("...", panel.Action, StringComparison.Ordinal);
            }

            // Its frame and its icon say how far it is - classes, so the theme's colors apply.
            Assert.Contains("done", tab.StatusBorderForTests(BoardStatusPart.Name).Classes);
            Assert.Contains("todo", tab.StatusBorderForTests(BoardStatusPart.Credits).Classes);
            Assert.Contains("optional", tab.StatusBorderForTests(BoardStatusPart.KiCad).Classes);
        });
    }

    // ###########################################################################################
    // With KiCad files in the draft the panels are shown at once, then again once the KiCad data is
    // read - into the SAME panels, so a button pressed while it arrives is not thrown away.
    // ###########################################################################################
    [Fact]
    public async Task The_KiCad_facts_arrive_into_the_same_panels()
    {
        WriteDraft(NewBoard, Started(), "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);

            var read = new TaskCompletionSource<KiCadProjectRoot?>();
            var root = new KiCadProjectRoot
            {
                Project = new KiCadProjectInfo { Views = [new KiCadProjectView { DisplayName = "PCB top" }] },
                Pcb =
                [
                    new KiCadPcb
                    {
                        Footprints =
                        [
                            new KiCadPcbFootprint { Reference = "U1", Pads = [new KiCadPcbPad { Number = "1" }] },
                            new KiCadPcbFootprint { Reference = "U2", Pads = [new KiCadPcbPad { Number = "1" }] },
                        ],
                    },
                ],
            };

            var asked = new List<string>();
            tab.BoardKiCadOverrideForTests = identity =>
            {
                asked.Add(identity);
                return read.Task;
            };

            tab.ShowBoard(entry, Started(), "PAL", null);
            await Until(() => asked.Count == 1);

            // Shown from the draft while the KiCad data is read.
            Assert.Contains(BoardStatusPage.ReadingKiCad, LinesOf(tab, BoardStatusPart.KiCad));
            Assert.Contains("1 KiCad file imported", LinesOf(tab, BoardStatusPart.KiCad));
            Assert.DoesNotContain(LinesOf(tab, BoardStatusPart.Components), line => line.Contains("not on the board yet", StringComparison.Ordinal));
            Button before = tab.StatusButtonForTests(BoardStatusPart.Components);
            int panels = tab.GetControl<StackPanel>("StatusPanelsHost").Children.Count;

            read.SetResult(root);
            await tab.ReadinessShownForTests;

            Assert.Equal([NewBoard], asked);
            Assert.DoesNotContain(BoardStatusPage.ReadingKiCad, LinesOf(tab, BoardStatusPart.KiCad));
            Assert.Contains("2 components can be marked automatically", LinesOf(tab, BoardStatusPart.KiCad));
            Assert.Contains("1 image known by KiCad data", LinesOf(tab, BoardStatusPart.Pictures));
            Assert.Contains("1 KiCad file not aligned yet", LinesOf(tab, BoardStatusPart.Pictures));
            Assert.Contains(LinesOf(tab, BoardStatusPart.Components), line => line.Contains("not on the board yet", StringComparison.Ordinal));
            Assert.Contains("1 component is not on the board yet", LinesOf(tab, BoardStatusPart.KiCad));

            Assert.Same(before, tab.StatusButtonForTests(BoardStatusPart.Components));
            Assert.Equal(panels, tab.GetControl<StackPanel>("StatusPanelsHost").Children.Count);
        });
    }

    // A panel's button opens its part: "Important KiCad signals" the board's list in the pane,
    // "Components" the list in the left column.
    [Fact]
    public async Task A_panels_button_opens_its_part()
    {
        WriteDraft(NewBoard, Started(), "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(null);
            tab.NetChoicesOverrideForTests = _ => Task.FromResult(KiCadNetNames.Choices([]));
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            tab.StatusButtonForTests(BoardStatusPart.Signals).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.Equal(PaneBoardPart.ImportantSignals, tab.ComponentPaneForTests.BoardPart);
            Assert.False(tab.GetControl<StackPanel>("StartPanel").IsVisible);

            tab.CloseComponent();
            Assert.True(tab.GetControl<StackPanel>("StartPanel").IsVisible);

            tab.StatusButtonForTests(BoardStatusPart.Components).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.Null(tab.ComponentPaneForTests.BoardPart);
            Assert.Null(tab.SelectedNavEntryForTests);
        });
    }

    // ###########################################################################################
    // Adding the KiCad components (the plan: "one press adds every KiCad component that isn't on the
    // board yet" - the KiCad import's since 2026-11-01, owner request: "Then you can delete this option
    // from 'Board status'"): into the draft held in memory as ONE step - saved, and one Undo away - with
    // what came of it said over the panels, and Components done once all are there.
    // ###########################################################################################
    [Fact]
    public async Task Add_from_KiCad_puts_every_missing_component_into_the_draft_as_one_step()
    {
        WriteDraft(NewBoard, Started(), "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);

            static KiCadPcbFootprint Chip(string reference, string value, int pads) => new()
            {
                Reference = reference,
                Value = value,
                FootprintName = $"Package_DIP:DIP-{pads}_W7.62mm",
                Pads = [.. Enumerable.Range(1, pads).Select(number => new KiCadPcbPad { Number = $"{number}", PinFunction = number == 1 ? "~{OE}" : "" })],
            };

            var root = new KiCadProjectRoot
            {
                Pcb =
                [
                    new KiCadPcb
                    {
                        Footprints =
                        [
                            Chip("U1", "6510", 40),
                            Chip("U2", "74LS373", 20),
                            new KiCadPcbFootprint { Reference = "C1", Value = "100nF", Pads = [new KiCadPcbPad { Number = "1" }, new KiCadPcbPad { Number = "2" }] },
                        ],
                    },
                ],
            };

            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(root);
            var reloaded = new List<string>();
            tab.ReloadBoardOverrideForTests = reloaded.Add;

            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            Assert.Equal(BoardStatusState.ToDo, tab.StatusPanelForTests(BoardStatusPart.Components).State);
            Assert.Contains(
                "2 components in the KiCad data are not on the board yet - \"KiCad: Files\" adds them",
                LinesOf(tab, BoardStatusPart.Components));

            await AddFromKiCad(tab);
            await tab.ReadinessShownForTests;

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard)!;
            Assert.Equal(["U1", "U2", "C1"], session.Content.Board.Components.Select(row => row.BoardLabel));
            Assert.Equal(20, session.Content.Board.ComponentPins.Count(row => row.BoardLabel == "U2"));
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal([NewBoard], reloaded);

            Assert.Equal("Added 2 components from KiCad - 1 with its pins.", tab.ReadinessNoteForTests);

            // Nothing of the KiCad data is left to add - marked on no image yet, they are still to do.
            Assert.DoesNotContain(LinesOf(tab, BoardStatusPart.Components), line => line.Contains("not on the board yet", StringComparison.Ordinal));
            Assert.Equal(BoardStatusState.ToDo, tab.StatusPanelForTests(BoardStatusPart.Components).State);

            // One step: one Undo takes all of them back.
            Assert.Equal(DraftStepWording.ComponentsFromKiCad(2), session.UndoDescription);
            session.Undo();
            Assert.Equal(["U1"], session.Content.Board.Components.Select(row => row.BoardLabel));
        });
    }

    // A KiCad board of three DIP-16 chips, U1 to U3, as a .kicad_pcb file.
    private const string ThreeChipPcb = """
    (kicad_pcb (version 20221018) (generator pcbnew)
      (net 0 "")
      (net 1 "GND")
      (footprint "Package_DIP:DIP-16_W7.62mm" (layer "F.Cu") (at 100 50) (property "Reference" "U1")
        (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 1 "GND")))
      (footprint "Package_DIP:DIP-16_W7.62mm" (layer "F.Cu") (at 120 50) (property "Reference" "U2")
        (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 1 "GND")))
      (footprint "Package_DIP:DIP-16_W7.62mm" (layer "F.Cu") (at 140 50) (property "Reference" "U3")
        (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 1 "GND")))
    )
    """;

    // ###########################################################################################
    // *** THE KiCad IMPORT ADDS THE KiCad BOARD'S COMPONENTS (owner request, 2026-11-01: "it should add
    // all components when importing the KiCad data. Then you can delete this option from 'Board status'
    // ... Have checkbox if it should add all data automatically also?"). *** Unticked - as remembered -
    // an import adds none, and "KiCad data"'s summary offers them, one press adding them as one step;
    // ticked - the choice remembered - an import adds them itself and says so.
    // ###########################################################################################
    [Fact]
    public async Task A_KiCad_import_adds_the_KiCad_components_while_its_box_is_ticked()
    {
        WriteDraft(NewBoard, Started());

        string source = Path.Combine(this.thisWorkspace.Root, "My KiCad project");
        Directory.CreateDirectory(source);
        File.WriteAllText(Path.Combine(source, "board.kicad_pcb"), ThreeChipPcb);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(KiCadWith("U1", "U2", "U3"));
            tab.ReloadBoardOverrideForTests = _ => { };

            bool adds = false;
            var remembered = new List<bool>();
            tab.KiCadImportAddsComponents = () => adds;
            tab.RememberKiCadImportAddsComponents = value =>
            {
                remembered.Add(value);
                adds = value;
            };

            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.OpenKiCadAsync();

            KiCadDataPanel panel = Assert.IsType<KiCadDataPanel>(tab.KiCadPanelForTests);
            await panel.Opening;

            CheckBox box = panel.GetControl<CheckBox>("AddComponentsCheckBox");
            Button addButton = panel.GetControl<Button>("AddKiCadComponentsButton");
            Assert.False(box.IsChecked);
            Assert.Empty(remembered);

            // Unticked: the import adds none, and the summary offers them.
            await panel.ImportKiCadFolderAsync(source);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard)!;
            Assert.Equal(["U1"], session.Content.Board.Components.Select(row => row.BoardLabel));
            Assert.Null(panel.ComponentsAddedNoteForTests);
            Assert.True(addButton.IsVisible);
            Assert.Equal("Add 2 components from KiCad", addButton.Content);
            Assert.Contains("Accept", addButton.Classes);

            // One press adds them, as one step.
            addButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => session.Content.Board.Components.Count == 3);
            await panel.DraftChanging;

            Assert.Equal(["U1", "U2", "U3"], session.Content.Board.Components.Select(row => row.BoardLabel));
            Assert.Equal(DraftStepWording.ComponentsFromKiCad(2), session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
            Assert.StartsWith("Added 2 components from KiCad", panel.ComponentsAddedNoteForTests);
            Assert.False(addButton.IsVisible);

            // Ticked - and remembered - the next import adds them itself.
            session.Undo();
            box.IsChecked = true;
            Assert.Equal([true], remembered);

            await panel.ImportKiCadFolderAsync(source);

            Assert.Equal(["U1", "U2", "U3"], session.Content.Board.Components.Select(row => row.BoardLabel));
            Assert.Equal(DraftStepWording.ComponentsFromKiCad(2), session.UndoDescription);
            Assert.StartsWith("Added 2 components from KiCad", panel.ComponentsAddedNoteForTests);
            Assert.False(addButton.IsVisible);
        });
    }

    // ###########################################################################################
    // *** "KiCad DATA" FOLLOWS THE HEADER'S UNDO AND REDO (code review, 2026-11-01). *** It was a modal
    // window until 2026-10-30; a panel stays on screen beside them, and went on saying the components
    // were added, with its "Add N components from KiCad" hidden and its report judging the KiCad data
    // against the labels from before - until something else was opened and it was opened again.
    // ###########################################################################################
    [Fact]
    public async Task The_KiCad_panel_follows_the_headers_undo_and_redo()
    {
        WriteDraft(NewBoard, Started());

        string source = Path.Combine(this.thisWorkspace.Root, "My KiCad project");
        Directory.CreateDirectory(source);
        File.WriteAllText(Path.Combine(source, "board.kicad_pcb"), ThreeChipPcb);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(KiCadWith("U1", "U2", "U3"));

            var reloaded = new List<string>();
            tab.ReloadBoardOverrideForTests = reloaded.Add;
            tab.KiCadImportAddsComponents = () => true;

            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.OpenKiCadAsync();

            KiCadDataPanel panel = Assert.IsType<KiCadDataPanel>(tab.KiCadPanelForTests);
            await panel.Opening;
            await panel.ImportKiCadFolderAsync(source);

            Button addButton = panel.GetControl<Button>("AddKiCadComponentsButton");
            Assert.StartsWith("Added 2 components from KiCad", panel.ComponentsAddedNoteForTests);
            Assert.False(addButton.IsVisible);
            Assert.DoesNotContain("U2", KiCadUnlabelled(panel));

            // The header's Undo: they are gone, and the panel says so.
            reloaded.Clear();
            await tab.UndoAsync(entry);
            await panel.DraftChanging;

            Assert.Null(panel.ComponentsAddedNoteForTests);
            Assert.True(addButton.IsVisible);
            Assert.Equal("Add 2 components from KiCad", addButton.Content);
            Assert.Contains("U2", KiCadUnlabelled(panel));

            // Followed once, and no change of its own: the board on screen read again once.
            Assert.Equal([NewBoard], reloaded);

            // The header's Redo: they are back.
            await tab.RedoAsync(entry);
            await panel.DraftChanging;

            Assert.False(addButton.IsVisible);
            Assert.DoesNotContain("U2", KiCadUnlabelled(panel));
        });
    }

    // The references the panel's report lists as not labelled on the board yet.
    private static List<string> KiCadUnlabelled(KiCadDataPanel panel) =>
        (panel.GetControl<ItemsControl>("UnusedList").ItemsSource ?? Array.Empty<string>()).Cast<string>().ToList();

    // ###########################################################################################
    // "Add N components from KiCad" - the way a contributor adds them since 2026-11-01: on the KiCad
    // data panel's summary, which catches up with the change as it ends, and back to Board status.
    // ###########################################################################################
    private static async Task AddFromKiCad(TabDrafts tab)
    {
        await tab.OpenKiCadAsync();

        KiCadDataPanel panel = Assert.IsType<KiCadDataPanel>(tab.KiCadPanelForTests);
        await panel.Opening;
        await panel.AddKiCadComponentsNowAsync();

        await tab.CloseKiCadAsync();
    }

    // The suggested matches' "Use" buttons, in the order shown, and "Use all".
    private static List<Button> UseButtons(TabDrafts tab) =>
        tab.GetControl<StackPanel>("StatusPanelsHost")
            .GetLogicalDescendants()
            .OfType<Button>()
            .Where(button => button.Tag is SchematicCadNameSuggestion)
            .ToList();

    private static Button UseAllButton(TabDrafts tab) =>
        tab.GetControl<StackPanel>("StatusPanelsHost")
            .GetLogicalDescendants()
            .OfType<Button>()
            .Single(button => button.Content is string label && label.StartsWith("Use all", StringComparison.Ordinal));

    // ###########################################################################################
    // The plan's "Match: suggested pairs ("...Top.png" with "PCB Top") to confirm": listed under
    // Schematic images once the KiCad data is read, each with a green "Use" - one taken is ONE step, its
    // CAD name in the draft held in memory and saved, the rest left standing in their same rows; "Use
    // all" takes the rest as one step.
    // ###########################################################################################
    [Fact]
    public async Task Suggested_matches_are_offered_and_each_Use_is_one_step()
    {
        var board = new BoardData
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Top (replica)", SchematicImageFile = "Images/Issue 4.F.png" },
                new BoardSchematicEntry { SchematicName = "Bottom (replica)", SchematicImageFile = "Images/Issue 4.B.png" },
                new BoardSchematicEntry { SchematicName = "Main sheet", SchematicImageFile = "Images/scan.png" },
            ],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
        };

        WriteDraft(NewBoard, board, "KiCad data/Issue 4.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);

            var root = new KiCadProjectRoot
            {
                Project = new KiCadProjectInfo
                {
                    Views =
                    [
                        new KiCadProjectView { DisplayName = "Issue 4 - PCB Top", Type = "pcb_top", SourceFile = "Issue 4.kicad_pcb" },
                        new KiCadProjectView { DisplayName = "Issue 4 - PCB Bottom", Type = "pcb_bottom", SourceFile = "Issue 4.kicad_pcb" },
                        new KiCadProjectView { DisplayName = "Main sheet", Type = "schematic", SourceFile = "issue4.kicad_sch" },
                    ],
                },
            };

            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(root);
            var reloaded = new List<string>();
            tab.ReloadBoardOverrideForTests = reloaded.Add;

            tab.ShowBoard(entry, board, "PAL", null);
            await tab.ReadinessShownForTests;

            List<Button> uses = UseButtons(tab);
            Assert.Equal(["Top (replica)", "Bottom (replica)", "Main sheet"], uses.Select(button => ((SchematicCadNameSuggestion)button.Tag!).SchematicName));
            Assert.All(uses, button => Assert.Contains("Accept", button.Classes));
            Assert.All(uses, button => Assert.Equal("Use", button.Content));
            Assert.Equal("Use all 3 suggestions", UseAllButton(tab).Content);
            Assert.Contains("Accept", UseAllButton(tab).Classes);

            uses[0].RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.ReadinessNoteForTests is not null);
            await tab.ReadinessShownForTests;

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard)!;
            Assert.Equal(["Issue 4 - PCB Top", "", ""], session.Content.Board.Schematics.Select(row => row.CadName));
            Assert.Equal(DraftStepWording.CadName("Top (replica)"), session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal([NewBoard], reloaded);
            Assert.Equal("[Top (replica)] now shows [Issue 4 - PCB Top].", tab.ReadinessNoteForTests);

            // The rest left standing - the very rows they were (a row made again under a press threw
            // that click away).
            List<Button> rest = UseButtons(tab);
            Assert.Equal([uses[1], uses[2]], rest);
            Assert.Equal("Use all 2 suggestions", UseAllButton(tab).Content);

            UseAllButton(tab).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.ReadinessNoteForTests == "2 pictures now show their KiCad views.");
            await tab.ReadinessShownForTests;

            Assert.Equal(["Issue 4 - PCB Top", "Issue 4 - PCB Bottom", "Main sheet"], session.Content.Board.Schematics.Select(row => row.CadName));
            Assert.Equal(DraftStepWording.PicturesMatched(["Bottom (replica)", "Main sheet"]), session.UndoDescription);
            Assert.Empty(UseButtons(tab));
            Assert.False(UseAllButton(tab).IsVisible);
            Assert.Contains("3 KiCad files mapped", LinesOf(tab, BoardStatusPart.Pictures));
        });
    }

    // ###########################################################################################
    // A suggestion the draft moved past meanwhile - its picture given a CAD name another way - is
    // taken as nothing: no step, nothing read again, and said.
    // ###########################################################################################
    [Fact]
    public async Task A_suggestion_the_draft_moved_past_is_no_step()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
        };

        WriteDraft(NewBoard, board, "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(null);
            var reloaded = new List<string>();
            tab.ReloadBoardOverrideForTests = reloaded.Add;

            tab.ShowBoard(entry, board, "PAL", null);
            await tab.ReadinessShownForTests;

            DraftSession session = (await DraftEdits.OpenForChangeAsync(NewBoard))!;
            session.Apply("By hand", content => content with { Board = SchematicRows.WithCadName(content.Board, new SchematicRowRef("Top", "Images/top.png"), "Elsewhere").Board });
            string? undo = session.UndoDescription;

            await tab.UseSuggestedMatchesAsync([new SchematicCadNameSuggestion("Top", "Images/top.png", "Board - PCB Top")]);
            await tab.ReadinessShownForTests;

            Assert.Equal("Elsewhere", session.Content.Board.Schematics[0].CadName);
            Assert.Equal(undo, session.UndoDescription);
            Assert.Empty(reloaded);
            Assert.Equal(BoardReadiness.MatchesUsed([]), tab.ReadinessNoteForTests);
        });
    }

    // KiCad data with these chips, each a DIP-16 on the first PCB.
    private static KiCadProjectRoot KiCadWith(params string[] references) => new()
    {
        Pcb =
        [
            new KiCadPcb
            {
                Footprints =
                [
                    .. references.Select(reference => new KiCadPcbFootprint
                    {
                        Reference = reference,
                        Value = "4164",
                        FootprintName = "Package_DIP:DIP-16_W7.62mm",
                        Pads = [.. Enumerable.Range(1, 16).Select(number => new KiCadPcbPad { Number = $"{number}" })],
                    }),
                ],
            },
        ],
    };

    // ###########################################################################################
    // Another new board shown: the page that was on screen goes at once, never left standing under the
    // next board's header (code review, 2026-10-26) - and for good when that board's draft cannot be
    // opened, its buttons acting on the board shown.
    // ###########################################################################################
    [Fact]
    public async Task Another_boards_page_never_stands_under_this_board()
    {
        const string Broken = "Commodore/C64/999002/board.json";

        WriteDraft(NewBoard, Started());

        // A draft with its marker and no board file - nothing to open.
        string marker = DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, Broken);
        Directory.CreateDirectory(Path.GetDirectoryName(marker)!);
        DraftMarkerStore.Save(marker, new DraftMarker { BoardKey = Broken });

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry first = Entry("999001", NewBoard, draftOnly: true);
            HardwareBoardEntry second = Entry("999002", Broken, draftOnly: true);
            TabDrafts tab = Tab(first, second);

            tab.ShowBoard(first, Started(), "PAL", null);
            await tab.ReadinessShownForTests;
            Assert.True(tab.GetControl<StackPanel>("StatusPanelsHost").IsVisible);

            tab.ShowBoard(second, new BoardData(), "PAL", null);

            // Nothing of the first board stands: what shows is this board's, as far as known.
            Assert.DoesNotContain("1 component", LinesOf(tab, BoardStatusPart.Components));

            // Its draft cannot be opened: nothing shows at all.
            await tab.ReadinessShownForTests;
            Assert.False(tab.GetControl<StackPanel>("StatusPanelsHost").IsVisible);
            Assert.Empty(tab.StatusPanelsForTests);
        });
    }

    // ###########################################################################################
    // The start panel back - the pane closed - shows the board's KiCad facts as they were while they
    // are read again (code review, 2026-10-26: each came back as if unknown - the KiCad components the
    // board lacks forgotten for a moment, and a click then opened the wrong thing).
    // ###########################################################################################
    [Fact]
    public async Task The_start_panel_back_keeps_the_KiCad_facts_while_they_are_read_again()
    {
        WriteDraft(NewBoard, Started(), "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(KiCadWith("U1", "U2"));
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;
            Assert.Contains(
                "1 component in the KiCad data is not on the board yet - \"KiCad: Files\" adds it",
                LinesOf(tab, BoardStatusPart.Components));

            var read = new TaskCompletionSource<KiCadProjectRoot?>();
            tab.BoardKiCadOverrideForTests = _ => read.Task;

            await tab.OpenComponentAsync("U1");
            tab.CloseComponent();
            await Until(() => tab.StatusPanelsForTests.Count > 0);

            Assert.Contains(
                "1 component in the KiCad data is not on the board yet - \"KiCad: Files\" adds it",
                LinesOf(tab, BoardStatusPart.Components));

            read.SetResult(KiCadWith("U1", "U2"));
            await tab.ReadinessShownForTests;
        });
    }

    // ###########################################################################################
    // The step is named for what it adds to the draft AS THE CHANGE FINDS IT (code review,
    // 2026-10-26): written outside CRT meanwhile with one of the two in it, it read "2 components
    // added from KiCad" and added one.
    // ###########################################################################################
    [Fact]
    public async Task Add_from_KiCad_counts_on_the_draft_as_the_change_finds_it()
    {
        WriteDraft(NewBoard, Started(), "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(KiCadWith("U1", "U2", "U3"));
            tab.ReloadBoardOverrideForTests = _ => { };
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            // U2 put in by another program while CRT shows the board.
            BoardData written = Started();
            written.Components.Add(new ComponentEntry { BoardLabel = "U2", Category = "IC" });
            WriteDraft(NewBoard, written, "KiCad data/board.kicad_pcb");

            await AddFromKiCad(tab);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard)!;
            Assert.Equal(["U1", "U2", "U3"], session.Content.Board.Components.Select(row => row.BoardLabel));
            Assert.Equal(DraftStepWording.ComponentsFromKiCad(1), session.UndoDescription);
            Assert.Equal("Added 1 component from KiCad - 1 with its pins.", tab.ReadinessNoteForTests);
        });
    }

    // What "Add N from KiCad" did is said only while the draft is as it left it - an Undo takes the
    // words with it (code review, 2026-10-26).
    [Fact]
    public async Task The_note_goes_once_the_draft_moves_on()
    {
        WriteDraft(NewBoard, Started(), "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(KiCadWith("U1", "U2"));
            tab.ReloadBoardOverrideForTests = _ => tab.ShowBoard(entry, Started(), "PAL", null);
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            await AddFromKiCad(tab);
            await tab.ReadinessShownForTests;
            Assert.NotNull(tab.ReadinessNoteForTests);

            DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard)!.Undo();
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            Assert.Null(tab.ReadinessNoteForTests);
            Assert.Contains(
                "1 component in the KiCad data is not on the board yet - \"KiCad: Files\" adds it",
                LinesOf(tab, BoardStatusPart.Components));
        });
    }

    // ###########################################################################################
    // A draft that cannot be opened is said to be gone - never "every component is on the board
    // already" (code review, 2026-10-26) - and nothing is read again when nothing changed.
    // ###########################################################################################
    [Fact]
    public async Task With_no_draft_to_add_to_it_says_so()
    {
        string marker = DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, NewBoard);
        Directory.CreateDirectory(Path.GetDirectoryName(marker)!);
        DraftMarkerStore.Save(marker, new DraftMarker { BoardKey = NewBoard });

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(KiCadWith("U1"));
            var reloaded = new List<string>();
            tab.ReloadBoardOverrideForTests = reloaded.Add;
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await AddFromKiCad(tab);

            Assert.Equal(DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.NoDraft), tab.DraftProblemForTests);
            Assert.Empty(reloaded);
        });
    }

    // KiCad data that cannot be read adds nothing - and reads nothing again (code review, 2026-10-26).
    [Fact]
    public async Task Nothing_added_reads_nothing_again()
    {
        WriteDraft(NewBoard, Started(), "KiCad data/board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.BoardKiCadOverrideForTests = _ => Task.FromResult<KiCadProjectRoot?>(null);
            var reloaded = new List<string>();
            tab.ReloadBoardOverrideForTests = reloaded.Add;
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            await AddFromKiCad(tab);

            Assert.Equal(BoardReadiness.KiCadNotRead, tab.ReadinessNoteForTests);
            Assert.Empty(reloaded);
        });
    }

    // ###########################################################################################
    // The board read again with nothing missing changed keeps "What's missing"'s buttons - the very
    // controls, so a press going down as Main's reload ends is not thrown away (code review,
    // 2026-10-26; the drafts list's "click it twice", 2026-10-10).
    // ###########################################################################################
    [Fact]
    public void The_board_read_again_keeps_the_missing_components_buttons()
    {
        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] };
        WriteDraft(OldBoard, board);

        UiTest.Run(() =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, board, "PAL", null);

            Button before = ComponentButtons(tab)[0];

            tab.ShowBoard(entry, board, "PAL", null);

            Assert.Same(before, ComponentButtons(tab)[0]);
        });
    }

    // ###########################################################################################
    // The lists are worked out ONCE per board shown, and not at all while the tab is off screen - then
    // when it comes on (code review, 2026-10-26: twice per board, for every board load in any tab).
    // ###########################################################################################
    [Fact]
    public async Task The_lists_are_worked_out_once_and_only_on_screen()
    {
        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }] };
        WriteDraft(OldBoard, board);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            HardwareBoardEntry other = Entry("250425", "Commodore/C64/250425/board.json", draftOnly: false);
            TabDrafts tab = Tab(entry, other);
            tab.PublishedBoardOverrideForTests = _ => board;
            tab.ShowBoard(entry, board, "PAL", null);

            // Another board shown while the pane is open: the start panel comes back with it.
            await tab.OpenComponentAsync("U8");
            int asks = tab.ReadinessAsksForTests;
            tab.ShowBoard(other, board, "PAL", null);
            Assert.Equal(asks + 1, tab.ReadinessAsksForTests);

            // Off screen: nothing worked out - until the tab is shown.
            bool onScreen = false;
            tab.IsOnScreenOverrideForTests = () => onScreen;
            asks = tab.ReadinessAsksForTests;
            tab.ShowBoard(entry, board, "PAL", null);
            Assert.Equal(asks, tab.ReadinessAsksForTests);

            onScreen = true;
            var window = new Window { Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(asks + 1, tab.ReadinessAsksForTests);
            Assert.Equal(["U8"], tab.MissingForTests!.WithoutDescription);
            window.Close();
        });
    }

    // ###########################################################################################
    // An existing board says what it lacks - its components by label, each opening in the pane - and
    // works it out again from the draft held in memory when the start panel comes back.
    // ###########################################################################################
    [Fact]
    public async Task An_existing_board_shows_what_it_lacks_and_each_component_opens()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" },
                new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "RAM", Description = "4164" },
            ],
        };

        WriteDraft(OldBoard, board);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => board;
            tab.ShowBoard(entry, board, "PAL", null);

            Assert.True(tab.GetControl<StackPanel>("StatusPanelsHost").IsVisible);
            Assert.Equal(["U8"], tab.MissingForTests!.WithoutDescription);
            Assert.Equal(["U8", "U9"], tab.MissingForTests.Unmarked);

            List<Button> labels = ComponentButtons(tab);
            Assert.Equal(["U8", "U8", "U9"], labels.Select(button => button.Content as string));

            labels[0].RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);
            Assert.Equal("U8", tab.ComponentPaneForTests.Label);

            // U8 given its short description in the draft held in memory, then the pane closed: the
            // board shown is behind, the list is not.
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, OldBoard)!;
            tab.CloseComponent();
            session.Apply("Description", content => content with
            {
                Board = new BoardData
                {
                    Components =
                    [
                        new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU", Description = "6510" },
                        new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "RAM", Description = "4164" },
                    ],
                },
            });

            await tab.OpenComponentAsync("U9");
            tab.CloseComponent();

            Assert.Empty(tab.MissingForTests!.WithoutDescription);
            Assert.Equal(
                [BoardReadiness.UnmarkedHeading(2)],
                tab.GetControl<StackPanel>("StatusPanelsHost").GetLogicalDescendants().OfType<TextBlock>()
                    .Select(text => text.Text).Where(text => text?.StartsWith("Warning - not marked", StringComparison.Ordinal) == true));
        });
    }

    // ###########################################################################################
    // Nothing lacking - a picture and a credit there, no component to mark: its panels done, no Errors
    // and warnings panel. The groups under Components are kept, hidden, for the next board (made once -
    // code review, 2026-10-26).
    // ###########################################################################################
    [Fact]
    public void A_board_lacking_nothing_shows_its_panels_done()
    {
        WriteDraft(OldBoard, Whole(), "Images/top.png");

        UiTest.Run(() =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Whole(), "PAL", null);

            Assert.Equal([], tab.MissingForTests!.BoardGaps.Select(gap => gap.Text));

            Assert.DoesNotContain(tab.StatusPanelsForTests, panel => panel.Part == BoardStatusPart.Problems);
            Assert.False(tab.StatusBorderForTests(BoardStatusPart.Problems).IsVisible);

            Assert.Equal(BoardStatusState.Done, tab.StatusPanelForTests(BoardStatusPart.Pictures).State);
            Assert.Equal(BoardStatusState.Done, tab.StatusPanelForTests(BoardStatusPart.Credits).State);
            Assert.Empty(ComponentButtons(tab));
        });

        static BoardData Whole() => new()
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
            Credits = [new CreditEntry { Category = "Schematic image", NameOrHandle = "Anna" }],
        };
    }

    // ###########################################################################################
    // BOARD STATUS FLAGS WHAT THE BOARD LACKS (owner request, 2026-10-30), on the panel of the part that
    // lacks it - here no schematic image and no credit, the Credits panel's button opening that part in
    // the pane.
    // ###########################################################################################
    [Fact]
    public async Task Board_status_flags_what_the_board_lacks_on_its_panels()
    {
        WriteDraft(OldBoard, new BoardData());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => new BoardData();
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            Assert.Contains(BoardReadiness.NoPicturesGap, LinesOf(tab, BoardStatusPart.Pictures));
            Assert.Equal(BoardStatusState.ToDo, tab.StatusPanelForTests(BoardStatusPart.Pictures).State);
            Assert.Contains(BoardReadiness.NoCreditsGap, LinesOf(tab, BoardStatusPart.Credits));
            Assert.Equal(BoardReadiness.CreditsAction, tab.StatusButtonForTests(BoardStatusPart.Credits).Content);

            tab.StatusButtonForTests(BoardStatusPart.Credits).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.Equal(PaneBoardPart.Credits, tab.ComponentPaneForTests.BoardPart);
            Assert.Equal("CreditsNavButton", tab.SelectedNavEntryForTests);
        });
    }

    // ###########################################################################################
    // *** A ROW WITH A FIELD IT NEEDS EMPTY IS FLAGGED IN BOARD STATUS, BY ITS LIST, AND STOPS SUBMIT
    // (owner request, 2026-11-01: "These errors must be flagged in 'Board status' and then it should
    // not be possible to submit the draft board"). *** A credit with no name or handle is an error of
    // the draft's row - Submit says it sends nothing - and Board status says it on the Credits panel,
    // an error there, never also in Errors and warnings: its button opens the credits, where the box is
    // marked.
    // ###########################################################################################
    [Fact]
    public async Task A_credit_missing_its_name_is_flagged_in_Board_status_and_stops_Submit()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
            Credits = [new CreditEntry { Category = "Schematic image" }],
        };

        WriteDraft(OldBoard, board, "Images/top.png");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => new BoardData();
            tab.ShowBoard(entry, board, "PAL", null);

            DraftListItem row = Assert.Single(tab.Drafts);
            Assert.True(row.HasErrors);
            Assert.Equal("1 error", row.ErrorsText);
            Assert.Contains("to fix first", row.SubmitTooltip, StringComparison.Ordinal);

            Assert.DoesNotContain(tab.StatusPanelsForTests, panel => panel.Part == BoardStatusPart.Problems);
            Assert.Equal(BoardStatusState.Error, tab.StatusPanelForTests(BoardStatusPart.Credits).State);
            Assert.Contains(BoardReadiness.IncompleteGap(1, "credit", "credits"), LinesOf(tab, BoardStatusPart.Credits));

            tab.StatusButtonForTests(BoardStatusPart.Credits).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.Equal(PaneBoardPart.Credits, tab.ComponentPaneForTests.BoardPart);
            Assert.True(tab.ComponentPaneForTests.CreditsForTests[0].NameOrHandleMissing);
        });
    }

    // A long list names its first labels and counts the rest.
    [Fact]
    public void A_long_list_of_components_names_the_first_and_counts_the_rest()
    {
        var board = new BoardData
        {
            Components = [.. Enumerable.Range(1, 11).Select(number => new ComponentEntry { BoardLabel = $"U{number}", Category = "IC", Description = "x" })],
        };

        WriteDraft(OldBoard, board);

        UiTest.Run(() =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, board, "PAL", null);

            Assert.Equal(BoardReadiness.ListedAtMost, ComponentButtons(tab).Count);
            Assert.Contains(
                tab.GetControl<StackPanel>("StatusPanelsHost").GetLogicalDescendants().OfType<TextBlock>(),
                text => text.Text == "and 3 more");
        });
    }
}
