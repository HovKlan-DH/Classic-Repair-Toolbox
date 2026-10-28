using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE START PANEL'S LISTS (TabDrafts.Readiness.cs - CRT 4.0 step 3, slice 3a: "Getting this board
// ready ... On an existing board, the same panel shows What's missing").
//
// A new board - one only a draft has - shows its checklist, each step done by what the draft held in
// memory says, with a button to the tool that does it; the KiCad data's facts arrive a moment later
// into the SAME rows. An existing board shows what it lacks, each component a click away, worked out
// again when the start panel comes back.
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

    // The step buttons, by the step they do.
    private static Dictionary<ReadinessStepKind, Button> StepButtons(TabDrafts tab) =>
        tab.GetControl<StackPanel>("ReadinessStepsPanel")
            .GetLogicalDescendants()
            .OfType<Button>()
            .Where(button => button.Tag is ReadinessStepKind)
            .ToDictionary(button => (ReadinessStepKind)button.Tag!);

    private static async Task Until(Func<bool> condition)
    {
        for (int attempt = 0; attempt < 50 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }
    }

    // ###########################################################################################
    // A new board without KiCad data: the checklist from its draft, the KiCad steps skipped with no
    // button, the next step's button there - and the two buttons that go ahead and make something
    // green (CLAUDE.md, button colours). Nothing of "What's missing".
    // ###########################################################################################
    [Fact]
    public async Task A_new_board_shows_its_checklist_from_its_draft()
    {
        WriteDraft(NewBoard, Started());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("999001", NewBoard, draftOnly: true);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            Assert.True(tab.GetControl<StackPanel>("ReadinessPanel").IsVisible);
            Assert.False(tab.GetControl<StackPanel>("MissingPanel").IsVisible);

            Dictionary<ReadinessStepKind, ReadinessStep> steps = tab.ReadinessForTests.ToDictionary(step => step.Kind);
            Assert.Equal(9, steps.Count);
            Assert.Equal(ReadinessState.Done, steps[ReadinessStepKind.Pictures].State);
            Assert.Equal(ReadinessState.Skipped, steps[ReadinessStepKind.Match].State);
            Assert.Equal(ReadinessState.Next, steps[ReadinessStepKind.Mark].State);

            Dictionary<ReadinessStepKind, Button> buttons = StepButtons(tab);
            Assert.False(buttons[ReadinessStepKind.Name].IsVisible);
            Assert.False(buttons[ReadinessStepKind.Match].IsVisible);
            Assert.False(buttons[ReadinessStepKind.Signals].IsVisible);
            Assert.True(buttons[ReadinessStepKind.Mark].IsVisible);
            Assert.Equal("Mark components", buttons[ReadinessStepKind.Mark].Content);
            Assert.Equal("KiCad data", buttons[ReadinessStepKind.KiCad].Content);

            Assert.Contains("Accept", buttons[ReadinessStepKind.Components].Classes);
            Assert.Contains("Accept", buttons[ReadinessStepKind.Submit].Classes);
            Assert.DoesNotContain("Accept", buttons[ReadinessStepKind.Mark].Classes);
            Assert.DoesNotContain("Accept", buttons[ReadinessStepKind.KiCad].Classes);

            // No button label ends in "..." (CLAUDE.md).
            Assert.All(buttons.Values, button => Assert.DoesNotContain("...", button.Content as string ?? string.Empty, StringComparison.Ordinal));
        });
    }

    // ###########################################################################################
    // With KiCad files in the draft the steps are shown at once, then again once the KiCad data is
    // read - into the SAME rows, so a button pressed while it arrives is not thrown away.
    // ###########################################################################################
    [Fact]
    public async Task The_KiCad_facts_arrive_into_the_same_rows()
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
            Assert.Equal("Reading the KiCad data...", tab.ReadinessForTests.Single(step => step.Kind == ReadinessStepKind.Match).Detail);
            Assert.Equal("Components (1)", tab.ReadinessForTests.Single(step => step.Kind == ReadinessStepKind.Components).Title);
            Button before = StepButtons(tab)[ReadinessStepKind.Components];

            read.SetResult(root);
            await tab.ReadinessShownForTests;

            Assert.Equal([NewBoard], asked);
            Dictionary<ReadinessStepKind, ReadinessStep> steps = tab.ReadinessForTests.ToDictionary(step => step.Kind);
            Assert.Equal("Components - 2 found in KiCad", steps[ReadinessStepKind.Components].Title);
            Assert.Equal(ReadinessState.Done, steps[ReadinessStepKind.Match].State);
            Assert.Equal(ReadinessState.Next, steps[ReadinessStepKind.Align].State);

            Assert.Same(before, StepButtons(tab)[ReadinessStepKind.Components]);
            Assert.Equal(9, tab.GetControl<StackPanel>("ReadinessStepsPanel").Children.Count);
            Assert.True(StepButtons(tab)[ReadinessStepKind.Match].IsVisible);
        });
    }

    // A step's button opens its tool: "Important signals" the board's list in the pane, "Add
    // component" the pane's form for a new one.
    [Fact]
    public async Task A_steps_button_opens_the_tool_that_does_it()
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

            StepButtons(tab)[ReadinessStepKind.Signals].RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.Equal(PaneBoardPart.ImportantSignals, tab.ComponentPaneForTests.BoardPart);
            Assert.False(tab.GetControl<StackPanel>("StartPanel").IsVisible);

            tab.CloseComponent();
            Assert.True(tab.GetControl<StackPanel>("StartPanel").IsVisible);

            StepButtons(tab)[ReadinessStepKind.Components].RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.True(tab.ComponentPaneForTests.IsNewComponentForm);
        });
    }

    // ###########################################################################################
    // "Add N from KiCad" (the plan: "one press adds every KiCad component that isn't on the board
    // yet"): into the draft held in memory as ONE step - saved, and one Undo away - with what came of
    // it said under the checklist, and the step done once all are there.
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

            Button add = StepButtons(tab)[ReadinessStepKind.Components];
            Assert.Equal("Add 2 from KiCad", add.Content);
            Assert.Contains("Accept", add.Classes);

            add.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.ReadinessNoteForTests is not null);
            await tab.ReadinessShownForTests;

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard)!;
            Assert.Equal(["U1", "U2", "C1"], session.Content.Board.Components.Select(row => row.BoardLabel));
            Assert.Equal(20, session.Content.Board.ComponentPins.Count(row => row.BoardLabel == "U2"));
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal([NewBoard], reloaded);

            Assert.Equal("Added 2 components from KiCad - 1 with its pins.", tab.ReadinessNoteForTests);
            ReadinessStep components = tab.ReadinessForTests.Single(step => step.Kind == ReadinessStepKind.Components);
            Assert.Equal(ReadinessState.Done, components.State);
            Assert.Equal("Add component", StepButtons(tab)[ReadinessStepKind.Components].Content);

            // One step: one Undo takes all of them back.
            Assert.Equal(DraftStepWording.ComponentsFromKiCad(2), session.UndoDescription);
            session.Undo();
            Assert.Equal(["U1"], session.Content.Board.Components.Select(row => row.BoardLabel));
        });
    }

    // The suggested matches' "Use" buttons, in the order shown, and "Use all".
    private static List<Button> UseButtons(TabDrafts tab) =>
        tab.GetControl<StackPanel>("ReadinessStepsPanel")
            .GetLogicalDescendants()
            .OfType<Button>()
            .Where(button => button.Tag is SchematicCadNameSuggestion)
            .ToList();

    private static Button UseAllButton(TabDrafts tab) =>
        tab.GetControl<StackPanel>("ReadinessStepsPanel")
            .GetLogicalDescendants()
            .OfType<Button>()
            .Single(button => button.Content is string label && label.StartsWith("Use all", StringComparison.Ordinal));

    // ###########################################################################################
    // The plan's "Match: suggested pairs ("...Top.png" with "PCB Top") to confirm": listed under the
    // Match step once the KiCad data is read, each with a green "Use" - one taken is ONE step, its
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
            Assert.Equal(ReadinessState.Done, tab.ReadinessForTests.Single(step => step.Kind == ReadinessStepKind.Match).State);
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
    // Another new board shown: the checklist that was on screen goes at once, never left standing
    // under the next board's header (code review, 2026-10-26) - and for good when that board's draft
    // cannot be opened, its buttons acting on the board shown.
    // ###########################################################################################
    [Fact]
    public async Task Another_boards_checklist_never_stands_under_this_board()
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
            Assert.True(tab.GetControl<StackPanel>("ReadinessPanel").IsVisible);

            tab.ShowBoard(second, new BoardData(), "PAL", null);
            Assert.False(tab.GetControl<StackPanel>("ReadinessPanel").IsVisible);
            Assert.Empty(tab.ReadinessForTests);

            await tab.ReadinessShownForTests;
            Assert.False(tab.GetControl<StackPanel>("ReadinessPanel").IsVisible);
        });
    }

    // ###########################################################################################
    // The start panel back - the pane closed - shows the board's KiCad facts as they were while they
    // are read again (code review, 2026-10-26: each came back as if unknown, "Add N from KiCad" turned
    // into "Add component" for a moment, and a click then opened the wrong thing).
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
            Assert.Equal("Add 1 from KiCad", StepButtons(tab)[ReadinessStepKind.Components].Content);

            var read = new TaskCompletionSource<KiCadProjectRoot?>();
            tab.BoardKiCadOverrideForTests = _ => read.Task;

            await tab.OpenComponentAsync("U1");
            tab.CloseComponent();
            await Until(() => tab.ReadinessForTests.Count > 0);

            Assert.Equal("Add 1 from KiCad", StepButtons(tab)[ReadinessStepKind.Components].Content);

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

            await tab.AddFromKiCadAsync();

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

            await tab.AddFromKiCadAsync();
            await tab.ReadinessShownForTests;
            Assert.NotNull(tab.ReadinessNoteForTests);

            DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard)!.Undo();
            tab.ShowBoard(entry, Started(), "PAL", null);
            await tab.ReadinessShownForTests;

            Assert.Null(tab.ReadinessNoteForTests);
            Assert.Equal("Add 1 from KiCad", StepButtons(tab)[ReadinessStepKind.Components].Content);
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

            await tab.AddFromKiCadAsync();

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

            await tab.AddFromKiCadAsync();

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

            Button before = tab.GetControl<StackPanel>("MissingGroupsPanel").GetLogicalDescendants().OfType<Button>().First();

            tab.ShowBoard(entry, board, "PAL", null);

            Assert.Same(before, tab.GetControl<StackPanel>("MissingGroupsPanel").GetLogicalDescendants().OfType<Button>().First());
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

            Assert.False(tab.GetControl<StackPanel>("ReadinessPanel").IsVisible);
            Assert.True(tab.GetControl<StackPanel>("MissingPanel").IsVisible);
            Assert.Equal(["U8"], tab.MissingForTests!.WithoutDescription);
            Assert.Equal(["U8", "U9"], tab.MissingForTests.Unmarked);

            List<Button> labels = tab.GetControl<StackPanel>("MissingGroupsPanel").GetLogicalDescendants().OfType<Button>().ToList();
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
                tab.GetControl<StackPanel>("MissingGroupsPanel").GetLogicalDescendants().OfType<TextBlock>()
                    .Select(text => text.Text).Where(text => text?.StartsWith("Not marked", StringComparison.Ordinal) == true));
        });
    }

    // Nothing lacking: one line saying so, no groups.
    [Fact]
    public void A_board_lacking_nothing_says_so()
    {
        WriteDraft(OldBoard, new BoardData());

        UiTest.Run(() =>
        {
            HardwareBoardEntry entry = Entry("250407", OldBoard, draftOnly: false);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            // The groups are kept, hidden, for the next board (made once - code review, 2026-10-26).
            StackPanel groups = tab.GetControl<StackPanel>("MissingGroupsPanel");
            Assert.Equal(BoardReadiness.NothingMissing, Assert.IsType<TextBlock>(Assert.Single(groups.Children, child => child.IsVisible)).Text);
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

            StackPanel groups = tab.GetControl<StackPanel>("MissingGroupsPanel");
            Assert.Equal(BoardReadiness.ListedAtMost, groups.GetLogicalDescendants().OfType<Button>().Count());
            Assert.Contains(
                groups.GetLogicalDescendants().OfType<TextBlock>(),
                text => text.Text == "and 3 more");
        });
    }
}
