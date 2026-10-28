using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE WORKSPACE (TabDrafts.Workspace.cs - CRT 4.0 step 2, 2026-10-25: "One tab instead of
// Contribute and Drafts ... A My drafts list in the header switches between boards ... The table
// view is one view of the same draft").
//
// The tab is about the board shown in CRT's drop-down lists (ShowBoard): its draft's header - or
// what it is, with none - its parts and components on the left, and the table view on the right.
// Every action on the left works on the board's draft, made first when it has none.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsWorkspaceTests : IDisposable
{
    private const string BoardA = "Commodore/C64/250407/board.json";
    private const string BoardB = "Commodore/C64/250425/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsWorkspaceTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Entry(string board, string excelDataFile) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = board,
        ExcelDataFile = excelDataFile,
    };

    private static void WriteDraft(string excelDataFile, BoardData? content = null)
    {
        var board = content ?? new BoardData();

        if (content is null)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = "U0", Category = "IC", FriendlyName = "CPU" });
        }

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { BoardKey = excelDataFile });
    }

    // A board of three components - one PAL only, one NTSC only - with every part counted.
    private static BoardData Board() => new()
    {
        RevisionDate = "2026-October-4",
        Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "VIC", Region = "PAL" },
            new ComponentEntry { BoardLabel = "U2", Category = "IC", FriendlyName = "VIC", Region = "NTSC" },
            new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" },
        ],
        KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "CLK", KiCadNetName = "CLK" }],
        BoardLinks = [new BoardLinkEntry { Name = "Forum" }],
        Credits = [],
    };

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

    private static string Text(TabDrafts tab, string name) => tab.GetControl<TextBlock>(name).Text ?? string.Empty;

    // ###########################################################################################
    // The board shown, in the left column: its parts counted, its components in the region shown,
    // and the ones its draft changed marked.
    // ###########################################################################################
    [Fact]
    public void The_left_column_counts_the_boards_parts_and_lists_its_components_in_the_region_shown()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            tab.ShowBoard(Entry("250407", BoardA), Board(), "PAL", new HashSet<string>(["U8"], StringComparer.OrdinalIgnoreCase));

            Assert.Equal("1", Text(tab, "PicturesCountText"));
            Assert.Equal("1", Text(tab, "SignalsCountText"));
            Assert.Equal("1", Text(tab, "FilesCountText"));
            Assert.Equal("none", Text(tab, "CreditsCountText"));

            Assert.Equal(["U1", "U8"], tab.ComponentsShownForTests.Select(item => item.BoardLabel));
            Assert.True(tab.ComponentsShownForTests.Single(item => item.BoardLabel == "U8").IsDrafted);
            Assert.False(tab.ComponentsShownForTests.Single(item => item.BoardLabel == "U1").IsDrafted);
            Assert.Contains("2026-October-4", Text(tab, "PublishedOnText"), StringComparison.Ordinal);
        });
    }

    // The search box narrows the list by its words; one that matches nothing says so.
    [Fact]
    public void The_search_narrows_the_components_and_says_when_nothing_matches()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            tab.ShowBoard(Entry("250407", BoardA), Board(), "PAL", null);

            TextBox search = tab.GetControl<TextBox>("ComponentSearchBox");

            // TextChanged is raised through the dispatcher.
            search.Text = "cpu";
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["U8"], tab.ComponentsShownForTests.Select(item => item.BoardLabel));
            Assert.False(tab.GetControl<TextBlock>("NoComponentsText").IsVisible);

            search.Text = "zilog";
            Dispatcher.UIThread.RunJobs();
            Assert.Empty(tab.ComponentsShownForTests);
            Assert.True(tab.GetControl<TextBlock>("NoComponentsText").IsVisible);
            Assert.Equal("No component matches the search.", Text(tab, "NoComponentsText"));
        });
    }

    // ###########################################################################################
    // With no board shown nothing on the left can be used - but "Add a new board" can: it is what
    // someone does when the board they want is not in the lists at all.
    // ###########################################################################################
    [Fact]
    public void Without_a_board_the_left_column_is_off_and_Add_a_new_board_is_not()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            tab.ShowBoard(null, null, "PAL", null);

            Assert.False(tab.GetControl<Button>("AddComponentButton").IsEffectivelyEnabled);
            Assert.False(tab.GetControl<Button>("TableViewButton").IsEffectivelyEnabled);
            Assert.True(tab.GetControl<Button>("AddNewBoardButton").IsEffectivelyEnabled);
            Assert.Equal("No board shown", Text(tab, "NoDraftBoardText"));
        });
    }

    // Both "add" buttons are green - they go ahead and make something (CLAUDE.md, button colours).
    [Fact]
    public void Add_component_and_Add_a_new_board_are_green()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();

            Assert.Contains("Accept", tab.GetControl<Button>("AddComponentButton").Classes);
            Assert.Contains("Accept", tab.GetControl<Button>("AddNewBoardButton").Classes);
            Assert.DoesNotContain("Accept", tab.GetControl<Button>("TableViewButton").Classes);
        });
    }

    // ###########################################################################################
    // The header: a board with a draft shows the draft's row; one with none says what it is, and
    // that the first change makes the draft.
    // ###########################################################################################
    [Fact]
    public void The_header_is_the_drafts_row_or_says_the_board_has_no_draft_yet()
    {
        WriteDraft(BoardA);

        UiTest.Run(() =>
        {
            HardwareBoardEntry withDraft = Entry("250407", BoardA);
            HardwareBoardEntry withoutDraft = Entry("250425", BoardB);
            TabDrafts tab = Tab(withDraft, withoutDraft);

            tab.ShowBoard(withDraft, Board(), "PAL", null);

            ContentControl header = tab.GetControl<ContentControl>("CurrentDraftRow");
            Assert.True(header.IsVisible);
            Assert.Equal(BoardA, Assert.IsType<DraftListItem>(header.Content).ExcelDataFile);
            Assert.False(tab.GetControl<Border>("NoDraftPanel").IsVisible);

            tab.ShowBoard(withoutDraft, Board(), "PAL", null);

            Assert.False(header.IsVisible);
            Assert.True(tab.GetControl<Border>("NoDraftPanel").IsVisible);
            Assert.Equal("Commodore 64 - 250425", Text(tab, "NoDraftBoardText"));
            Assert.Equal(ContributeBoardParts.NoDraftLine(hasBoard: true), Text(tab, "NoDraftText"));
        });
    }

    // ###########################################################################################
    // "My drafts" shows the board picked - through Main's drop-downs, so all of CRT is on it - and
    // falls back to its placeholder, so it reads as a menu, never as a second board name.
    // ###########################################################################################
    [Fact]
    public void Picking_a_draft_in_My_drafts_shows_its_board_and_the_list_falls_back()
    {
        WriteDraft(BoardA);
        WriteDraft(BoardB);

        UiTest.Run(() =>
        {
            TabDrafts tab = Tab(Entry("250407", BoardA), Entry("250425", BoardB));
            var shown = new List<string>();
            tab.SelectBoardOverrideForTests = entry => shown.Add(entry.ExcelDataFile);

            ComboBox myDrafts = tab.GetControl<ComboBox>("MyDraftsComboBox");
            Assert.Equal("My drafts (2)", myDrafts.PlaceholderText);

            myDrafts.SelectedItem = tab.Drafts.Single(item => item.ExcelDataFile == BoardB);

            Assert.Equal([BoardB], shown);
            Assert.Null(myDrafts.SelectedItem);
        });
    }

    // ###########################################################################################
    // *** A FIRST CHANGE MAKES THE DRAFT. *** On a board with none, an action on the left asks for
    // the draft first (Main.EnsureDraftAsync) and goes on once there is one; when none could be made,
    // it does nothing.
    // ###########################################################################################
    [Fact]
    public async Task An_action_on_a_board_with_no_draft_makes_one_first_and_stops_when_none_is_made()
    {
        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            var asked = new List<string>();
            var acted = new List<string>();

            tab.EnsureDraftOverrideForTests = board =>
            {
                asked.Add(board.ExcelDataFile);
                return Task.FromResult(false);
            };

            await tab.OnCurrentDraftAsync(board => { acted.Add(board.ExcelDataFile); return Task.CompletedTask; });

            Assert.Equal([BoardA], asked);
            Assert.Empty(acted);

            tab.EnsureDraftOverrideForTests = board =>
            {
                asked.Add(board.ExcelDataFile);
                WriteDraft(board.ExcelDataFile);
                return Task.FromResult(true);
            };

            await tab.OnCurrentDraftAsync(board => { acted.Add(board.ExcelDataFile); return Task.CompletedTask; });

            Assert.Equal([BoardA, BoardA], asked);
            Assert.Equal([BoardA], acted);

            // Has one now: nothing more is asked.
            await tab.OnCurrentDraftAsync(board => { acted.Add(board.ExcelDataFile); return Task.CompletedTask; });

            Assert.Equal(2, asked.Count);
            Assert.Equal(2, acted.Count);
            Assert.Single(tab.Drafts);
        });
    }

    // ###########################################################################################
    // The board's own lists open in the pane on the right (slice 4a) - its draft made first, the
    // table closed - and the important signals get the board's KiCad nets once they are read.
    // ###########################################################################################
    [Fact]
    public async Task Credits_and_important_signals_open_in_the_pane_and_the_signals_get_the_nets()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            tab.NetChoicesOverrideForTests = _ => Task.FromResult(KiCadNetNames.Choices([("CLK", "CLK", false)]));
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.ShowTableOnSheetAsync(BoardWorkbookSchema.SheetCredits);
            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);

            Assert.False(tab.IsTableOpen);
            Assert.True(tab.IsComponentOpen);
            Assert.Equal(PaneBoardPart.Credits, tab.ComponentPaneForTests.BoardPart);

            await tab.OpenBoardPartAsync(PaneBoardPart.ImportantSignals);
            await Until(() => tab.ComponentPaneForTests.SignalsForTests.Count == 1 && tab.ComponentPaneForTests.SignalsForTests[0].NetChoices.Count == 1);

            Assert.Equal(["CLK"], tab.ComponentPaneForTests.SignalsForTests[0].NetChoices);
            Assert.False(tab.ComponentPaneForTests.SignalsForTests[0].HasProblem);

            // The pictures and their colours too (slice 4b).
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
            Assert.Equal(PaneBoardPart.Pictures, tab.ComponentPaneForTests.BoardPart);
            Assert.Equal(["Top"], tab.ComponentPaneForTests.BoardPicturesForTests.Select(row => row.Name));
        });
    }

    // ###########################################################################################
    // The Pictures part gets the board's KiCad views to match its pictures to (step 3 - the
    // "Schematic images" window's drop-downs, moved onto the cards), read for it once it is open.
    // ###########################################################################################
    [Fact]
    public async Task The_Pictures_part_gets_the_boards_KiCad_views_for_its_pictures()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            tab.KiCadViewsOverrideForTests = _ => Task.FromResult<IReadOnlyList<string>>(["PCB top", "PCB bottom"]);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow top = tab.ComponentPaneForTests.BoardPicturesForTests.Single();
            await Until(() => top.CadNameChoices.Count == 3);

            Assert.Equal(["", "PCB top", "PCB bottom"], top.CadNameChoices.Select(item => item.CadName));
            Assert.True(top.ShowsCadNames);
        });
    }

    // ###########################################################################################
    // A draft made only because something was opened on the board - a component ("Edit this
    // component" on everybody's right-click menu), a board part, the table - goes again when that is
    // closed with nothing in the draft (code review, 2026-10-26): it stayed, the tab's title turned
    // red, and an empty draft sat in "My drafts". A draft with a change, or one there before, stays.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_made_only_to_look_goes_again_when_closed_with_nothing_in_it()
    {
        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.EnsureDraftOverrideForTests = _ =>
            {
                WriteDraft(BoardA, new BoardData());
                return Task.FromResult(true);
            };
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, BoardA));

            tab.CloseComponent();

            Assert.False(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, BoardA));

            // Made again to look - and this time changed: it stays.
            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;
            Assert.True(session.Apply("Credits", content => content with { Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Board data", NameOrHandle = "Bo" }]) }));
            Assert.Equal(DraftWorkbookEditOutcome.Saved, await session.SaveAsync());

            tab.CloseComponent();

            Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, BoardA));
        });
    }

    // ###########################################################################################
    // Submit reads the board as published for its credit suggestions under CRT's "please wait" (code
    // review, 2026-10-26): a large board not yet read was read with nothing on screen and the window
    // open to clicks, between the "reading your changes" wait and the dialog.
    // ###########################################################################################
    [Fact]
    public async Task The_published_board_for_the_credits_is_read_under_please_wait()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            var overlay = new BusyOverlay();
            var root = new Grid();
            root.Children.Add(tab);
            root.Children.Add(overlay);
            var window = new Window { Content = root, Width = 1300, Height = 900 };
            window.Show();

            bool busyWhileReading = false;
            tab.PublishedBoardOverrideForTests = _ =>
            {
                busyWhileReading = overlay.IsBusy;
                return Board();
            };

            DraftStatus status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, BoardA)!;
            Assert.NotNull(await tab.PublishedBoardForCreditsAsync(entry, status));

            Assert.True(busyWhileReading);
            window.Close();
        });
    }

    // ###########################################################################################
    // A draft made only to look whose pictures were MOVED is kept (code review, 2026-10-28): order is
    // no change the counts see, so closing the part deleted the draft and the new order with it.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_made_only_to_look_keeps_its_pictures_new_order()
    {
        BoardData Pictures() => new()
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" },
                new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "bottom.png" },
            ],
        };

        string dataRoot = Path.Combine(this.thisWorkspace.Root, "Data");
        string published = Path.Combine(dataRoot, DataTreeLayout.BoardFileFor(BoardA).Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(published)!);
        CachedWorkbooks.Write(published, Pictures());
        DataManager.LoadFrom(dataRoot, "does-not-exist.xlsx");

        try
        {
            await UiTest.RunAsync(async () =>
            {
                HardwareBoardEntry entry = Entry("250407", BoardA);
                TabDrafts tab = Tab(entry);
                tab.PublishedBoardOverrideForTests = _ => Pictures();
                tab.EnsureDraftOverrideForTests = _ =>
                {
                    WriteDraft(BoardA, Pictures());
                    return Task.FromResult(true);
                };
                tab.ShowBoard(entry, Pictures(), "PAL", null);

                await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
                tab.ComponentPaneForTests.MovePicture(tab.ComponentPaneForTests.BoardPicturesForTests[1], -1);
                DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;
                Assert.Equal(DraftWorkbookEditOutcome.Saved, await session.SaveAsync());

                tab.CloseComponent();

                Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, BoardA));
            });
        }
        finally
        {
            DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        }
    }

    // ###########################################################################################
    // A picture's name the Pictures part cannot take keeps the part open, the reason under its
    // heading, when another part or the table is asked for (code review, 2026-10-28: the typed name
    // and its reason went without a word). Once the name is put back, the other part opens.
    // ###########################################################################################
    [Fact]
    public async Task A_picture_name_that_cannot_be_taken_keeps_the_Pictures_part_open()
    {
        BoardData TwoPictures()
        {
            BoardData board = Board();
            board.Schematics.Add(new BoardSchematicEntry { SchematicName = "Bottom" });
            return board;
        }

        WriteDraft(BoardA, TwoPictures());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => TwoPictures();
            tab.ShowBoard(entry, TwoPictures(), "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
            ComponentPaneBoardPictureRow bottom = tab.ComponentPaneForTests.BoardPicturesForTests[1];
            bottom.NameDraft = "Top";

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Assert.Equal(PaneBoardPart.Pictures, tab.ComponentPaneForTests.BoardPart);
            Assert.Equal(SchematicRename.NameTaken("Top"), tab.ComponentPaneForTests.PicturesNoteForTests);

            await tab.ShowTableOnSheetAsync(BoardWorkbookSchema.SheetCredits);
            Assert.False(tab.IsTableOpen);

            bottom.NameDraft = "Bottom";
            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            Assert.Equal(PaneBoardPart.Credits, tab.ComponentPaneForTests.BoardPart);
        });
    }

    // A draft that was there before is never let go of, however little it holds.
    [Fact]
    public async Task A_draft_there_before_stays_whatever_it_holds()
    {
        WriteDraft(BoardA, new BoardData());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);
            tab.CloseComponent();

            Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, BoardA));
        });
    }

    // ###########################################################################################
    // A picture's preview marks the components of the region CRT shows (code review, 2026-10-26): the
    // part used the pane's region, which only opening a component set - empty first, hiding every
    // regional component, or the region of the last component opened.
    // ###########################################################################################
    [Fact]
    public async Task A_pictures_preview_marks_the_components_of_the_region_shown()
    {
        BoardData Marked()
        {
            BoardData board = Board();
            board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" });
            board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U2", X = "20", Y = "20", Width = "5", Height = "5" });
            return board;
        }

        WriteDraft(BoardA, Marked());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Marked();
            tab.ShowBoard(entry, Marked(), "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            // U1 is PAL only, U2 NTSC only: the PAL one alone.
            Assert.Equal([new Avalonia.Rect(1, 1, 5, 5)], tab.ComponentPaneForTests.BoardPicturesForTests.Single().Highlights);
        });
    }

    // ###########################################################################################
    // "Mark components" on a picture (slice 6): handed to Main as "Show on board" is - with the board on
    // screen read again first when the pane's saves left it behind, and that reload handed over once.
    // ###########################################################################################
    [Fact]
    public async Task A_picture_tool_goes_to_the_Schematics_tab_on_the_board_as_the_draft_now_is()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            tab.ReloadBoardOverrideForTests = _ => { };
            tab.ShowBoard(entry, Board(), "PAL", null);

            var started = new List<(PictureTool Tool, string Picture, bool ReadBoardFirst)>();
            tab.PictureToolOverrideForTests = (tool, picture, readBoardFirst) => started.Add((tool, picture, readBoardFirst));

            // On screen, as in CRT: a save of the pane's then leaves the board on screen behind.
            var window = new Window { Content = tab, Width = 1300, Height = 900 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
            await tab.StartPictureToolAsync(PictureTool.MarkComponents, "Top");

            // The board on screen is the draft as shown: nothing to read again.
            Assert.Equal([(PictureTool.MarkComponents, "Top", false)], started);

            // A change the pane saves leaves the board on screen behind: read it again first, once.
            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "CPU 6510";
            await tab.StartPictureToolAsync(PictureTool.AlignKiCad, "Top");
            await tab.StartPictureToolAsync(PictureTool.AlignKiCad, "Top");

            Assert.Equal([(PictureTool.AlignKiCad, "Top", true), (PictureTool.AlignKiCad, "Top", false)], started.Skip(1));

            window.Close();
        });
    }

    // ###########################################################################################
    // The credits ticked in the Submit dialog (2026-10-26) go into the draft as ONE step and are
    // saved - and the board then handed back, the one a submit sends, carries them, with the draft's
    // fingerprint as it now is.
    // ###########################################################################################
    [Fact]
    public async Task Credits_from_the_Submit_dialog_go_into_the_draft_and_the_board_sent_carries_them()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            DraftStatus status = DraftStatusReader.Resolve(this.thisWorkspace.Root, DraftManager.DraftsRoot, BoardA)!;
            string before = DraftFingerprint.Compute(status.WorkbookPath, DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, BoardA));

            (BoardData Board, string Fingerprint)? added = await tab.AddCreditsToDraftAsync(
                entry,
                status,
                [new CreditEntry { Category = "Board data", NameOrHandle = "Bo", Contact = "RHD" }]);

            Assert.NotNull(added);
            Assert.Contains(added!.Value.Board.Credits, credit => credit.NameOrHandle == "Bo" && credit.Category == "Board data");
            Assert.NotEqual(before, added.Value.Fingerprint);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;
            Assert.Equal("Credits", session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
        });
    }

    // ###########################################################################################
    // "Copy from another board" (slice 5) looks only at the OTHER boards - never the one the pane is
    // open on - and offers their components with the same part number.
    // ###########################################################################################
    [Fact]
    public async Task Copy_from_another_board_looks_at_the_other_boards_only()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            HardwareBoardEntry other = Entry("250425", BoardB);
            TabDrafts tab = Tab(entry, other);

            // Both boards have a 6510 - only the other one's may be offered.
            BoardData WithCpu(string label) => new()
            {
                Components = [new ComponentEntry { BoardLabel = label, FriendlyName = "CPU", PartNumber = "6510", Category = "IC", Package = "DIP-40" }],
            };

            tab.PublishedBoardOverrideForTests = board => board.ExcelDataFile == BoardA ? WithCpu("U8") : WithCpu("U7");
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("PartNumberBox").Text = "6510";
            await tab.ComponentPaneForTests.OpenCopyPanelAsync();

            ComponentCopySource offered = Assert.Single(tab.ComponentPaneForTests.GetControl<ComboBox>("CopySourceBox").Items.OfType<ComponentCopySource>());
            Assert.Equal(other.ToString(), offered.BoardName);
            Assert.Equal("U7", offered.Row.BoardLabel);
        });
    }

    // ###########################################################################################
    // "Also on other boards" end to end (slice 5): what this draft changed of the part goes into the
    // OTHER board's own draft as one step named after this board, and is saved; a board of which no
    // draft can be made is said, not changed.
    // ###########################################################################################
    [Fact]
    public async Task Also_on_other_boards_changes_the_part_in_the_other_boards_own_draft()
    {
        BoardData Cpu(string label) => new()
        {
            Components = [new ComponentEntry { BoardLabel = label, FriendlyName = "CPU", PartNumber = "6510", Category = "IC" }],
        };

        WriteDraft(BoardA, Cpu("U8"));
        WriteDraft(BoardB, Cpu("U7"));

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            HardwareBoardEntry other = Entry("250425", BoardB);
            HardwareBoardEntry noDraft = Entry("250469", "Commodore/C64/" + Guid.NewGuid().ToString("N") + "/board.json");
            TabDrafts tab = Tab(entry, other, noDraft);
            tab.PublishedBoardOverrideForTests = board => board.ExcelDataFile == BoardA ? Cpu("U8") : board.ExcelDataFile == BoardB ? Cpu("U7") : Cpu("U1");
            tab.ShowBoard(entry, Cpu("U8"), "PAL", null);

            await tab.OpenComponentAsync("U8");
            ComponentPane pane = tab.ComponentPaneForTests;
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-40";
            pane.CommitHeld();

            await pane.OpenAlsoPanelAsync();
            await pane.ChangeOtherBoardsAsync();

            DraftSession otherDraft = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardB)!;
            Assert.Equal("DIP-40", otherDraft.Content.Board.Components.Single().Package);
            Assert.Equal($"[U7] copied from {entry}", otherDraft.UndoDescription);
            Assert.False(otherDraft.HasUnsavedChanges);

            string said = pane.GetControl<TextBlock>("PaneMessageText").Text ?? string.Empty;
            Assert.Contains($"Changed on 1 board, in its own draft: {other}.", said, StringComparison.Ordinal);
            Assert.Contains($"Not changed on {noDraft} - no draft could be made of it.", said, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // An other board's draft that cannot be saved is said so - and is left as it was (code review,
    // 2026-10-26): the step stayed in its draft held in memory, its header said "Could not save" and
    // the timer that saves drafts no editor holds could save it later, after CRT had said it had not.
    // ###########################################################################################
    [Fact]
    public async Task Also_on_other_boards_takes_its_change_back_out_of_a_draft_that_cannot_be_saved()
    {
        BoardData Cpu(string label) => new()
        {
            Components = [new ComponentEntry { BoardLabel = label, FriendlyName = "CPU", PartNumber = "6510", Category = "IC" }],
        };

        WriteDraft(BoardA, Cpu("U8"));
        WriteDraft(BoardB, Cpu("U7"));

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            HardwareBoardEntry other = Entry("250425", BoardB);
            TabDrafts tab = Tab(entry, other);
            tab.PublishedBoardOverrideForTests = board => board.ExcelDataFile == BoardA ? Cpu("U8") : Cpu("U7");
            tab.ShowBoard(entry, Cpu("U8"), "PAL", null);

            DraftSession otherDraft = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardB)!;
            otherDraft.WhileWritingForTests = () => throw new IOException("held open by another program");

            await tab.OpenComponentAsync("U8");
            ComponentPane pane = tab.ComponentPaneForTests;
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-40";
            pane.CommitHeld();

            await pane.OpenAlsoPanelAsync();
            await pane.ChangeOtherBoardsAsync();

            Assert.Contains($"Not changed on {other} - its draft could not be saved.", pane.GetControl<TextBlock>("PaneMessageText").Text ?? string.Empty, StringComparison.Ordinal);
            Assert.True(string.IsNullOrEmpty(otherDraft.Content.Board.Components.Single().Package));
            Assert.False(otherDraft.HasUnsavedChanges);
            Assert.False(otherDraft.CanUndo);

            otherDraft.WhileWritingForTests = null;
        });
    }

    // ###########################################################################################
    // A component opened - or just added - has its pins read from its board's KiCad data and handed
    // to the pane, which then offers "Fill from KiCad" (the plan: "On a KiCad board, Fill from KiCad
    // adds each pin's net and its name from the symbol").
    // ###########################################################################################
    [Fact]
    public async Task A_component_shown_gets_its_pins_from_the_boards_KiCad_data()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();

            var asked = new List<string>();
            tab.KiCadPinsOverrideForTests = (identity, label) =>
            {
                asked.Add($"{identity}|{label}");
                return Task.FromResult<IReadOnlyList<KiCadPin>>([new KiCadPin("1", "VSS", "Ground", "GND")]);
            };

            tab.ShowBoard(entry, Board(), "PAL", null);
            await tab.OpenComponentAsync("U8");
            Dispatcher.UIThread.RunJobs();

            Assert.Equal([$"{BoardA}|U8"], asked);
            Assert.Single(tab.ComponentPaneForTests.KiCadPinsForTests);
            Assert.True(tab.ComponentPaneForTests.GetControl<Button>("FillFromKiCadButton").IsVisible);

            // A component added is shown too - and its pins asked for.
            await tab.AddComponentAsync();
            tab.ComponentPaneForTests.GetControl<TextBox>("NewLabelBox").Text = "U12";
            tab.ComponentPaneForTests.GetControl<AutoCompleteBox>("NewCategoryBox").Text = "IC";
            Assert.True(tab.ComponentPaneForTests.AddNewComponent());
            Dispatcher.UIThread.RunJobs();

            Assert.Equal([$"{BoardA}|U8", $"{BoardA}|U12"], asked);
        });
    }

    // ###########################################################################################
    // "Show on board" on an important signal: the pane's net handed to Main - with the board read
    // again first when the pane's saves left the board on screen behind, as for a component.
    // ###########################################################################################
    [Fact]
    public async Task Show_on_board_on_a_signal_hands_its_net_to_Main()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            tab.NetChoicesOverrideForTests = _ => Task.FromResult(KiCadNetNames.Choices([("CLK", "CLK", false)]));

            var shown = new List<(string Net, bool ReadBoardFirst)>();
            tab.ShowNetOverrideForTests = (net, readBoardFirst) => shown.Add((net, readBoardFirst));

            tab.ShowBoard(entry, Board(), "PAL", null);
            await tab.OpenBoardPartAsync(PaneBoardPart.ImportantSignals);
            Dispatcher.UIThread.RunJobs();

            tab.ComponentPaneForTests.ShowNetOnBoard("CLK");
            Dispatcher.UIThread.RunJobs();
            await tab.ComponentPaneForTests.LastSave;
            Dispatcher.UIThread.RunJobs();

            Assert.Equal([("CLK", false)], shown);
        });
    }

    // ###########################################################################################
    // EVERY component of a board with the part number is changed - each one listed is ticked (code
    // review, 2026-10-26: the boards were taken once each, and only a board's FIRST 4164 got the pins
    // while the pane said the board was changed).
    // ###########################################################################################
    [Fact]
    public async Task Also_on_other_boards_changes_every_ticked_component_of_a_board()
    {
        BoardData Rams(params string[] labels) => new()
        {
            Components = labels.Select(label => new ComponentEntry { BoardLabel = label, FriendlyName = "RAM", PartNumber = "4164", Category = "IC" }).ToList(),
        };

        WriteDraft(BoardA, Rams("U9"));
        WriteDraft(BoardB, Rams("U9", "U10"));

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            HardwareBoardEntry other = Entry("250425", BoardB);
            TabDrafts tab = Tab(entry, other);
            tab.PublishedBoardOverrideForTests = board => board.ExcelDataFile == BoardA ? Rams("U9") : Rams("U9", "U10");
            tab.ShowBoard(entry, Rams("U9"), "PAL", null);

            await tab.OpenComponentAsync("U9");
            ComponentPane pane = tab.ComponentPaneForTests;
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-16";
            pane.CommitHeld();

            await pane.OpenAlsoPanelAsync();
            Assert.Equal(2, pane.GetControl<StackPanel>("AlsoBoardsPanel").Children.OfType<CheckBox>().Count());
            await pane.ChangeOtherBoardsAsync();

            DraftSession otherDraft = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardB)!;
            Assert.Equal(["DIP-16", "DIP-16"], otherDraft.Content.Board.Components.Select(component => component.Package));
        });
    }

    // ###########################################################################################
    // A board given a file but whose part could not be changed says so (code review, 2026-10-26: the
    // file going in listed the board as "Changed", and the part's failed save was never said).
    // ###########################################################################################
    [Fact]
    public async Task A_part_not_changed_on_a_board_that_took_the_file_is_said()
    {
        byte[] old = [37, 80, 68, 70, 49];
        byte[] mine = [37, 80, 68, 70, 50];

        BoardData Cpu(string label) => new()
        {
            Components = [new ComponentEntry { BoardLabel = label, FriendlyName = "CPU", PartNumber = "6510", Category = "IC" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = label, Name = "Datasheet", File = "Component local files/6510.pdf" }],
        };

        WriteDraftWithFiles(BoardA, Cpu("U8"), ("Component local files/6510.pdf", old));
        WriteDraftWithFiles(BoardB, Cpu("U7"), ("Component local files/6510.pdf", old));

        var published = new Dictionary<string, (BoardData Board, string File)>(StringComparer.OrdinalIgnoreCase)
        {
            [BoardA] = (Cpu("U8"), this.PublishedFile(Cpu("U8"), ("Component local files/6510.pdf", old))),
            [BoardB] = (Cpu("U7"), this.PublishedFile(Cpu("U7"), ("Component local files/6510.pdf", old))),
        };

        string picked = Path.Combine(this.thisWorkspace.Root, "Picked", "6510 rev B.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(picked)!);
        File.WriteAllBytes(picked, mine);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry a = Entry("250407", BoardA);
            HardwareBoardEntry b = Entry("250425", BoardB);
            TabDrafts tab = Tab(a, b);
            tab.PublishedBoardOverrideForTests = entry => published[entry.ExcelDataFile].Board;
            tab.PublishedBoardFileOverrideForTests = entry => published[entry.ExcelDataFile].File;
            tab.ShowBoard(a, published[BoardA].Board, "PAL", null);

            await tab.OpenComponentAsync("U8");
            ComponentPane pane = tab.ComponentPaneForTests;
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-40";
            pane.CommitHeld();
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);
            await pane.ChangeFileAsync(pane.FilesForTests[0]);

            await pane.OpenAlsoPanelAsync();

            // The part's save fails once; the file's, after it, is made.
            DraftSession bDraft = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardB)!;
            int writes = 0;
            bDraft.WhileWritingForTests = () =>
            {
                if (writes++ == 0)
                {
                    throw new IOException("held open by another program");
                }
            };

            await pane.ChangeOtherBoardsAsync();
            bDraft.WhileWritingForTests = null;

            Assert.True(string.IsNullOrEmpty(bDraft.Content.Board.Components.Single().Package));
            string newName = Assert.Single(bDraft.Content.Board.ComponentLocalFiles).File;
            Assert.Equal(Sha(mine), bDraft.ListedFiles[newName].Sha256);

            string said = pane.GetControl<TextBlock>("PaneMessageText").Text ?? string.Empty;
            Assert.Contains($"Changed on 1 board, in its own draft: {b}.", said, StringComparison.Ordinal);
            Assert.Contains($"[U7] not changed on {b} - its draft could not be saved.", said, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // The component's KiCad pins are read again when the draft's KiCad data changes while it is open
    // (code review, 2026-10-26: a KiCad project imported with the component open left "Fill from
    // KiCad" hidden and the nets empty until it was opened again).
    // ###########################################################################################
    [Fact]
    public async Task The_KiCad_pins_are_read_again_when_the_drafts_KiCad_data_changes()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();

            int asked = 0;
            tab.KiCadPinsOverrideForTests = (_, _) =>
            {
                asked++;
                return Task.FromResult<IReadOnlyList<KiCadPin>>([]);
            };

            tab.ShowBoard(entry, Board(), "PAL", null);
            await tab.OpenComponentAsync("U8");
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(1, asked);

            // A KiCad project imported into the draft: copied in, then taken into its list.
            string folder = DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, BoardA);
            Directory.CreateDirectory(Path.Combine(folder, "KiCad data"));
            File.WriteAllText(Path.Combine(folder, "KiCad data", "board.kicad_pcb"), "(kicad_pcb)");
            Assert.True(await DraftEdits.TakeInNewFilesAsync(BoardA));

            Dispatcher.UIThread.RunJobs();
            Assert.Equal(2, asked);
        });
    }

    // ###########################################################################################
    // A draft made for a board where nothing then went in is let go of again (code review, 2026-10-26:
    // every such board was left with an empty draft in "My drafts", to be discarded by hand).
    // ###########################################################################################
    [Fact]
    public async Task A_draft_made_for_a_board_that_took_nothing_is_let_go_of()
    {
        BoardData Cpu(string label, string package) => new()
        {
            Components = [new ComponentEntry { BoardLabel = label, FriendlyName = "CPU", PartNumber = "6510", Category = "IC", Package = package }],
        };

        WriteDraft(BoardA, Cpu("U8", ""));

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            HardwareBoardEntry other = Entry("250425", BoardB);
            TabDrafts tab = Tab(entry, other);

            // B has the package already - and no draft, so one is made for it.
            tab.PublishedBoardOverrideForTests = board => board.ExcelDataFile == BoardA ? Cpu("U8", "") : Cpu("U7", "DIP-40");
            tab.SeedDraftOverrideForTests = identity =>
            {
                WriteDraft(identity, Cpu("U7", "DIP-40"));
                return true;
            };

            tab.ShowBoard(entry, Cpu("U8", ""), "PAL", null);
            await tab.OpenComponentAsync("U8");
            ComponentPane pane = tab.ComponentPaneForTests;
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-40";
            pane.CommitHeld();

            await pane.OpenAlsoPanelAsync();
            await pane.ChangeOtherBoardsAsync();

            Assert.Contains($"Not changed on {other} - it is like that there already.", pane.GetControl<TextBlock>("PaneMessageText").Text ?? string.Empty, StringComparison.Ordinal);
            Assert.False(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, BoardB));
        });
    }

    // ###########################################################################################
    // A file put on another board is ONE step of that board's draft, which its Undo takes back (code
    // review, 2026-10-26: written straight into the draft's file list, it was no step at all - and a
    // board file had no other way back but discarding the whole draft).
    // ###########################################################################################
    [Fact]
    public async Task A_file_put_on_another_board_is_one_step_that_boards_Undo_takes_back()
    {
        byte[] old = [37, 80, 68, 70, 49];
        byte[] mine = [37, 80, 68, 70, 50];

        BoardData WithManual() => new()
        {
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf" }],
        };

        WriteDraftWithFiles(BoardA, WithManual(), ("Board local files/manual.pdf", old));
        WriteDraftWithFiles(BoardB, WithManual(), ("Board local files/manual.pdf", old));

        var published = new Dictionary<string, (BoardData Board, string File)>(StringComparer.OrdinalIgnoreCase)
        {
            [BoardA] = (WithManual(), this.PublishedFile(WithManual(), ("Board local files/manual.pdf", old))),
            [BoardB] = (WithManual(), this.PublishedFile(WithManual(), ("Board local files/manual.pdf", old))),
        };

        string picked = Path.Combine(this.thisWorkspace.Root, "Picked", "manual 2nd edition.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(picked)!);
        File.WriteAllBytes(picked, mine);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry a = Entry("250407", BoardA);
            HardwareBoardEntry b = Entry("250425", BoardB);
            TabDrafts tab = Tab(a, b);
            tab.PublishedBoardOverrideForTests = entry => published[entry.ExcelDataFile].Board;
            tab.PublishedBoardFileOverrideForTests = entry => published[entry.ExcelDataFile].File;
            tab.ShowBoard(a, published[BoardA].Board, "PAL", null);

            await tab.OpenBoardPartAsync(PaneBoardPart.FilesAndLinks);
            ComponentPane pane = tab.ComponentPaneForTests;
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);
            await pane.ChangeBoardFileAsync(pane.BoardFilesForTests.Single());

            await pane.OpenBoardAlsoPanelAsync();
            await pane.ChangeBoardFilesOnOtherBoardsAsync();

            DraftSession bDraft = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardB)!;
            string newName = Assert.Single(bDraft.Content.Board.BoardLocalFiles).File;
            Assert.Equal(Sha(mine), bDraft.ListedFiles[newName].Sha256);

            Assert.True(bDraft.CanUndo);
            Assert.True(bDraft.Undo());
            Assert.Equal("Board local files/manual.pdf", Assert.Single(bDraft.Content.Board.BoardLocalFiles).File);
        });
    }

    // ###########################################################################################
    // The pins a LATER read found are kept - a read started before a KiCad import that ends after the
    // read started by it never puts the old pins back (code review, 2026-10-26).
    // ###########################################################################################
    [Fact]
    public async Task An_older_KiCad_pin_read_ending_last_does_not_win()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();

            var reads = new List<TaskCompletionSource<IReadOnlyList<KiCadPin>>>();
            tab.KiCadPinsOverrideForTests = (_, _) =>
            {
                var read = new TaskCompletionSource<IReadOnlyList<KiCadPin>>();
                reads.Add(read);
                return read.Task;
            };

            tab.ShowBoard(entry, Board(), "PAL", null);
            await tab.OpenComponentAsync("U8");

            string folder = DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, BoardA);
            Directory.CreateDirectory(Path.Combine(folder, "KiCad data"));
            File.WriteAllText(Path.Combine(folder, "KiCad data", "board.kicad_pcb"), "(kicad_pcb)");
            Assert.True(await DraftEdits.TakeInNewFilesAsync(BoardA));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(2, reads.Count);

            reads[1].SetResult([new KiCadPin("1", "VSS", "Ground", "GND")]);
            Dispatcher.UIThread.RunJobs();
            reads[0].SetResult([new KiCadPin("1", "OLD", "", "")]);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("VSS", Assert.Single(tab.ComponentPaneForTests.KiCadPinsForTests).Name);
        });
    }

    // ---- also on other boards, for files ----------------------------------------------------------

    private static string Sha(byte[] bytes) => Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes));

    // A draft of `board` listing `files`, their bytes in the draft's own store.
    private static void WriteDraftWithFiles(string excelDataFile, BoardData board, params (string Name, byte[] Bytes)[] files)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        string folder = Path.GetDirectoryName(boardFile)!;
        Directory.CreateDirectory(folder);

        var listed = new List<BoardDocument.StoredFileRow>();

        foreach ((string name, byte[] bytes) in files)
        {
            string stored = Path.Combine(folder, DataTreeLayout.StoredFilePath(Sha(bytes), name).Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(stored)!);
            File.WriteAllBytes(stored, bytes);
            listed.Add(new BoardDocument.StoredFileRow { Name = name, Sha256 = Sha(bytes), Size = bytes.Length });
        }

        BoardFiles.Write(boardFile, BoardDocumentMapping.ToDocument(new BoardContent(board, [], []), name => name, listed));
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile), new DraftMarker { BoardKey = excelDataFile });
    }

    // The published board file of `board`, listing `files` - in a folder of the test's own.
    private string PublishedFile(BoardData board, params (string Name, byte[] Bytes)[] files)
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Published", Guid.NewGuid().ToString("N"), DataTreeLayout.BoardFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        BoardFiles.Write(path, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            files.Select(file => new BoardDocument.StoredFileRow { Name = file.Name, Sha256 = Sha(file.Bytes), Size = file.Bytes.Length }).ToList()));

        return path;
    }

    // ###########################################################################################
    // "Also on other boards" for a file, end to end (the plan: "This datasheet is also used by 11
    // other boards - change it there too?"): a datasheet replaced on this board goes into the draft of
    // each other board that has the same file, under THAT board's name for it - never over a version
    // that board's draft has of its own, and never onto bytes another board uses as a scope baseline.
    // ###########################################################################################
    [Fact]
    public async Task A_datasheet_replaced_goes_into_the_other_boards_drafts_that_have_the_same_file()
    {
        const string BoardC = "Commodore/C64/250469/board.json";
        const string BoardD = "Commodore/C64/326298/board.json";

        byte[] old = [37, 80, 68, 70, 49];
        byte[] mine = [37, 80, 68, 70, 50];
        byte[] theirs = [37, 80, 68, 70, 51];

        BoardData WithDatasheet(string label, string file) => new()
        {
            Components = [new ComponentEntry { BoardLabel = label, FriendlyName = "CPU", Category = "IC" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = label, Name = "Datasheet", File = file }],
        };

        var asBaseline = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U1", Pin = "1", File = "Scope baseline/odd.pdf", TimeDiv = "1us" }],
        };

        WriteDraftWithFiles(BoardA, WithDatasheet("U8", "Component local files/6510.pdf"), ("Component local files/6510.pdf", old));
        WriteDraftWithFiles(BoardB, WithDatasheet("U7", "Component local files/6510.pdf"), ("Component local files/6510.pdf", old));
        WriteDraftWithFiles(BoardC, WithDatasheet("U2", "Component local files/MOS 6510.pdf"), ("Component local files/MOS 6510.pdf", theirs));

        var published = new Dictionary<string, (BoardData Board, string File)>(StringComparer.OrdinalIgnoreCase)
        {
            [BoardA] = (WithDatasheet("U8", "Component local files/6510.pdf"), this.PublishedFile(WithDatasheet("U8", "Component local files/6510.pdf"), ("Component local files/6510.pdf", old))),
            [BoardB] = (WithDatasheet("U7", "Component local files/6510.pdf"), this.PublishedFile(WithDatasheet("U7", "Component local files/6510.pdf"), ("Component local files/6510.pdf", old))),
            [BoardC] = (WithDatasheet("U2", "Component local files/MOS 6510.pdf"), this.PublishedFile(WithDatasheet("U2", "Component local files/MOS 6510.pdf"), ("Component local files/MOS 6510.pdf", old))),
            [BoardD] = (asBaseline, this.PublishedFile(asBaseline, ("Scope baseline/odd.pdf", old))),
        };

        string picked = Path.Combine(this.thisWorkspace.Root, "Picked", "6510 rev B.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(picked)!);
        File.WriteAllBytes(picked, mine);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry a = Entry("250407", BoardA);
            HardwareBoardEntry b = Entry("250425", BoardB);
            HardwareBoardEntry c = Entry("250469", BoardC);
            HardwareBoardEntry d = Entry("326298", BoardD);
            TabDrafts tab = Tab(a, b, c, d);
            tab.PublishedBoardOverrideForTests = entry => published[entry.ExcelDataFile].Board;
            tab.PublishedBoardFileOverrideForTests = entry => published[entry.ExcelDataFile].File;
            tab.ShowBoard(a, published[BoardA].Board, "PAL", null);

            await tab.OpenComponentAsync("U8");
            ComponentPane pane = tab.ComponentPaneForTests;
            pane.PickFileOverrideForTests = _ => Task.FromResult<string?>(picked);
            await pane.ChangeFileAsync(pane.FilesForTests[0]);

            await pane.OpenAlsoPanelAsync();

            Assert.Equal(
                [b.ToString(), $"{c} - as MOS 6510.pdf"],
                pane.GetControl<StackPanel>("AlsoFilesPanel").Children.OfType<CheckBox>().Select(box => box.Content as string));

            await pane.ChangeOtherBoardsAsync();

            // B has the new file under a name of its own, its row pointed at it as ONE step named after
            // this board, saved; C keeps its own version.
            DraftSession bDraft = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardB)!;
            Assert.Equal("Component local files/6510 rev B.pdf", Assert.Single(ComponentPaneEdits.FilesOf(bDraft.Content.Board, "U7")).File);
            Assert.Equal(Sha(mine), bDraft.ListedFiles["Component local files/6510 rev B.pdf"].Sha256);
            Assert.Equal($"[Datasheet] copied from {a}", bDraft.UndoDescription);
            Assert.False(bDraft.HasUnsavedChanges);
            Assert.Equal(
                Sha(mine),
                BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardB))!.Files.Single().Sha256);

            DraftSession cDraft = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardC)!;
            Assert.Equal(Sha(theirs), cDraft.ListedFiles["Component local files/MOS 6510.pdf"].Sha256);

            string said = pane.GetControl<TextBlock>("PaneMessageText").Text ?? string.Empty;
            Assert.Contains($"Changed on 1 board, in its own draft: {b}.", said, StringComparison.Ordinal);
            Assert.Contains($"[Datasheet (6510.pdf)] not put on {c} - that board's draft has a version of its own.", said, StringComparison.Ordinal);
            Assert.DoesNotContain(d.ToString(), said, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // Several drafts in one submit (slice 5): the Submit dialog is offered every OTHER draft, and one
    // ticked is made ready as its own Submit would - a change held only in memory saved into its file
    // first, the credit typed put in for its own work - while one with errors is not sent and says so.
    // ###########################################################################################
    [Fact]
    public async Task Another_draft_is_offered_and_made_ready_to_send_with_the_credit_typed()
    {
        const string BoardC = "Commodore/C64/250469/board.json";

        WriteDraft(BoardA, Board());
        WriteDraft(BoardB, new BoardData { Components = [new ComponentEntry { BoardLabel = "U7", Category = "IC", FriendlyName = "CPU" }] });
        WriteDraft(BoardC, new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "" }],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC", FriendlyName = "CPU" }],
        });

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry a = Entry("250407", BoardA);
            HardwareBoardEntry b = Entry("250425", BoardB);
            HardwareBoardEntry c = Entry("250469", BoardC);
            TabDrafts tab = Tab(a, b, c);

            IReadOnlyList<OtherDraftOffer> offers = tab.OtherDraftOffers(a);
            Assert.Equal([BoardB, BoardC], offers.Select(offer => offer.Entry.ExcelDataFile));
            Assert.Null(offers[0].NotOffered);
            Assert.EndsWith("to fix first", offers[1].NotOffered, StringComparison.Ordinal);

            // A change held only in memory - as "Also on other boards" leaves one while its save runs.
            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardB)!;
            Assert.True(session.Apply("Links", content => content with { Board = BoardPartEdits.WithBoardLinks(content.Board, [new BoardLinkEntry { Name = "Forum", Url = "https://example.org" }]) }));
            Assert.True(session.HasUnsavedChanges);

            OtherDraftReady ready = await tab.PrepareOtherDraftAsync(b, "Bo", "RHD");

            SubmitDraftTarget target = Assert.IsType<SubmitDraftTarget>(ready.Target);
            Assert.Equal(b.ToString(), target.DisplayName);
            Assert.Equal("Commodore/C64/250425", target.Identity.BoardId);
            Assert.Equal("Forum", Assert.Single(target.Board.BoardLinks).Name);
            Assert.Contains(target.Board.Credits, credit => credit.NameOrHandle == "Bo" && credit.Contact == "RHD");
            Assert.True(target.Credited);
            Assert.False(session.HasUnsavedChanges);
            Assert.NotEmpty(target.Fingerprint);

            // Its errors cannot be shown from inside the dialog, so it is not sent.
            OtherDraftReady withErrors = await tab.PrepareOtherDraftAsync(c, "Bo", "RHD");
            Assert.Null(withErrors.Target);
            Assert.EndsWith("to fix first", withErrors.Reason, StringComparison.Ordinal);

            // No name typed: nothing more is credited.
            OtherDraftReady uncredited = await tab.PrepareOtherDraftAsync(b, " ", "");
            Assert.False(Assert.IsType<SubmitDraftTarget>(uncredited.Target).Credited);
            Assert.Equal(target.Board.Credits.Count, uncredited.Target.Board.Credits.Count);
        });
    }

    // The table view opens on a sheet asked for (what the board parts did until slice 4a).
    [Fact]
    public async Task Credits_open_the_table_view_on_its_sheet()
    {
        WriteDraft(BoardA);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.ShowTableOnSheetAsync(BoardWorkbookSchema.SheetCredits);

            Assert.True(tab.IsTableOpen);
            Assert.Equal(BoardWorkbookSchema.SheetCredits, tab.TableEditor.CurrentSheet!.Name);
            Assert.Equal("Close table view", tab.GetControl<Button>("TableViewButton").Content);

            await tab.ToggleTableViewAsync();

            Assert.False(tab.IsTableOpen);
            Assert.Equal("Table view", tab.GetControl<Button>("TableViewButton").Content);
        });
    }

    // ###########################################################################################
    // *** THE TABLE IS A VIEW OF THE BOARD SHOWN. *** Another board shown in the drop-downs closes a
    // table open on the first board's draft - what it holds saved first.
    // ###########################################################################################
    [Fact]
    public async Task Showing_another_board_closes_a_table_open_on_the_first()
    {
        WriteDraft(BoardA);
        WriteDraft(BoardB);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry first = Entry("250407", BoardA);
            HardwareBoardEntry second = Entry("250425", BoardB);
            TabDrafts tab = Tab(first, second);
            tab.ShowBoard(first, Board(), "PAL", null);

            await tab.ToggleTableViewAsync();
            Assert.True(tab.IsTableOpen);

            tab.ShowBoard(second, Board(), "PAL", null);

            for (int attempt = 0; attempt < 50 && tab.IsTableOpen; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.False(tab.IsTableOpen);
            Assert.Equal(BoardB, Assert.IsType<DraftListItem>(tab.GetControl<ContentControl>("CurrentDraftRow").Content).ExcelDataFile);
        });
    }

    // Until something is true or half a second has passed - for what a click starts and does not wait for.
    private static async Task Until(Func<bool> condition)
    {
        for (int attempt = 0; attempt < 50 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }
    }

    // ###########################################################################################
    // Clicking a component opens it in the component pane on the right, in place of the start panel
    // (the plan: "A component opens in a pane, never in a maximized window").
    // ###########################################################################################
    [Fact]
    public async Task Clicking_a_component_opens_it_in_the_pane()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            var window = new Window { Content = tab, Width = 1300, Height = 900 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Button u8 = tab.GetControl<ItemsControl>("ComponentsItemsControl")
                .GetRealizedContainers()
                .Select(container => Avalonia.VisualTree.VisualExtensions.FindDescendantOfType<Button>(container))
                .First(button => button?.Tag is ComponentListItem { BoardLabel: "U8" })!;

            u8.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            await Until(() => tab.IsComponentOpen);

            Assert.True(tab.IsComponentOpen);
            Assert.Equal("U8", tab.ComponentPaneForTests.Label);
            Assert.True(tab.ComponentPaneForTests.IsVisible);
            Assert.False(tab.GetControl<StackPanel>("StartPanel").IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // The right side shows one thing: the table opened closes the component pane, a component
    // opened closes the table - and what the pane held is in the draft the table shows.
    // ###########################################################################################
    [Fact]
    public async Task The_table_and_the_component_pane_take_turns_on_the_right()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "Processor";

            await tab.ToggleTableViewAsync();
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.IsTableOpen);
            Assert.False(tab.IsComponentOpen);
            Assert.False(tab.ComponentPaneForTests.IsVisible);
            Assert.Equal("Processor", ComponentPaneEdits.RowShown(tab.TableEditor.DraftSessionForTests!.Content.Board, "U8", "PAL")!.FriendlyName);

            await tab.OpenComponentAsync("U8");

            Assert.False(tab.IsTableOpen);
            Assert.True(tab.IsComponentOpen);
            Assert.True(tab.ComponentPaneForTests.IsVisible);
            Assert.False(tab.TableEditor.IsVisible);
        });
    }

    // ###########################################################################################
    // *** CLOSING THE TABLE SETTLES ITS OWN BOARD'S DRAFT, NOT EVERY ONE (code review, 2026-10-26). ***
    // Another board's draft with a change that cannot be saved made closing this table ask about
    // "the table's" changes, every time - and Discard read that other draft's file back, its change
    // gone without it ever being shown. It waits for its own board (or quitting) instead.
    // ###########################################################################################
    [Fact]
    public async Task Closing_the_table_leaves_another_boards_unsaved_draft_alone()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry, Entry("250425", BoardB));
            bool asked = false;
            tab.UnsavedTableEditsAnswerForTests = _ => { asked = true; return UnsavedTableEditsChoice.Discard; };
            tab.ShowBoard(entry, Board(), "PAL", null);

            // Board B's draft holds a change its file was changed under, so it cannot be saved.
            DraftSession other = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardB)!;
            other.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });
            BoardData elsewhere = Board();
            elsewhere.Components.Add(new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "Added elsewhere" });
            WriteDraft(BoardB, elsewhere);

            await tab.ToggleTableViewAsync();
            Assert.True(tab.IsTableOpen);

            Assert.True(await tab.CloseTableAsync());

            Assert.False(asked);
            Assert.False(tab.IsTableOpen);
            Assert.True(other.HasUnsavedChanges);
            Assert.Equal("906122", ComponentPaneEdits.RowShown(other.Content.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** ANOTHER BOARD SHOWN CLOSES THE TABLE, WHATEVER ITS SAVE DOES (code review, 2026-10-26). ***
    // CRT shows the other board already, so the close cannot be called off: asked, Cancel left the
    // table open on the previous board's draft under the new board's header. What it held stays in
    // that draft, held in memory, for its own board or quitting to save or ask about.
    // ###########################################################################################
    [Fact]
    public async Task Showing_another_board_closes_the_table_without_asking_and_keeps_what_could_not_be_saved()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entryA = Entry("250407", BoardA);
            HardwareBoardEntry entryB = Entry("250425", BoardB);
            TabDrafts tab = Tab(entryA, entryB);
            bool asked = false;
            tab.UnsavedTableEditsAnswerForTests = _ => { asked = true; return UnsavedTableEditsChoice.Cancel; };
            tab.ShowBoard(entryA, Board(), "PAL", null);

            await tab.ToggleTableViewAsync();
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Typed in the table", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            // Written somewhere else meanwhile, so its save is refused.
            BoardData elsewhere = Board();
            elsewhere.Components.Add(new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "Added elsewhere" });
            WriteDraft(BoardA, elsewhere);

            tab.ShowBoard(entryB, Board(), "PAL", null);
            await Until(() => session.SaveState == DraftSaveState.Failed);

            Assert.False(asked);
            Assert.False(tab.IsTableOpen);
            Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.SaveFailure);
            Assert.Equal("906122", ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** A DRAFT NO EDITOR HOLDS IS SAVED TOO (code review, 2026-10-26). *** Only the table and the
    // pane had a save timer, each for its own draft; a change in a draft neither showed was saved only
    // if whoever let go of it remembered to. The tab's one timer saves every draft with a save due
    // that neither holds - those two save their own, which the board on screen follows.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_no_editor_holds_is_saved_when_due_and_the_panes_is_left_to_the_pane()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry, Entry("250425", BoardB));
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.AutoSave = false;
            DraftSession shown = tab.ComponentPaneForTests.Session!;
            shown.Apply("In the pane's draft", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            DraftSession other = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardB)!;
            other.Apply("Nobody shows it", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "251715") });

            await tab.SaveDueDraftsNoEditorHoldsAsync(DateTimeOffset.UtcNow.AddMinutes(1));

            Assert.False(other.HasUnsavedChanges);
            Assert.Equal("251715", ComponentPaneEdits.RowShown(BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardB))!.Board, "U8", "PAL")!.PartNumber);
            Assert.True(shown.HasUnsavedChanges);
        });
    }

    // ###########################################################################################
    // The component pane's "Capture from my oscilloscope" (slice 3b) follows the Oscilloscope tab's
    // session as Main hands it to the tab - offered while connected, gone when not.
    // ###########################################################################################
    [Fact]
    public async Task The_pane_offers_the_scope_capture_while_Main_says_a_scope_is_connected()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);
            await tab.OpenComponentAsync("U8");

            Button capture = tab.ComponentPaneForTests.GetControl<Button>("CaptureBaselineButton");
            Assert.False(capture.IsVisible);

            tab.UseScopeCapture(_ => Task.FromResult<byte[]?>(null), connected: true);
            Assert.True(capture.IsVisible);

            tab.ShowScopeConnected(false);
            Assert.False(capture.IsVisible);
        });
    }

    // ###########################################################################################
    // Opening the table waits for the pane's close-save before it reads the draft - under CRT's
    // "please wait" (code review, 2026-10-26): the contributor pressed for it, and a large board's
    // save takes a second or two in which the tab took clicks on a half-switched workspace.
    // ###########################################################################################
    [Fact]
    public async Task Opening_the_table_waits_for_the_panes_save_under_please_wait()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            var overlay = new BusyOverlay();
            _ = new Grid { Children = { tab, overlay } };
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");

            var save = new TaskCompletionSource();
            tab.ComponentPaneForTests.UseLastSaveForTests(save.Task);

            Task opening = tab.ToggleTableViewAsync();
            Dispatcher.UIThread.RunJobs();

            Assert.True(overlay.IsBusy);
            Assert.False(tab.IsTableOpen);

            save.SetResult();
            await opening;

            Assert.False(overlay.IsBusy);
            Assert.True(tab.IsTableOpen);
        });
    }

    // ###########################################################################################
    // *** THE PANE IS A VIEW OF THE BOARD SHOWN. *** Another board shown closes it - what it held
    // handed to its draft and saved as it goes.
    // ###########################################################################################
    [Fact]
    public async Task Showing_another_board_closes_the_pane_and_saves_what_it_held()
    {
        WriteDraft(BoardA, Board());
        WriteDraft(BoardB);

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry first = Entry("250407", BoardA);
            HardwareBoardEntry second = Entry("250425", BoardB);
            TabDrafts tab = Tab(first, second);
            tab.ShowBoard(first, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            DraftSession session = tab.ComponentPaneForTests.Session!;
            Assert.Equal("U8", tab.ComponentPaneForTests.Label);
            Assert.True(tab.ComponentPaneForTests.GetControl<ScrollViewer>("SectionsScrollViewer").IsVisible);
            tab.ComponentPaneForTests.GetControl<TextBox>("DescriptionBox").Text = "Runs the show";

            tab.ShowBoard(second, Board(), "PAL", null);
            Assert.False(tab.IsComponentOpen);
            Assert.Equal("Runs the show", ComponentPaneEdits.RowShown(session.Content.Board, "U8", "PAL")!.Description);

            // The save the pane started as it closed - finished before the file is read, and before
            // the next test runs (a save still finishing then refreshed the tab against that test's
            // drafts folder).
            await tab.ComponentPaneForTests.LastSave;
            Dispatcher.UIThread.RunJobs();
            Assert.Null(session.SaveFailure);
            Assert.False(session.HasUnsavedChanges);

            string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA);
            Assert.Equal("Runs the show", ComponentPaneEdits.RowShown(BoardFiles.ReadContent(boardFile)!.Board, "U8", "PAL")!.Description);
        });
    }

    // ###########################################################################################
    // What the pane holds counts as unsaved for quitting and installing an update - and is saved
    // before either (Main.ConfirmLeavingTablesAsync).
    // ###########################################################################################
    [Fact]
    public async Task What_the_pane_holds_counts_as_unsaved_and_is_saved_before_leaving()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.AutoSave = false;
            tab.ComponentPaneForTests.GetControl<TextBox>("PartNumberBox").Text = "906122";

            Assert.True(tab.CommitCellsAndHasUnsavedTableEdits());
            Assert.True(tab.HasUnsavedTableEditsFor(BoardA));
            Assert.True(await tab.SaveDraftEditorsNowAsync());
            Assert.False(tab.CommitCellsAndHasUnsavedTableEdits());

            BoardContent written = BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA))!;
            Assert.Equal("906122", ComponentPaneEdits.RowShown(written.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** A DRAFT NEITHER THE PANE NOR THE TABLE SHOWS STILL COUNTS. *** A save the pane started as
    // it closed can fail (the file held open for a moment); the change then waits in the draft held
    // in memory, and quitting must save it - it asked only about an open pane or table, and the
    // change was gone.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_with_changes_not_saved_that_nothing_shows_is_saved_before_leaving()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            Assert.False(tab.IsComponentOpen);
            Assert.False(tab.IsTableOpen);
            Assert.True(tab.CommitCellsAndHasUnsavedTableEdits());

            Assert.True(await tab.SaveDraftEditorsNowAsync());
            Assert.False(tab.CommitCellsAndHasUnsavedTableEdits());

            BoardContent written = BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA))!;
            Assert.Equal("906122", ComponentPaneEdits.RowShown(written.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ###########################################################################################
    // *** SUBMIT AND "MANAGE FILES" SETTLE THE DRAFT ITSELF TOO (code review, 2026-10-26). *** They
    // saved only an open table or pane, so a change waiting in the draft held in memory that neither
    // shows - the pane's close-save refused, the header's Undo with nothing open - was left out: the
    // submission sent the board file without it.
    // ###########################################################################################
    [Fact]
    public async Task A_change_nothing_shows_is_saved_before_the_draft_is_submitted()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            Assert.True(await tab.SettleTableEditsAsync(entry));

            Assert.False(session.HasUnsavedChanges);
            BoardContent written = BoardFiles.ReadContent(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA))!;
            Assert.Equal("906122", ComponentPaneEdits.RowShown(written.Board, "U8", "PAL")!.PartNumber);
        });
    }

    // ...and when that save cannot be made, the contributor is asked - and Cancel stops the submit.
    [Fact]
    public async Task A_change_nothing_shows_that_cannot_be_saved_is_asked_about_before_the_draft_is_submitted()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);
            UnsavedTableEditsPrompt? asked = null;
            tab.UnsavedTableEditsAnswerForTests = prompt => { asked = prompt; return UnsavedTableEditsChoice.Cancel; };

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, BoardA)!;
            session.Apply("Waiting", content => content with { Board = ComponentPaneEdits.WithField(content.Board, "U8", "PAL", ComponentField.PartNumber, "906122") });

            // Written somewhere else meanwhile, so the save refuses rather than write over it.
            BoardData elsewhere = Board();
            elsewhere.Components.Add(new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "Added elsewhere" });
            WriteDraft(BoardA, elsewhere);

            Assert.False(await tab.SettleTableEditsAsync(entry));

            Assert.Equal(UnsavedTableEditsPrompt.DraftChangedOnDisk, asked);
            Assert.True(session.HasUnsavedChanges);
        });
    }

    // ###########################################################################################
    // The left column follows the pane at once - a name typed there is in the list, with the draft's
    // "Changed" mark - while the board on screen catches up later.
    // ###########################################################################################
    [Fact]
    public async Task The_left_column_follows_the_pane()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "Processor";
            tab.ComponentPaneForTests.CommitHeld();

            ComponentListItem u8 = tab.ComponentsShownForTests.Single(item => item.BoardLabel == "U8");
            Assert.Contains("Processor", u8.DisplayText, StringComparison.Ordinal);
            Assert.True(u8.IsDrafted);
        });
    }

    // ###########################################################################################
    // A draft discarded while its component is open: the pane lets go - nothing handed over, nothing
    // written back into the folder that is gone.
    // ###########################################################################################
    [Fact]
    public async Task Discarding_the_draft_lets_go_of_the_pane_without_writing_it_back()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "Processor";

            tab.DiscardConfirmed(entry);

            Assert.False(tab.IsComponentOpen);
            Assert.True(tab.ComponentPaneForTests.LastSave.IsCompleted);
            Assert.False(File.Exists(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardA)));
        });
    }

    // ###########################################################################################
    // "Add component" (slice 3c) opens the pane's form on the board shown; the component added is in
    // the left column at once, marked as the draft's change.
    // ###########################################################################################
    [Fact]
    public async Task Add_component_opens_the_panes_form_and_the_list_follows_the_component_added()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.AddComponentAsync();

            Assert.True(tab.IsComponentOpen);
            Assert.True(tab.ComponentPaneForTests.IsNewComponentForm);

            tab.ComponentPaneForTests.GetControl<TextBox>("NewLabelBox").Text = "U9";
            tab.ComponentPaneForTests.GetControl<AutoCompleteBox>("NewCategoryBox").Text = "IC";
            Assert.True(tab.ComponentPaneForTests.AddNewComponent());

            ComponentListItem added = tab.ComponentsShownForTests.Single(item => item.BoardLabel == "U9");
            Assert.True(added.IsDrafted);
        });
    }

    // ###########################################################################################
    // "Show on board": the pane's change saved first, and - since the board on screen is then behind
    // the draft - Main told to read it again before it selects the component; with nothing saved
    // since, it selects at once. Leaving the tab reads nothing again after that.
    // ###########################################################################################
    [Fact]
    public async Task Show_on_board_saves_first_and_has_the_board_read_again_only_when_it_is_behind()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            var reloads = new List<string>();
            tab.ReloadBoardOverrideForTests = reloads.Add;
            var shown = new List<(string Label, bool ReadFirst)>();
            tab.ShowOnBoardOverrideForTests = (label, readFirst) => shown.Add((label, readFirst));
            tab.ShowBoard(entry, Board(), "PAL", null);

            // On screen, as in CRT: a save of the pane's then leaves the board on screen behind.
            var window = new Window { Content = tab, Width = 1300, Height = 900 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            await tab.OpenComponentAsync("U8");
            await tab.ShowComponentOnBoardAsync("U8");

            tab.ComponentPaneForTests.GetControl<TextBox>("FriendlyNameBox").Text = "Processor";
            await tab.ShowComponentOnBoardAsync("U8");

            Assert.Equal([("U8", false), ("U8", true)], shown);
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardA)!;
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal("Processor", session.Content.Board.Components.Single(row => row.BoardLabel == "U8").FriendlyName);

            // Handed over: the board on screen has nothing left to catch up with.
            tab.CatchUpBoardOnScreen();
            Assert.Empty(reloads);

            window.Close();
        });
    }

    // "Preview popup" is of the draft as it is in memory - what the boxes hold included.
    [Fact]
    public async Task Preview_popup_is_of_the_draft_held_in_memory()
    {
        WriteDraft(BoardA, Board());

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = Entry("250407", BoardA);
            TabDrafts tab = Tab(entry);
            tab.PublishedBoardOverrideForTests = _ => Board();
            var previews = new List<(BoardData Board, string Label, string Identity)>();
            tab.PreviewOverrideForTests = (board, label, identity) => previews.Add((board, label, identity));
            tab.ShowBoard(entry, Board(), "PAL", null);

            await tab.OpenComponentAsync("U8");
            tab.ComponentPaneForTests.GetControl<TextBox>("DescriptionBox").Text = "Runs the show";
            tab.ComponentPaneForTests.GetControl<Button>("PreviewPopupButton").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            (BoardData board, string label, string identity) = Assert.Single(previews);
            Assert.Equal("U8", label);
            Assert.Equal(BoardA, identity);
            Assert.Equal("Runs the show", board.Components.Single(row => row.BoardLabel == "U8").Description);
        });
    }

    // ###########################################################################################
    // Why no draft was made is about the board it was pressed on: showing another board clears it
    // (code review, 2026-10-09); reading the same board again does not.
    // ###########################################################################################
    [Fact]
    public void A_problem_making_the_draft_is_cleared_when_another_board_is_shown()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = Tab();
            HardwareBoardEntry first = Entry("250407", BoardA);
            tab.ShowBoard(first, Board(), "PAL", null);

            tab.ShowDraftProblem("No draft could be made of this board: the published board could not be read.");
            tab.ShowBoard(first, Board(), "PAL", null);
            Assert.StartsWith("No draft could be made", tab.DraftProblemForTests);

            tab.ShowBoard(Entry("250425", BoardB), Board(), "PAL", null);
            Assert.Null(tab.DraftProblemForTests);
        });
    }
}
