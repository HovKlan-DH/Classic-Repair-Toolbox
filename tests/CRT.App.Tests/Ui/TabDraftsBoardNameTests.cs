using Avalonia.Controls;
using Avalonia.LogicalTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Board name" on the Contribute tab (owner request, 2026-11-01: "how to edit those details again
// afterwards - I think this should be the very first item in the list in 'Board contents'"): a page
// on the right, a new board's names changed there, a published board's only shown.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsBoardNameTests : IDisposable
{
    private const string NewBoard = "Comodore/C64/999001/board.json";
    private const string Published = "Commodore/Commodore 64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsBoardNameTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry NewEntry() => new()
    {
        HardwareName = "C64",
        BoardName = "999001",
        ExcelDataFile = NewBoard,
        IsDraftOnly = true,
    };

    private static HardwareBoardEntry PublishedEntry() => new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        ExcelDataFile = Published,
    };

    private static void SeedNewBoard(string notes = "") =>
        Assert.True(DraftSeeder.CreateNewBoard(DraftManager.DraftsRoot, new NewBoardRegistration
        {
            HardwareName = "C64",
            BoardName = "999001",
            HardwareNotes = notes,
            ExcelDataFile = NewBoard,
        }).Created);

    private static TabDrafts Tab(params HardwareBoardEntry[] entries)
    {
        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [.. entries],
            PublishedBoardOverrideForTests = _ => null,
            ReceiptsOverrideForTests = [],
        };

        tab.RefreshDrafts();
        return tab;
    }

    private static BoardNameFields FieldsOf(TabDrafts tab) => tab.BoardNamePanelForTests.GetControl<BoardNameFields>("Fields");

    private static void Fill(TabDrafts tab, string manufacturer, string hardware, string board)
    {
        BoardNameFields fields = FieldsOf(tab);

        fields.GetControl<AutoCompleteBox>("ManufacturerBox").Text = manufacturer;
        fields.GetControl<TextBox>("HardwareBox").Text = hardware;
        fields.GetControl<TextBox>("BoardBox").Text = board;
        tab.BoardNamePanelForTests.RefreshForTests();
    }

    // The very first entry of the left column, and selected while its page is what is open.
    [Fact]
    public async Task Board_name_comes_after_Board_status_and_is_selected_while_open()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            // Owner request, 2026-11-01: "must be shown first, as this is the overview for the board".
            List<Button> entries = tab.GetControl<Grid>("WorkspaceNav").GetLogicalDescendants().OfType<Button>().ToList();
            Assert.Equal("StatusNavButton", entries[0].Name);
            Assert.Equal("NameNavButton", entries[1].Name);

            await tab.OpenBoardNameAsync();

            Assert.True(tab.IsBoardNameOpen);
            Assert.True(tab.BoardNamePanelForTests.IsVisible);
            Assert.False(tab.GetControl<StackPanel>("StartPanel").IsVisible);
            Assert.Equal("NameNavButton", tab.SelectedNavEntryForTests);

            // Board status goes back; the page closes.
            await tab.ShowBoardStatusAsync();

            Assert.False(tab.IsBoardNameOpen);
            Assert.False(tab.BoardNamePanelForTests.IsVisible);
            Assert.Equal("StatusNavButton", tab.SelectedNavEntryForTests);
        });
    }

    // Anything else opened lets the page go - one thing on the right at a time.
    [Fact]
    public async Task Another_entry_closes_the_page()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();
            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);

            Assert.False(tab.IsBoardNameOpen);
            Assert.False(tab.BoardNamePanelForTests.IsVisible);
            Assert.Equal("CreditsNavButton", tab.SelectedNavEntryForTests);
        });
    }

    // ###########################################################################################
    // A new board's page holds what it was made with - Save off until something else is typed, and off
    // for a name another board has, saying so. Saved, the draft has moved and Main is told.
    // ###########################################################################################
    [Fact]
    public async Task A_new_boards_name_is_changed_on_its_page()
    {
        SeedNewBoard(notes: "My notes");

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            HardwareBoardEntry other = PublishedEntry();
            TabDrafts tab = Tab(entry, other);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            var told = new List<(HardwareBoardEntry Entry, NewBoardRegistration Renamed)>();
            tab.BoardRenamedOverrideForTests = (from, renamed) => told.Add((from, renamed));

            await tab.OpenBoardNameAsync();

            BoardNameFields fields = FieldsOf(tab);
            Assert.Equal("Comodore", fields.GetControl<AutoCompleteBox>("ManufacturerBox").Text);
            Assert.Equal("C64", fields.GetControl<TextBox>("HardwareBox").Text);
            Assert.Equal("999001", fields.GetControl<TextBox>("BoardBox").Text);
            Assert.Equal("My notes", fields.GetControl<TextBox>("NotesBox").Text);
            Assert.False(tab.BoardNamePanelForTests.CanSaveForTests);

            // Another board's name is refused, as when a board is made.
            Fill(tab, "Commodore", "Commodore 64", "250407");
            Assert.False(tab.BoardNamePanelForTests.CanSaveForTests);
            Assert.Contains("already exists", tab.BoardNamePanelForTests.GetControl<TextBlock>("ValidationText").Text, StringComparison.Ordinal);

            Fill(tab, "Commodore", "Commodore 64", "999001");
            Assert.True(tab.BoardNamePanelForTests.CanSaveForTests);

            await tab.BoardNamePanelForTests.SaveAsync();

            // Refused, the page says why - which is what is wanted when this fails.
            Assert.True(told.Count == 1, $"The rename was refused: [{tab.BoardNamePanelForTests.GetControl<TextBlock>("ValidationText").Text}]");

            (HardwareBoardEntry from, NewBoardRegistration renamed) = Assert.Single(told);
            Assert.Same(entry, from);
            Assert.Equal("Commodore/Commodore 64/999001/board.json", renamed.ExcelDataFile);

            Assert.False(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, NewBoard));
            Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, renamed.ExcelDataFile));
            Assert.Null(DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard));

            // Moved, the board shown is another board: Board status follows.
            Assert.False(tab.IsBoardNameOpen);
        });
    }

    // What the draft holds goes along: saved before the folder moves.
    [Fact]
    public async Task What_the_draft_holds_is_saved_before_it_moves()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);
            tab.BoardRenamedOverrideForTests = (_, _) => { };

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, NewBoard)!;
            session.Apply("Waiting", content => content with
            {
                Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Schematic image", NameOrHandle = "Anna" }]),
            });

            await tab.OpenBoardNameAsync();
            Fill(tab, "Commodore", "C64", "999001");
            await tab.BoardNamePanelForTests.SaveAsync();

            string moved = NewBoardIdentity.BuildExcelDataFile("Commodore", "C64", "999001");
            BoardDocument read = BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, moved))!;

            Assert.Equal("Anna", Assert.Single(read.Credits).NameOrHandle);
        });
    }

    // ###########################################################################################
    // Undo steps go only with a draft that really moved (code review, 2026-11-01): the session was closed
    // before anything was tried, so notes alone, and a rename refused with "Nothing was changed", lost
    // them too - while the page warns only of a new name.
    // ###########################################################################################
    [Fact]
    public async Task Notes_alone_keep_the_drafts_Undo_steps()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);
            tab.BoardRenamedOverrideForTests = (_, _) => { };

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, NewBoard)!;
            session.Apply("Waiting", content => content with
            {
                Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Schematic image", NameOrHandle = "Anna" }]),
            });

            await tab.OpenBoardNameAsync();
            FieldsOf(tab).GetControl<TextBox>("NotesBox").Text = "Now with notes";
            tab.BoardNamePanelForTests.RefreshForTests();
            await tab.BoardNamePanelForTests.SaveAsync();

            Assert.Same(session, DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard));
            Assert.True(session.CanUndo);
            Assert.False(session.FileChangedOutside());
        });
    }

    [Fact]
    public async Task A_rename_refused_keeps_the_drafts_Undo_steps()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);
            tab.BoardRenamedOverrideForTests = (_, _) => { };

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, NewBoard)!;
            session.Apply("Waiting", content => content with
            {
                Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Schematic image", NameOrHandle = "Anna" }]),
            });

            // A folder already at the new place refuses the rename.
            Directory.CreateDirectory(DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, NewBoardIdentity.BuildExcelDataFile("Commodore", "C64", "999001")));

            await tab.OpenBoardNameAsync();
            Fill(tab, "Commodore", "C64", "999001");
            await tab.BoardNamePanelForTests.SaveAsync();

            Assert.True(tab.IsBoardNameOpen);
            Assert.Same(session, DraftManager.Sessions.Find(DraftManager.DraftsRoot, NewBoard));
            Assert.True(session.CanUndo);
            Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, NewBoard));
        });
    }

    // A picture tool on holds the rename back: its Done writes the draft as it holds it.
    [Fact]
    public async Task A_picture_tool_on_holds_the_rename_back()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);
            tab.PictureToolOnOverrideForTests = () => true;

            await tab.OpenBoardNameAsync();
            Fill(tab, "Commodore", "C64", "999001");
            await tab.BoardNamePanelForTests.SaveAsync();

            Assert.Equal(PictureToolWording.PictureToolOn, tab.BoardNamePanelForTests.GetControl<TextBlock>("ValidationText").Text);
            Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, NewBoard));
        });
    }

    // Notes alone move nothing: the page stays, and says it saved.
    [Fact]
    public async Task Notes_alone_stay_on_the_page()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);
            tab.BoardRenamedOverrideForTests = (_, _) => { };

            await tab.OpenBoardNameAsync();

            FieldsOf(tab).GetControl<TextBox>("NotesBox").Text = "Now with notes";
            tab.BoardNamePanelForTests.RefreshForTests();
            Assert.True(tab.BoardNamePanelForTests.CanSaveForTests);

            await tab.BoardNamePanelForTests.SaveAsync();

            Assert.True(tab.IsBoardNameOpen);
            Assert.Equal(BoardNameWording.Saved, tab.BoardNamePanelForTests.GetControl<TextBlock>("SavedText").Text);
            Assert.False(tab.BoardNamePanelForTests.CanSaveForTests);
            Assert.Equal("Now with notes", DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, NewBoard))!.NewBoard!.HardwareNotes);
        });
    }

    // ###########################################################################################
    // A published board's name is only shown, why said - and looking at it makes no draft.
    // ###########################################################################################
    [Fact]
    public async Task A_published_boards_name_is_only_shown_and_makes_no_draft()
    {
        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = PublishedEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();

            BoardNamePanel panel = tab.BoardNamePanelForTests;
            Assert.True(panel.GetControl<Border>("WhyNotPanel").IsVisible);
            Assert.Equal(BoardNameWording.Published, panel.GetControl<TextBlock>("WhyNotText").Text);
            Assert.False(panel.GetControl<StackPanel>("SavePanel").IsVisible);
            Assert.False(FieldsOf(tab).GetControl<TextBox>("HardwareBox").IsEnabled);
            Assert.False(FieldsOf(tab).GetControl<AutoCompleteBox>("ManufacturerBox").IsEnabled);
            Assert.Equal("Commodore 64", FieldsOf(tab).GetControl<TextBox>("HardwareBox").Text);

            Assert.False(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, Published));
        });
    }

    // ###########################################################################################
    // Sent for review, a new board keeps its name while a maintainer has it - the server knows it by
    // the name, and a submission under a new one would be a second board. Turned down, it can change.
    // ###########################################################################################
    [Fact]
    public async Task A_new_board_sent_for_review_keeps_its_name_until_that_is_decided()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ReceiptsOverrideForTests = [new SubmissionReceipt { SubmissionId = 7, BoardId = "Comodore/C64/999001", LastKnownState = "pending" }];
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();

            Assert.Equal(BoardNameWording.SentForReview, tab.BoardNamePanelForTests.GetControl<TextBlock>("WhyNotText").Text);
            Assert.False(tab.BoardNamePanelForTests.GetControl<StackPanel>("SavePanel").IsVisible);

            tab.ReceiptsOverrideForTests = [new SubmissionReceipt { SubmissionId = 7, BoardId = "Comodore/C64/999001", LastKnownState = "rejected" }];
            await tab.OpenBoardNameAsync();

            Assert.False(tab.BoardNamePanelForTests.GetControl<Border>("WhyNotPanel").IsVisible);
            Assert.True(tab.BoardNamePanelForTests.GetControl<StackPanel>("SavePanel").IsVisible);
        });
    }

    // ###########################################################################################
    // *** ONLY A DECISION AGAINST IT FREES THE NAME (code review, 2026-11-01). *** A submission accepted
    // or published is final too, so "no longer open" freed the name of a board that had reached the data
    // - on a machine that had not synced yet it still reads as a new board, and the next Submit under a
    // new name made a second board beside the published one.
    // ###########################################################################################
    [Theory]
    [InlineData("accepted", BoardNameWording.Published)]
    [InlineData("published", BoardNameWording.Published)]
    [InlineData("merged", BoardNameWording.SentForReview)]
    public async Task A_new_board_that_reached_the_data_keeps_its_name(string state, string reason)
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ReceiptsOverrideForTests = [new SubmissionReceipt { SubmissionId = 7, BoardId = "Comodore/C64/999001", LastKnownState = state }];
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();

            Assert.Equal(reason, tab.BoardNamePanelForTests.GetControl<TextBlock>("WhyNotText").Text);
            Assert.False(tab.BoardNamePanelForTests.GetControl<StackPanel>("SavePanel").IsVisible);
        });
    }

    [Theory]
    [InlineData("rejected")]
    [InlineData("abandoned")]
    [InlineData("withdrawn")]
    public async Task A_new_board_whose_submission_was_decided_against_can_be_renamed(string state)
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ReceiptsOverrideForTests = [new SubmissionReceipt { SubmissionId = 7, BoardId = "Comodore/C64/999001", LastKnownState = state }];
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();

            Assert.False(tab.BoardNamePanelForTests.GetControl<Border>("WhyNotPanel").IsVisible);
            Assert.True(tab.BoardNamePanelForTests.GetControl<StackPanel>("SavePanel").IsVisible);
        });
    }

    // ###########################################################################################
    // *** A SEND THAT NEVER FINISHED LEAVES THE NAME FREE (code review, 2026-11-02). *** A dropped
    // connection leaves the receipt with no state and then "uploading": no maintainer has it, yet the
    // page said "has been sent for review ... until the maintainer has decided" - for up to a day, until
    // the sweep marked it abandoned. CRT never resumes an old upload, so nothing of it goes on.
    // ###########################################################################################
    [Theory]
    [InlineData("")]
    [InlineData("uploading")]
    public async Task A_send_that_never_finished_leaves_the_name_free(string state)
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ReceiptsOverrideForTests = [new SubmissionReceipt { SubmissionId = 7, BoardId = "Comodore/C64/999001", LastKnownState = state }];
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();

            Assert.False(tab.BoardNamePanelForTests.GetControl<Border>("WhyNotPanel").IsVisible);
            Assert.True(tab.BoardNamePanelForTests.GetControl<StackPanel>("SavePanel").IsVisible);
        });
    }

    // ###########################################################################################
    // A finalise request that went IS with the server, whatever became of its answer (UnfinishedSubmission):
    // the submission may be queued and the maintainers mailed. With nothing heard since, the name stays -
    // "uploading" from a check made BEFORE the request says nothing about it. A check after it does.
    // ###########################################################################################
    [Theory]
    [InlineData("", false)]
    [InlineData("uploading", false)]
    [InlineData("uploading", true)]
    public async Task A_finalise_whose_answer_was_lost_keeps_the_name_until_the_server_has_said(string state, bool askedSince)
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            DateTimeOffset sent = DateTimeOffset.UtcNow;

            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ReceiptsOverrideForTests =
            [
                new SubmissionReceipt
                {
                    SubmissionId = 7,
                    BoardId = "Comodore/C64/999001",
                    LastKnownState = state,
                    FinaliseSentUtc = sent,
                    LastCheckedUtc = sent.AddMinutes(askedSince ? 1 : -5),
                }
            ];
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();

            Assert.Equal(askedSince, tab.BoardNamePanelForTests.GetControl<StackPanel>("SavePanel").IsVisible);

            if (!askedSince)
            {
                Assert.Equal(BoardNameWording.SentForReview, tab.BoardNamePanelForTests.GetControl<TextBlock>("WhyNotText").Text);
            }
        });
    }

    // ###########################################################################################
    // *** A NEW BOARD'S NOTES ARE PART OF WHAT IT SENDS (code review, 2026-11-02). *** They live in the
    // marker, which the fingerprint skips as bookkeeping - so notes changed after a rejection said
    // "Saved." and Submit stayed grey with "Change something to send it again". Taken from
    // PrepareOtherDraftAsync (what a receipt remembers) and read back through the row, so the two
    // places a fingerprint is taken must agree: changed back, the row greys out again.
    // ###########################################################################################
    [Fact]
    public async Task Notes_changed_after_a_rejection_bring_Submit_back()
    {
        SeedNewBoard(notes: "First notes");

        // The names it was made with - a save of the page writes them into the board file, so a fixture
        // without them would change the file and bring Submit back by itself.
        CachedWorkbooks.Write(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, NewBoard), new BoardData
        {
            HardwareName = "C64",
            BoardName = "999001",
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            Credits =
            [
                new CreditEntry { Category = NewBoardCredits.BoardData, NameOrHandle = "Anna" },
                new CreditEntry { Category = NewBoardCredits.OscilloscopeBaseline, NameOrHandle = "Anna" },
            ],
        });

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);
            tab.BoardRenamedOverrideForTests = (_, _) => tab.RefreshDrafts();

            OtherDraftReady ready = await tab.PrepareOtherDraftAsync(entry, string.Empty, string.Empty);
            SubmitDraftTarget target = Assert.IsType<SubmitDraftTarget>(ready.Target);

            tab.ReceiptsOverrideForTests =
            [
                new SubmissionReceipt
                {
                    SubmissionId = 7,
                    BoardId = "Comodore/C64/999001",
                    SentUtc = DateTimeOffset.UtcNow.AddMinutes(1),
                    LastKnownState = "rejected",
                    DraftFingerprint = target.Fingerprint,
                }
            ];

            tab.RefreshDrafts();
            Assert.False(Assert.Single(tab.Drafts).CanSubmit);

            await tab.OpenBoardNameAsync();
            FieldsOf(tab).GetControl<TextBox>("NotesBox").Text = "Better notes";
            tab.BoardNamePanelForTests.RefreshForTests();
            await tab.BoardNamePanelForTests.SaveAsync();

            Assert.True(Assert.Single(tab.Drafts).CanSubmit);

            FieldsOf(tab).GetControl<TextBox>("NotesBox").Text = "First notes";
            tab.BoardNamePanelForTests.RefreshForTests();
            await tab.BoardNamePanelForTests.SaveAsync();

            Assert.False(Assert.Single(tab.Drafts).CanSubmit);
        });
    }

    // ###########################################################################################
    // A new name can land on a listed board's folder as a new board's can (code review, 2026-11-02):
    // the fields' one rule, on this page too.
    // ###########################################################################################
    [Fact]
    public async Task A_new_name_landing_on_a_listed_boards_folder_is_refused()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            var listed = new HardwareBoardEntry
            {
                HardwareName = "Commodore 64",
                BoardName = "250407 (long board)",
                ExcelDataFile = "Commodore/C64/250407/board.json",
            };

            TabDrafts tab = Tab(entry, listed);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();

            Fill(tab, "Commodore", "C64", "250407");

            Assert.False(tab.BoardNamePanelForTests.CanSaveForTests);
            Assert.Contains("250407 (long board)", tab.BoardNamePanelForTests.GetControl<TextBlock>("ValidationText").Text, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // The list read again - every window activation - with this page open works nothing out behind it
    // (code review, 2026-11-01): Board status came back for a moment, opened the draft's session, ran
    // every check and read the KiCad data, each time.
    // ###########################################################################################
    [Fact]
    public async Task Reading_the_list_again_works_out_no_Board_status_behind_the_page()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);

            await tab.OpenBoardNameAsync();
            int asks = tab.ReadinessAsksForTests;

            tab.RefreshDrafts();

            Assert.Equal(asks, tab.ReadinessAsksForTests);
            Assert.True(tab.IsBoardNameOpen);
        });
    }

    // Board status's panel opens the same page.
    [Fact]
    public async Task Board_statuss_panel_opens_the_page()
    {
        SeedNewBoard();

        await UiTest.RunAsync(async () =>
        {
            HardwareBoardEntry entry = NewEntry();
            TabDrafts tab = Tab(entry);
            tab.ShowBoard(entry, new BoardData(), "PAL", null);
            await tab.ReadinessShownForTests;

            tab.StatusButtonForTests(BoardStatusPart.Name).RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            for (int attempt = 0; attempt < 50 && !tab.IsBoardNameOpen; attempt++)
            {
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.True(tab.IsBoardNameOpen);
        });
    }
}
