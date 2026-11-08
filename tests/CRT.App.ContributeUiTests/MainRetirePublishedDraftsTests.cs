using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE WHOLE RETIRE STEP, ON A REAL Main (code review, 2026-11-02): Main.RetirePublishedDraftsAsync with
// its test-server gate passed in, over a published board in a temp data tree, a seeded draft and a
// published receipt. WHETHER a draft may go is DraftRetirement's, tested in CRT.Data; HOW the deleting is
// sequenced is PublishedDraftRetirer's, tested with delegates. What only this reaches is the wiring
// between them: the gate shut, no test or pre-release ever ran it, so taking "anything open" out of what
// Main asks left the suite green, and its first real run would have been on users' drafts after the
// release. StartAsync is never called (MainWindowTests' header).
//
// Every draft here equals the published board and was made before the receipt was sent, so each is
// retired unless the thing the test is about holds it.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainRetirePublishedDraftsTests : IDisposable
{
    private static readonly MainFiles.Board C64 = new("Commodore/C64/250407", "Commodore 64", "250407");

    private readonly TempWorkspace thisWorkspace = new();

    private static readonly HardwareBoardEntry Entry = new()
    {
        HardwareName = C64.HardwareName,
        BoardName = C64.BoardName,
        ExcelDataFile = C64.Identity,
    };

    private static BoardData Board() => new()
    {
        RevisionDate = "2026-August-7",
        Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" }],
    };

    public MainRetirePublishedDraftsTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        SubmissionReceiptStore.LoadFrom(this.thisWorkspace.Path_("submissions.json"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

        string dataRoot = this.thisWorkspace.Path_("Data");
        Directory.CreateDirectory(dataRoot);

        string main = MainFiles.Write(dataRoot, MainRetirePublishedDraftsTests.C64);
        CachedWorkbooks.Write(DraftBoardSource.PublishedPathOf(dataRoot, MainRetirePublishedDraftsTests.C64.Identity), MainRetirePublishedDraftsTests.Board());
        DataManager.LoadFrom(dataRoot, main);

        DraftSeedResult seeded = DraftSeeder.SeedFromPublishedFile(DraftManager.DraftsRoot, dataRoot, MainRetirePublishedDraftsTests.C64.Identity);
        Assert.True(seeded.Created, seeded.Reason);
    }

    public void Dispose()
    {
        DraftManager.Sessions.Close(DraftManager.DraftsRoot, MainRetirePublishedDraftsTests.C64.Identity);
        DraftManager.LoadFrom(string.Empty);
        SubmissionReceiptStore.LoadFrom(string.Empty);
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.json");
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static bool HasDraft => DraftBoardSource.HasDraft(DraftManager.DraftsRoot, MainRetirePublishedDraftsTests.C64.Identity);

    // The board's submission reached stable - sent after the draft was made, as a draft's work is.
    private static void RecordPublished() =>
        SubmissionReceiptStore.Record(new SubmissionReceipt
        {
            SubmissionId = 8,
            UploadToken = "tok8",
            BoardId = MainRetirePublishedDraftsTests.C64.BoardId,
            Summary = "Fixed U1",
            SentUtc = DateTimeOffset.UtcNow.AddHours(1),
            LastKnownState = "published",
        });

    private static CRT.Main Window()
    {
        var window = new CRT.Main();

        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static TabDrafts ShownOn(CRT.Main window)
    {
        TabDrafts tab = window.TabDrafts;
        tab.PublishedBoardOverrideForTests = _ => MainRetirePublishedDraftsTests.Board();
        tab.ShowBoard(MainRetirePublishedDraftsTests.Entry, MainRetirePublishedDraftsTests.Board(), "PAL", null);

        return tab;
    }

    // Undone and saved: the file says what the published board says, with the change still to be redone.
    private static DraftSession ChangedAndUndone()
    {
        DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, MainRetirePublishedDraftsTests.C64.Identity)!;

        Assert.True(session.Apply("Credits", content => content with
        {
            Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Board data", NameOrHandle = "Bo" }]),
        }));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.True(session.Undo());
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.False(session.HasUnsavedChanges);
        Assert.True(session.CanRedo);

        return session;
    }

    // The set-up's own proof: with nothing holding it, the step retires the draft.
    [Fact]
    public async Task A_draft_nothing_holds_goes()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainRetirePublishedDraftsTests.Window();
            MainRetirePublishedDraftsTests.RecordPublished();

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.False(MainRetirePublishedDraftsTests.HasDraft);

            window.Close();
        });
    }

    // The test instance's stable data does not last, so nothing is retired against it.
    [Fact]
    public async Task Nothing_goes_while_CRT_talks_to_the_test_server()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainRetirePublishedDraftsTests.Window();
            MainRetirePublishedDraftsTests.RecordPublished();

            await window.RetirePublishedDraftsAsync(testServer: true);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);

            window.Close();
        });
    }

    // ###########################################################################################
    // What is open on the draft holds it through the step, and the next step after it closes retires it.
    // ###########################################################################################
    [Fact]
    public async Task A_table_open_on_the_draft_holds_it_until_it_closes()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainRetirePublishedDraftsTests.Window();
            TabDrafts tab = MainRetirePublishedDraftsTests.ShownOn(window);
            MainRetirePublishedDraftsTests.RecordPublished();

            await tab.OpenTableAsync(MainRetirePublishedDraftsTests.Entry);

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);
            Assert.True(tab.IsTableOpen);

            Assert.True(await tab.CloseTableAsync());

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.False(MainRetirePublishedDraftsTests.HasDraft);

            window.Close();
        });
    }

    [Fact]
    public async Task The_pane_open_on_the_draft_holds_it_until_it_closes()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainRetirePublishedDraftsTests.Window();
            TabDrafts tab = MainRetirePublishedDraftsTests.ShownOn(window);
            MainRetirePublishedDraftsTests.RecordPublished();

            await tab.OpenBoardPartAsync(PaneBoardPart.Credits);

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);
            Assert.True(tab.IsComponentOpen);

            tab.CloseComponent();
            await tab.ComponentPaneForTests.LastSave;

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.False(MainRetirePublishedDraftsTests.HasDraft);

            window.Close();
        });
    }

    [Fact]
    public async Task The_KiCad_panel_open_on_the_draft_holds_it_until_it_closes()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainRetirePublishedDraftsTests.Window();
            TabDrafts tab = MainRetirePublishedDraftsTests.ShownOn(window);
            MainRetirePublishedDraftsTests.RecordPublished();

            await tab.OpenKiCadAsync();

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);
            Assert.True(tab.IsKiCadOpen);

            await tab.CloseKiCadAsync();

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.False(MainRetirePublishedDraftsTests.HasDraft);

            window.Close();
        });
    }

    // A change the session holds and its file does not: an editor's save that could not be made, say.
    [Fact]
    public async Task A_change_not_saved_yet_holds_the_draft()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainRetirePublishedDraftsTests.Window();
            MainRetirePublishedDraftsTests.RecordPublished();

            DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, MainRetirePublishedDraftsTests.C64.Identity)!;
            Assert.True(session.Apply("Credits", content => content with
            {
                Board = BoardPartEdits.WithCredits(content.Board, [new CreditEntry { Category = "Board data", NameOrHandle = "Bo" }]),
            }));

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);
            Assert.True(session.HasUnsavedChanges);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** REDO WENT WITH THE DRAFT (code review, 2026-11-02). *** The contributor's work reached stable,
    // and the draft was changed again; the header's Undo, with nothing open on the right, saved it equal
    // to the published board - and the next status check deleted it, the undone change with it, though
    // the header still offered Redo. Held while the session can redo; sessions end with the run, so the
    // next launch retires it as ever.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_with_a_step_to_redo_stays_this_run_and_goes_the_next()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = MainRetirePublishedDraftsTests.Window();
            MainRetirePublishedDraftsTests.RecordPublished();

            DraftSession session = MainRetirePublishedDraftsTests.ChangedAndUndone();

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);
            Assert.True(session.Redo());
            Assert.Single(session.Content.Board.Credits);

            Assert.True(session.Undo());
            Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

            // The run ends, and the session with it.
            DraftManager.Sessions.Close(DraftManager.DraftsRoot, MainRetirePublishedDraftsTests.C64.Identity);

            await window.RetirePublishedDraftsAsync(testServer: false);

            Assert.False(MainRetirePublishedDraftsTests.HasDraft);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A DRAFT HELD DRAWS NOTHING NOBODY ASKED FOR (code review, 2026-11-02). *** Found but not touched,
    // it still set the redraw, so every row and both drop-down templates were rebuilt at each minute check
    // while it stayed - what the 2026-10-04 rule stopped. The minute check asks for nothing when no
    // receipt moved; the badge says whether the tab was drawn.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_held_at_the_minute_check_does_not_draw_the_tab()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main();
            MainRetirePublishedDraftsTests.RecordPublished();
            MainRetirePublishedDraftsTests.ChangedAndUndone();

            await window.RetirePublishedDraftsAsync(refreshDraftsTab: false, testServer: false);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);
            Assert.Null(window.ContributeTabBadgeForTests);

            // What the caller asked for it still gets: a receipt moved, or the launch check.
            await window.RetirePublishedDraftsAsync(refreshDraftsTab: true, testServer: false);

            Assert.True(MainRetirePublishedDraftsTests.HasDraft);
            Assert.NotNull(window.ContributeTabBadgeForTests);
        });
    }

    // A draft actually retired is a change worth drawing, whatever the caller asked.
    [Fact]
    public async Task A_draft_retired_at_the_minute_check_draws_the_tab()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main();
            MainRetirePublishedDraftsTests.RecordPublished();

            await window.RetirePublishedDraftsAsync(refreshDraftsTab: false, testServer: false);

            Assert.False(MainRetirePublishedDraftsTests.HasDraft);
            Assert.NotNull(window.ContributeTabBadgeForTests);
        });
    }
}
