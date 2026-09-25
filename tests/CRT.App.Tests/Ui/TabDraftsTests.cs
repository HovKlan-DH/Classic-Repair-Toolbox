using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// TabDrafts.RefreshDrafts - listing every system with a local draft (NewContributeStrategy.md
// Phase 2, session 2b, task 7) and the empty-state fallback. Drives DraftManager's real static
// state via LoadFrom (its own test seam), and TabDrafts.HardwareBoardsOverrideForTests to avoid
// needing a real main Excel workbook loaded through DataManager - the same override-then-real
// pattern TabWorkbooks.BoardKeyOverrideForTests uses, for the same reason.
[Collection("HeadlessUi")]
public sealed class TabDraftsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);

        // The drift tests below point DataManager at the workspace so a relative ExcelDataFile
        // resolves to a real board workbook. Reset it so a temp path never outlives this class -
        // collections never run in parallel (xunit.runner.json), so the "DataManager" collection
        // cannot be running concurrently, but a stale root would still confuse whatever ran next.
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");

        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry BoardEntry(string hardware, string board, string excelDataFile) => new()
    {
        HardwareName = hardware,
        BoardName = board,
        ExcelDataFile = excelDataFile
    };

    // ###########################################################################################
    // *** A DRAFT IS A BOARD FOLDER NOW, so these helpers write one (Phase 6, 2026-09-23). ***
    //
    // They used to save a BoardDraft holding N delta rows, and the tab counted those rows for
    // its "N rows changed" chip. The chip is now DERIVED by comparing the draft workbook against
    // the published one, so a test wanting a count of N has to create a genuine difference of N
    // rows - which is what WriteDraftWithChanges does: it publishes an empty board and drafts one
    // holding N components, so every one of them reads as an addition.
    //
    // baseRevision drives the drift check and is written into the marker, where it now lives.
    // ###########################################################################################
    private static void WriteDraftWithChanges(
        string excelDataFile,
        int rowCount = 1,
        string baseRevision = "",
        NewSystemRegistration? registration = null)
    {
        var board = new BoardData();
        for (int i = 0; i < rowCount; i++)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = $"U{i}" });
        }

        string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
        BoardWorkbookWriter.Write(workbook, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker
            {
                SystemKey = excelDataFile,
                BaseRevision = baseRevision,
                NewSystem = registration,
            });
    }

    [Fact]
    public void With_no_drafts_at_all_the_empty_state_shows_and_the_list_is_hidden()
    {
        UiTest.Run(() =>
        {
            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", "Commodore/C64/250407/Data.xlsx")
            };

            tab.RefreshDrafts();

            Assert.Empty(tab.Drafts);
            Assert.True(tab.GetControl<TextBlock>("EmptyStateText").IsVisible);
            Assert.False(tab.GetControl<ScrollViewer>("DraftsScrollViewer").IsVisible);
        });
    }

    [Fact]
    public void A_drafted_system_appears_in_the_list_and_the_empty_state_is_hidden()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            var item = Assert.Single(tab.Drafts);
            Assert.Equal("Commodore 64 - 250407", item.DisplayName);
            Assert.True(tab.GetControl<ScrollViewer>("DraftsScrollViewer").IsVisible);
            Assert.False(tab.GetControl<TextBlock>("EmptyStateText").IsVisible);
        });
    }

    [Fact]
    public void A_system_with_no_draft_is_not_listed_alongside_one_that_has_one()
    {
        UiTest.Run(() =>
        {
            string withDraft = "Commodore/C64/250407/Data.xlsx";
            string withoutDraft = "Commodore/C64/250425/Data.xlsx";
            WriteDraftWithChanges(withDraft);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", withDraft),
                BoardEntry("Commodore 64", "250425", withoutDraft)
            };

            tab.RefreshDrafts();

            var item = Assert.Single(tab.Drafts);
            Assert.Equal("Commodore 64 - 250407", item.DisplayName);
        });
    }

    [Fact]
    public void The_row_summary_counts_every_changed_row_across_all_sections()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile, rowCount: 3);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            Assert.Equal("3 rows changed", Assert.Single(tab.Drafts).RowSummary);
        });
    }

    [Fact]
    public void A_single_changed_row_reads_in_the_singular()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile, rowCount: 1);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            Assert.Equal("1 row changed", Assert.Single(tab.Drafts).RowSummary);
        });
    }

    [Fact]
    public void Multiple_drafted_systems_are_sorted_by_their_short_hardware_board_label()
    {
        // ShortHardwareBoardLabel is "<segments[^3]>/<segments[^2]>" of ExcelDataFile - the
        // immediate two folders above the file itself (HardwareFolder/BoardFolder), NOT
        // HardwareName/BoardName (the main Excel sheet's own display names, which is what
        // DisplayName below reads). Varying only the topmost path segment (as an earlier version
        // of this test did) left both fixtures sharing the same last-two-segments label and
        // proved nothing about sort order - the board FOLDER name is what has to differ.
        UiTest.Run(() =>
        {
            string zSystem = "Commodore/Zeta board/rev1/Data.xlsx";
            string aSystem = "Commodore/Alpha board/rev1/Data.xlsx";
            WriteDraftWithChanges(zSystem);
            WriteDraftWithChanges(aSystem);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore", "Zeta board", zSystem),
                BoardEntry("Commodore", "Alpha board", aSystem)
            };

            tab.RefreshDrafts();

            Assert.Equal(2, tab.Drafts.Count);
            Assert.Equal("Commodore - Alpha board", tab.Drafts[0].DisplayName);
            Assert.Equal("Commodore - Zeta board", tab.Drafts[1].DisplayName);
        });
    }

    // ------------------------------------ A system that exists only as a draft (task 9)

    private static NewSystemRegistration NewRegistration(string excelDataFile) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = "MyBoard",
        ExcelDataFile = excelDataFile,
    };

    // The IsEmpty fix proven at the UI layer: a system created through "Add a new system" has zero
    // rows in every section, and if it did not count as a draft it would never be listed here -
    // which is also the only place it could be discarded from.
    [Fact]
    public void A_newly_created_system_with_no_rows_at_all_is_still_listed()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
            WriteDraftWithChanges(excelDataFile, rowCount: 0, registration: NewRegistration(excelDataFile));

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "MyBoard", excelDataFile)
            };

            tab.RefreshDrafts();

            Assert.Single(tab.Drafts);
            Assert.Equal("Commodore 64 - MyBoard", tab.Drafts[0].DisplayName);
        });
    }

    // "0 rows changed" would be actively misleading on a brand-new system: nothing was CHANGED
    // because the whole board is new, and the count says nothing about what was actually made.
    [Fact]
    public void A_newly_created_system_is_described_as_new_rather_than_by_a_row_count()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
            WriteDraftWithChanges(excelDataFile, rowCount: 0, registration: NewRegistration(excelDataFile));

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "MyBoard", excelDataFile)
            };

            tab.RefreshDrafts();

            Assert.Equal("New system, nothing added yet", tab.Drafts[0].RowSummary);
            Assert.DoesNotContain("changed", tab.Drafts[0].RowSummary);
        });
    }

    [Fact]
    public void A_new_system_that_has_been_worked_on_says_how_much_is_in_it()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
            WriteDraftWithChanges(
                excelDataFile,
                rowCount: 1,
                registration: NewRegistration(excelDataFile));

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "MyBoard", excelDataFile)
            };

            tab.RefreshDrafts();

            Assert.Equal("New system, 1 row so far", tab.Drafts[0].RowSummary);
        });
    }

    // An ordinary draft over a synced system keeps the original wording - what it lists really is
    // a set of changes to a board that already exists.
    [Fact]
    public void An_ordinary_draft_still_reads_as_rows_changed()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile, rowCount: 2);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            Assert.Equal("2 rows changed", tab.Drafts[0].RowSummary);
        });
    }

    // ------------------------------------------------------ The drift warning (session 2d)

    // A board path unique to each test. BoardDataReader's load cache is a PROCESS-WIDE static keyed
    // by the ExcelDataFile string, and this class is not in the "BoardData" collection that clears
    // it - so a shared path would let one test's cached revision leak into another's answer. A
    // fresh path per test is a guaranteed cache miss, the same reasoning (and the same fix)
    // DataManagerDraftOverlayTests documents for itself.
    private readonly string thisExcelDataFile =
        $"Commodore/C64/{Guid.NewGuid():N}/Data C64 drift.xlsx";

    // Drift needs a real board workbook on disk to read an official revision from, plus a draft
    // recording a different base. DataManager is pointed at the workspace so the relative
    // ExcelDataFile resolves.
    private string WriteBoardWithRevision(string excelDataFile, string revisionDate)
    {
        string path = Path.Combine(
            this.thisWorkspace.Root,
            excelDataFile.Replace('/', Path.DirectorySeparatorChar));

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookBuilder.WriteCompleteBoard(path, revisionDate);

        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");

        return path;
    }

    // The base revision drives the drift check, and since Phase 6 it lives in the MARKER
    // rather than in the draft's own rows.
    private static void WriteDraftBasedOn(string excelDataFile, string baseRevision) =>
        WriteDraftWithChanges(excelDataFile, rowCount: 1, baseRevision: baseRevision);

    private TabDrafts BuildTabFor(string excelDataFile, string hardware = "Commodore 64", string board = "250407")
    {
        var tab = new TabDrafts();
        tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
        {
            BoardEntry(hardware, board, excelDataFile)
        };

        tab.RefreshDrafts();
        return tab;
    }

    [Fact]
    public void A_draft_whose_board_has_been_updated_officially_is_flagged()
    {
        UiTest.Run(() =>
        {
            string ExcelDataFile = this.thisExcelDataFile;
            this.WriteBoardWithRevision(ExcelDataFile, "2026-August-21");
            WriteDraftBasedOn(ExcelDataFile, "2026-May-12");

            var tab = this.BuildTabFor(ExcelDataFile);

            var item = Assert.Single(tab.Drafts);
            Assert.True(item.HasDrift);
            Assert.Contains("has been updated", item.DriftSummary);

            // The reassurance is load-bearing: without it the warning reads as "your work may be
            // lost", which is false - Drafts/ and Data/ are separate trees.
            Assert.Contains("Your edits are still applied", item.DriftSummary);
        });
    }

    // Worded to match what the data supports: when an ORDER could not be established, the message
    // must say "changed", never "updated".
    [Fact]
    public void An_unorderable_revision_difference_says_changed_rather_than_updated()
    {
        UiTest.Run(() =>
        {
            string ExcelDataFile = this.thisExcelDataFile;
            this.WriteBoardWithRevision(ExcelDataFile, "revision two");
            WriteDraftBasedOn(ExcelDataFile, "revision one");

            var item = Assert.Single(this.BuildTabFor(ExcelDataFile).Drafts);

            Assert.True(item.HasDrift);
            Assert.Contains("has changed", item.DriftSummary);
            Assert.DoesNotContain("has been updated", item.DriftSummary);
        });
    }

    [Fact]
    public void A_draft_started_from_the_current_revision_is_not_flagged()
    {
        UiTest.Run(() =>
        {
            string ExcelDataFile = this.thisExcelDataFile;
            this.WriteBoardWithRevision(ExcelDataFile, "2026-August-21");
            WriteDraftBasedOn(ExcelDataFile, "2026-August-21");

            var item = Assert.Single(this.BuildTabFor(ExcelDataFile).Drafts);

            Assert.False(item.HasDrift);
            Assert.Equal(string.Empty, item.DriftSummary);
        });
    }

    // A draft written before every save path recorded a base revision. An unrecorded base is not
    // evidence of drift, so it must stay silent rather than warn on a guess.
    [Fact]
    public void A_draft_with_no_recorded_base_revision_is_not_flagged()
    {
        UiTest.Run(() =>
        {
            string ExcelDataFile = this.thisExcelDataFile;
            this.WriteBoardWithRevision(ExcelDataFile, "2026-August-21");
            WriteDraftBasedOn(ExcelDataFile, string.Empty);

            Assert.False(Assert.Single(this.BuildTabFor(ExcelDataFile).Drafts).HasDrift);
        });
    }

    // A system that exists only as a draft has no official counterpart at all, so drift is
    // meaningless for it - and its blank base revision is deliberate, not legacy.
    [Fact]
    public void A_draft_only_system_is_never_flagged_as_drifted()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/MyBoard/Data C64 MyBoard.xlsx";
            WriteDraftWithChanges(excelDataFile, rowCount: 0, registration: NewRegistration(excelDataFile));

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                new()
                {
                    HardwareName = "Commodore 64",
                    BoardName = "MyBoard",
                    ExcelDataFile = excelDataFile,
                    IsDraftOnly = true,
                }
            };

            tab.RefreshDrafts();

            Assert.False(Assert.Single(tab.Drafts).HasDrift);
        });
    }

    // The button has to be here as well as on the Contribute tab - this is where someone already
    // working on drafts looks for it.
    [Fact]
    public void The_tab_offers_a_button_to_add_a_new_system()
    {
        UiTest.Run(() =>
        {
            var tab = new TabDrafts();

            var button = tab.GetControl<Button>("AddNewSystemButton");

            Assert.NotNull(button);
            Assert.Equal("Add a new system", button.Content);
        });
    }

    // Offered unconditionally, including before anything has ever been sent: the window's own
    // empty state teaches what will appear there, whereas a button that materialises only after
    // the first submission is one nobody finds until they no longer need telling.
    [Fact]
    public void The_tab_offers_a_my_submissions_button_even_with_no_drafts()
    {
        UiTest.Run(() =>
        {
            var tab = new TabDrafts();

            var button = tab.GetControl<Button>("MySubmissionsButton");

            Assert.NotNull(button);
            Assert.True(button.IsEnabled);

            // *** THE LABEL MOVED INSIDE THE BUTTON when the unread badge was added. *** Content
            // is a StackPanel now, not a string, so this reads the label out of it rather than
            // off Content - the words a contributor sees are unchanged, and that is what this
            // test was always about.
            Assert.Equal("My submissions", TabDraftsTests.LabelOf(button));
        });
    }

    // Pulls the visible label out of a button whose Content is a panel. Returns the FIRST
    // TextBlock, which is the label; the badge's own count sits in a later one.
    private static string LabelOf(Button button) =>
        (button.Content as Panel)?
            .Children
            .OfType<TextBlock>()
            .FirstOrDefault()?
            .Text
        ?? button.Content as string
        ?? string.Empty;

    // ###########################################################################################
    // THE UNREAD-FEEDBACK BADGE on the My submissions button.
    //
    // *** WHY IT EXISTS: a comment nobody notices is a comment nobody reads. *** Contributing
    // needs no account, so the reviewer's sentence is the whole channel back to the contributor,
    // and it lived behind a button nobody had a reason to press. Somebody could be asked for
    // changes and never find out.
    //
    // Driven through UnreadCommentCountOverrideForTests rather than the real receipts store,
    // which is a static singleton reading the user's own AppData folder.
    // ###########################################################################################

    [Fact]
    public void The_badge_is_HIDDEN_when_there_is_no_unread_feedback()
    {
        UiTest.Run(() =>
        {
            var tab = new TabDrafts
            {
                HardwareBoardsOverrideForTests = [],
                UnreadCommentCountOverrideForTests = 0
            };

            tab.RefreshDrafts();

            // The ordinary button must look exactly as it did before this feature existed.
            Assert.False(tab.GetControl<Border>("MySubmissionsBadge").IsVisible);
        });
    }

    [Fact]
    public void The_badge_SHOWS_THE_COUNT_when_a_reviewer_has_said_something_unread()
    {
        UiTest.Run(() =>
        {
            var tab = new TabDrafts
            {
                HardwareBoardsOverrideForTests = [],
                UnreadCommentCountOverrideForTests = 2
            };

            tab.RefreshDrafts();

            Assert.True(tab.GetControl<Border>("MySubmissionsBadge").IsVisible);

            // The NUMBER matters, not just that something appeared: "(2)" tells the contributor
            // two different submissions want attention, which is what decides whether they open
            // the window now or later.
            Assert.Equal("2", tab.GetControl<TextBlock>("MySubmissionsBadgeText").Text);
        });
    }

    [Fact]
    public void With_NO_DRAFTS_but_unread_feedback_the_empty_state_says_so()
    {
        UiTest.Run(() =>
        {
            // *** THE ONE CASE THIS TAB APPEARS EMPTY. *** With no drafts it is normally hidden
            // outright; unread reviewer feedback is what holds it open
            // (Main.ApplyDraftsTabVisibility), because a contributor who submitted and then
            // discarded their draft would otherwise never see the reply.
            //
            // "No local drafts yet" would then be a true sentence answering the wrong question,
            // leaving somebody staring at an empty tab with no idea a reviewer had written.
            var tab = new TabDrafts
            {
                HardwareBoardsOverrideForTests = [],
                UnreadCommentCountOverrideForTests = 1
            };

            tab.RefreshDrafts();

            var empty = tab.GetControl<TextBlock>("EmptyStateText");

            Assert.True(empty.IsVisible);
            Assert.Contains("reviewer has replied", empty.Text!, StringComparison.Ordinal);
            Assert.Contains("My submissions", empty.Text!, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void With_no_drafts_and_NO_feedback_the_empty_state_is_the_ordinary_one()
    {
        UiTest.Run(() =>
        {
            // Somebody who has never contributed must see no trace of the review feature.
            var tab = new TabDrafts
            {
                HardwareBoardsOverrideForTests = [],
                UnreadCommentCountOverrideForTests = 0
            };

            tab.RefreshDrafts();

            var empty = tab.GetControl<TextBlock>("EmptyStateText");

            Assert.Contains("No local drafts yet", empty.Text!, StringComparison.Ordinal);
            Assert.DoesNotContain("reviewer", empty.Text!, StringComparison.OrdinalIgnoreCase);
        });
    }

    // ###########################################################################################
    // THE SUBMIT BUTTON (Phase 4) - sending a draft for review.
    //
    // The click itself opens a modal that talks to a real server, so these tests assert on the
    // row model the button is bound to (CanSubmit / SubmitTooltip / SubmitCommand) rather than
    // pressing it - the same division ConfigurationHelpIconTests documents for a button whose
    // handler reaches outside the app.
    // ###########################################################################################

    [Fact]
    public void A_draft_with_rows_can_be_submitted()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            DraftListItem row = Assert.Single(tab.Drafts);

            Assert.True(row.CanSubmit);
            Assert.NotNull(row.SubmitCommand);
            Assert.True(row.SubmitCommand.CanExecute(null));
        });
    }

    // ###########################################################################################
    // A system registered but not yet filled in is a REAL state, not a theoretical one: "Add a new
    // system" creates the registration, and this tab lists it from that moment. Submitting it
    // would put an empty system in front of a reviewer, so the button is disabled and says why.
    //
    // Disabled rather than hidden, deliberately - an action that disappears reads as the app
    // having lost it, and the tooltip is where the reason can actually be given.
    // ###########################################################################################
    [Fact]
    public void A_brand_new_system_with_nothing_in_it_yet_cannot_be_submitted_and_says_why()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data C64 250407.xlsx";

            // A draft is listed because the MARKER exists, so a brand-new system with nothing in
            // it yet still gets a row - which is the only surface it can be discarded from.
            WriteDraftWithChanges(
                excelDataFile,
                rowCount: 0,
                registration: new NewSystemRegistration
                {
                    HardwareName = "C64",
                    BoardName = "250407",
                    ExcelDataFile = excelDataFile,
                });

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("C64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            DraftListItem row = Assert.Single(tab.Drafts);

            Assert.False(row.CanSubmit);
            Assert.Contains("nothing to send yet", row.SubmitTooltip);
        });
    }

    // Drift must NOT block submitting. The server diffs against the base revision itself, and
    // refusing to send while the official data has moved would strand a contributor behind a
    // change somebody else made.
    [Fact]
    public void A_draft_whose_official_data_has_moved_can_still_be_submitted()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";

            WriteDraftWithChanges(excelDataFile, rowCount: 1, baseRevision: "2026-01-01");

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            DraftListItem row = Assert.Single(tab.Drafts);

            Assert.True(row.CanSubmit);
        });
    }

    // ###########################################################################################
    // THE BINDING ITSELF, through a real layout pass.
    //
    // Every test above reads the row MODEL, which proves the rule and not the wiring: a correct
    // CanSubmit bound to the wrong property, or to a button that is not there, ships broken and
    // every one of them still passes. Avalonia tolerates a failed binding silently (the same trap
    // CLAUDE.md records for a missing StaticResource key), so the only way to know the button is
    // really gated is to build the row and read the Button that comes out.
    //
    // This is also what catches the column count: the row Grid was widened from four columns to
    // five to make space, and a Submit button left at the Discard button's index would sit on top
    // of it rather than beside it.
    // ###########################################################################################
    [Fact]
    public void The_rendered_row_really_carries_a_Submit_button_that_is_gated_by_CanSubmit()
    {
        UiTest.Run(() =>
        {
            string withRows = "Commodore/C64/250407/Data.xlsx";
            string empty = "Commodore/VIC20/250408/Data VIC20 250408.xlsx";

            WriteDraftWithChanges(withRows);

            WriteDraftWithChanges(
                empty,
                rowCount: 0,
                registration: new NewSystemRegistration
                {
                    HardwareName = "VIC20",
                    BoardName = "250408",
                    ExcelDataFile = empty,
                });

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", withRows),
                BoardEntry("VIC20", "250408", empty)
            };

            tab.RefreshDrafts();

            var window = new Window { Content = tab, Width = 1100, Height = 600 };
            window.Show();
            window.Measure(new Size(1100, 600));
            window.Arrange(new Rect(0, 0, 1100, 600));
            Dispatcher.UIThread.RunJobs();

            List<Button> submitButtons = tab.GetVisualDescendants()
                .OfType<Button>()
                .Where(button => (button.Content as string) == "Submit")
                .ToList();

            // One per drafted system, not one for the tab.
            Assert.Equal(2, submitButtons.Count);

            // The row WITH rows is enabled; the empty new system is not. Asserted as a set rather
            // than by index, since the list is ordered by label and that ordering is not this
            // test's subject.
            Assert.Contains(submitButtons, button => button.IsEnabled);
            Assert.Contains(submitButtons, button => !button.IsEnabled);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** "Schematic images" AND "KiCad data" ARE TWO BUTTONS (maintainer request, 2026-09-24). ***
    //
    // They used to be one "Schematic images and KiCad data" button. Asserted on the RENDERED row
    // for the same reason as the Submit test above: a button bound to the wrong command, or left
    // at another button's column index and drawn on top of it, passes every model-level test. So
    // this checks each button exists, is wired to its OWN command, and sits clear of the others.
    // ###########################################################################################
    [Fact]
    public void The_rendered_row_has_separate_Schematic_images_and_KiCad_data_buttons()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            var window = new Window { Content = tab, Width = 1300, Height = 600 };
            window.Show();
            window.Measure(new Size(1300, 600));
            window.Arrange(new Rect(0, 0, 1300, 600));
            Dispatcher.UIThread.RunJobs();

            List<Button> buttons = tab.GetVisualDescendants().OfType<Button>().ToList();
            Button Named(string content) => Assert.Single(buttons, button => (button.Content as string) == content);

            DraftListItem row = Assert.Single(tab.Drafts);

            Assert.Same(row.ManageSchematicImagesCommand, Named("Schematic images").Command);
            Assert.Same(row.ManageKiCadDataCommand, Named("KiCad data").Command);
            Assert.Same(row.EditTableCommand, Named("Edit in table format").Command);
            Assert.DoesNotContain(buttons, button => (button.Content as string) == "Schematic images and KiCad data");

            // Every button of the row in a column of its own - two sharing an index draw on top of
            // each other, and nothing throws.
            var rowButtons = new[] { "Edit in table format", "Schematic images", "KiCad data", "Submit", "Discard" }
                .Select(Named)
                .ToList();

            Assert.Equal(rowButtons.Count, rowButtons.Select(Grid.GetColumn).Distinct().Count());

            window.Close();
        });
    }

    // The tooltip is the one place the no-account rule is stated where a contributor will actually
    // read it. It was the whole point of the Phase 4 rework: contributing must never ask anyone to
    // register, and a button offering to "submit" with no further word invites the assumption that
    // it will.
    [Fact]
    public void The_submit_tooltip_says_no_account_is_needed()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", excelDataFile)
            };

            tab.RefreshDrafts();

            string tooltip = Assert.Single(tab.Drafts).SubmitTooltip;

            Assert.Contains("do not need an account", tooltip);
            Assert.Contains("email address", tooltip);
        });
    }

    // ###########################################################################################
    // TABLE MODE - "Edit in table format" (maintainer request, 2026-09-24).
    //
    // The table itself is pinned by BoardTableEditorTests; these pin the TAB around it: that
    // opening one draft's table hides every other row but leaves `Drafts` whole (Main decides
    // whether this tab is shown at all from its count), and that unsaved table edits are never
    // lost without the contributor saying so. The prompt is answered through
    // UnsavedTableEditsAnswerForTests - the real window would block on ShowDialog.
    // ###########################################################################################
    private const string TableSystemA = "Commodore/C64/250407/Data.xlsx";
    private const string TableSystemB = "Commodore/C64/250425/Data.xlsx";

    private static TabDrafts TabWithTwoDrafts()
    {
        WriteDraftWithChanges(TabDraftsTests.TableSystemA);
        WriteDraftWithChanges(TabDraftsTests.TableSystemB);

        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = new List<HardwareBoardEntry>
            {
                BoardEntry("Commodore 64", "250407", TabDraftsTests.TableSystemA),
                BoardEntry("Commodore 64", "250425", TabDraftsTests.TableSystemB),
            },

            // Nothing published to colour against - these tests are about the tab, not colours.
            PublishedBoardOverrideForTests = _ => null,
        };

        tab.RefreshDrafts();

        return tab;
    }

    private static HardwareBoardEntry EntryA(TabDrafts tab) =>
        tab.HardwareBoardsOverrideForTests!.First(entry => entry.ExcelDataFile == TabDraftsTests.TableSystemA);

    // The draft rows the tab actually SHOWS: the list, or - in table mode - the open draft's row in
    // its own host, with the list hidden.
    private static IEnumerable<DraftListItem> ShownRows(TabDrafts tab)
    {
        ContentControl openRow = tab.GetControl<ContentControl>("OpenDraftRow");
        if (openRow.IsVisible)
        {
            return openRow.Content is DraftListItem item ? [item] : [];
        }

        return tab.GetControl<ScrollViewer>("DraftsScrollViewer").IsVisible
            ? tab.GetControl<ItemsControl>("DraftsItemsControl").ItemsSource!.Cast<DraftListItem>()
            : [];
    }

    // Types into the open table's first component, leaving an unsaved edit.
    private static void MakeUnsavedEdit(TabDrafts tab)
    {
        BoardTableEditor editor = tab.GetControl<BoardTableEditor>("TableEditor");
        BoardTableSheet components = editor.SessionForTests!.Document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendlyName = components.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);

        components.Rows[0].Cells[friendlyName].Text = "Typed in the table";
        Assert.True(tab.HasUnsavedTableEdits);
    }

    private static string FirstFriendlyNameOnDisk(string excelDataFile) =>
        DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, excelDataFile)!.Components[0].FriendlyName;

    [Fact]
    public void Every_draft_row_offers_Edit_in_table_format_while_no_table_is_open()
    {
        UiTest.Run(() =>
        {
            TabDrafts tab = TabWithTwoDrafts();

            Assert.All(tab.Drafts, row =>
            {
                Assert.False(row.IsTableOpen);
                Assert.Equal("Edit in table format", row.TableButtonText);
            });
            Assert.False(tab.GetControl<BoardTableEditor>("TableEditor").IsVisible);
        });
    }

    [Fact]
    public async Task Opening_the_table_shows_ONLY_that_drafts_row_with_the_table_below_it()
    {
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();

            await tab.OpenTableAsync(EntryA(tab));

            DraftListItem shown = Assert.Single(ShownRows(tab));
            Assert.Equal(TabDraftsTests.TableSystemA, shown.ExcelDataFile);
            Assert.True(shown.IsTableOpen);
            Assert.Equal("Close table", shown.TableButtonText);

            Assert.True(tab.GetControl<BoardTableEditor>("TableEditor").IsVisible);
            Assert.True(tab.GetControl<BoardTableEditor>("TableEditor").HasTable);
            Assert.False(tab.GetControl<ScrollViewer>("DraftsScrollViewer").IsVisible);

            // Drafts itself still lists BOTH - Main reads its count to decide whether this tab is
            // shown at all, and a table open on one draft must not hide the tab's reason to exist.
            Assert.Equal(2, tab.Drafts.Count);
        });
    }

    [Fact]
    public async Task In_table_mode_the_open_drafts_row_is_really_drawn_above_the_table()
    {
        // Measured, not just listed. The first version drew the open row at ZERO height: it
        // switched layouts by assigning a new RowDefinitions collection to the body Grid, which
        // kept measuring its cells against the old row kinds. The row, and its "Close table"
        // button (the only way out), was there yet drawn nowhere. This fails against that version.
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            var window = new Window { Content = tab, Width = 1300, Height = 700 };
            window.Show();

            await tab.OpenTableAsync(EntryA(tab));
            Dispatcher.UIThread.RunJobs();

            ContentControl openRow = tab.GetControl<ContentControl>("OpenDraftRow");
            BoardTableEditor editor = tab.GetControl<BoardTableEditor>("TableEditor");

            Assert.True(openRow.Bounds.Height > 0, $"The open draft's row is {openRow.Bounds.Height}px tall");
            Assert.True(editor.Bounds.Height > openRow.Bounds.Height, "The table should take the rest of the height");
            Assert.True(editor.Bounds.Top >= openRow.Bounds.Bottom, "The table should sit BELOW the row");

            Button close = tab.GetVisualDescendants().OfType<Button>().Single(button => (button.Content as string) == "Close table");
            Assert.True(close.Bounds.Height > 0);

            window.Close();
        });
    }

    [Fact]
    public async Task The_table_survives_a_refresh_of_the_list()
    {
        // Main refreshes this tab after every board load; that must not throw the table away.
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            tab.RefreshDrafts();

            Assert.True(tab.IsTableOpen);
            Assert.True(tab.HasUnsavedTableEdits);
            Assert.Single(ShownRows(tab));
        });
    }

    [Fact]
    public async Task Closing_a_table_with_nothing_unsaved_goes_straight_back_to_the_list()
    {
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            tab.UnsavedTableEditsAnswerForTests = _ => throw new InvalidOperationException("Nothing unsaved - nothing to ask.");

            await tab.OpenTableAsync(EntryA(tab));

            Assert.True(await tab.CloseTableAsync());

            Assert.False(tab.IsTableOpen);
            Assert.Equal(2, ShownRows(tab).Count());
            Assert.False(tab.GetControl<BoardTableEditor>("TableEditor").IsVisible);
        });
    }

    [Fact]
    public async Task Closing_with_unsaved_edits_and_CANCEL_keeps_the_table_and_the_edits()
    {
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            tab.UnsavedTableEditsAnswerForTests = _ => UnsavedTableEditsChoice.Cancel;

            Assert.False(await tab.CloseTableAsync());

            Assert.True(tab.IsTableOpen);
            Assert.True(tab.HasUnsavedTableEdits);
        });
    }

    [Fact]
    public async Task Closing_with_unsaved_edits_and_DISCARD_closes_without_writing_anything()
    {
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            tab.UnsavedTableEditsAnswerForTests = _ => UnsavedTableEditsChoice.Discard;

            Assert.True(await tab.CloseTableAsync());

            Assert.False(tab.IsTableOpen);
            Assert.NotEqual("Typed in the table", FirstFriendlyNameOnDisk(TabDraftsTests.TableSystemA));
        });
    }

    [Fact]
    public async Task Closing_with_unsaved_edits_and_SAVE_writes_the_draft_then_closes()
    {
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            tab.UnsavedTableEditsAnswerForTests = _ => UnsavedTableEditsChoice.Save;

            Assert.True(await tab.CloseTableAsync());

            Assert.False(tab.IsTableOpen);
            Assert.Equal("Typed in the table", FirstFriendlyNameOnDisk(TabDraftsTests.TableSystemA));
        });
    }

    [Fact]
    public async Task A_save_that_is_REFUSED_keeps_the_table_open_rather_than_losing_the_edits()
    {
        // Choosing Save and having it refused (the file changed underneath) must not be read as
        // "fine to leave" - the edits would be gone with nothing written.
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            WriteDraftWithChanges(TabDraftsTests.TableSystemA, rowCount: 3);
            tab.UnsavedTableEditsAnswerForTests = _ => UnsavedTableEditsChoice.Save;

            Assert.False(await tab.CloseTableAsync());

            Assert.True(tab.IsTableOpen);
            Assert.True(tab.HasUnsavedTableEdits);
        });
    }

    [Fact]
    public async Task Closing_when_the_draft_changed_on_disk_asks_WITHOUT_Save_and_discarding_closes()
    {
        // Reported as a loop: after "Save to draft" in the Contribute tab had written the same
        // draft, "Close table" offered Save, the save was refused, the table stayed open, and the
        // next Close asked again. The question now knows a save cannot land and does not offer it.
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            var asked = new List<UnsavedTableEditsPrompt>();
            tab.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return UnsavedTableEditsChoice.Discard;
            };

            WriteDraftWithChanges(TabDraftsTests.TableSystemA, rowCount: 3);

            Assert.True(await tab.CloseTableAsync());

            Assert.Equal([UnsavedTableEditsPrompt.DraftChangedOnDisk], asked);
            Assert.False(tab.IsTableOpen);
        });
    }

    [Fact]
    public async Task Closing_while_Excel_holds_the_draft_open_asks_WITHOUT_Save()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File sharing is advisory outside Windows.");

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            var asked = new List<UnsavedTableEditsPrompt>();
            tab.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return UnsavedTableEditsChoice.Cancel;
            };

            string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, TabDraftsTests.TableSystemA);
            File.WriteAllText(Path.Combine(Path.GetDirectoryName(workbook)!, "~$" + Path.GetFileName(workbook)), "owner");

            using (new FileStream(workbook, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                Assert.False(await tab.CloseTableAsync());
            }

            Assert.Equal([UnsavedTableEditsPrompt.DraftOpenElsewhere], asked);
            Assert.True(tab.IsTableOpen);
        });
    }

    [Fact]
    public async Task Closing_with_the_draft_unchanged_on_disk_asks_the_ordinary_question_with_Save()
    {
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));
            MakeUnsavedEdit(tab);

            var asked = new List<UnsavedTableEditsPrompt>();
            tab.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return UnsavedTableEditsChoice.Cancel;
            };

            Assert.False(await tab.CloseTableAsync());

            Assert.Equal([UnsavedTableEditsPrompt.Leaving], asked);
            Assert.True(tab.IsTableOpen);
        });
    }

    [Fact]
    public async Task Other_editors_are_told_about_unsaved_table_edits_only_for_the_same_board()
    {
        // The question the Contribute tab's "Save to draft" and the label editor ask before
        // writing a draft (2026-09-24): only unsaved edits in a table on THAT board stop them.
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();

            Assert.False(tab.HasUnsavedTableEditsFor(TabDraftsTests.TableSystemA));

            await tab.OpenTableAsync(EntryA(tab));
            Assert.False(tab.HasUnsavedTableEditsFor(TabDraftsTests.TableSystemA));

            MakeUnsavedEdit(tab);
            Assert.True(tab.HasUnsavedTableEditsFor(TabDraftsTests.TableSystemA));
            Assert.True(tab.HasUnsavedTableEditsFor(TabDraftsTests.TableSystemA.ToUpperInvariant()));
            Assert.False(tab.HasUnsavedTableEditsFor(TabDraftsTests.TableSystemB));
            Assert.False(tab.HasUnsavedTableEditsFor(null));

            Assert.Equal(DraftWorkbookEditOutcome.Saved, tab.TableEditor.Save());
            Assert.False(tab.HasUnsavedTableEditsFor(TabDraftsTests.TableSystemA));
        });
    }

    [Fact]
    public async Task When_the_table_follows_an_outside_change_the_drafts_row_count_follows_too()
    {
        // Reported: the table reloaded after Excel changed the draft, but the draft's row kept its
        // old "N rows changed". The Drafts tab now refreshes on that reload as it does on a save.
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));

            string before = tab.Drafts.Single(draft => draft.ExcelDataFile == TabDraftsTests.TableSystemA).RowSummary;

            WriteDraftWithChanges(TabDraftsTests.TableSystemA, rowCount: 5);
            tab.TableEditor.CheckDraftFile();

            string after = tab.Drafts.Single(draft => draft.ExcelDataFile == TabDraftsTests.TableSystemA).RowSummary;
            Assert.NotEqual(before, after);
            Assert.Equal("5 rows changed", after);
            Assert.True(tab.IsTableOpen);
        });
    }

    [Fact]
    public async Task A_table_whose_draft_disappears_closes_on_the_next_refresh()
    {
        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = TabWithTwoDrafts();
            await tab.OpenTableAsync(EntryA(tab));

            // Discarded from elsewhere (or retired after publishing).
            DraftManager.DiscardDraft(TabDraftsTests.TableSystemA);
            tab.RefreshDrafts();

            Assert.False(tab.IsTableOpen);
            Assert.False(tab.GetControl<BoardTableEditor>("TableEditor").HasTable);
            Assert.Equal(TabDraftsTests.TableSystemB, Assert.Single(ShownRows(tab)).ExcelDataFile);
        });
    }

    // ###########################################################################################
    // *** A DISCARD THAT CANNOT FINISH SAYS SO (code review, 2026-09-25). ***
    //
    // The delete used to throw straight out of a fire-and-forget command when a file in the draft
    // was locked - on Windows, its workbook open in Excel - leaving a half-deleted folder, a list
    // that never refreshed and no word to the contributor. It now reports the failure on the tab,
    // worded for the likely cause, and still refreshes the list.
    // ###########################################################################################
    [Fact]
    public void A_discard_blocked_by_a_LOCKED_file_is_reported_on_the_tab()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile);

            HardwareBoardEntry entry = BoardEntry("Commodore 64", "250407", excelDataFile);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry> { entry };
            tab.RefreshDrafts();

            string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);

            using (new FileStream(workbook, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                tab.DiscardConfirmed(entry);
            }

            string? message = tab.DiscardFailureTextForTests;

            Assert.NotNull(message);
            Assert.Contains(entry.ToString(), message);
            Assert.Contains("Excel", message);
        });
    }

    [Fact]
    public void A_discard_that_succeeds_shows_no_failure_and_empties_the_list()
    {
        UiTest.Run(() =>
        {
            string excelDataFile = "Commodore/C64/250407/Data.xlsx";
            WriteDraftWithChanges(excelDataFile);

            HardwareBoardEntry entry = BoardEntry("Commodore 64", "250407", excelDataFile);

            var tab = new TabDrafts();
            tab.HardwareBoardsOverrideForTests = new List<HardwareBoardEntry> { entry };
            tab.RefreshDrafts();
            Assert.Single(tab.Drafts);

            tab.DiscardConfirmed(entry);

            Assert.Null(tab.DiscardFailureTextForTests);
            Assert.Empty(tab.Drafts);
        });
    }
}
