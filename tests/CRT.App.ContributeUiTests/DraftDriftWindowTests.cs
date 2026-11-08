using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// DraftDriftWindow - the "what changed officially" report (NewContributeStrategy.md Phase 2,
// session 2d).
//
// WHAT THESE COVER: how one already-built DraftDriftReport is RENDERED, and the rebase action's
// one guarantee - that it writes the base revision and leaves every drafted row alone. The
// detector's own logic is covered exhaustively in CRT.Data.Tests.
[Collection("HeadlessUi")]
public sealed class DraftDriftWindowTests : IDisposable
{
    private const string ExcelDataFile = "Commodore/C64/250407/Data.xlsx";

    private readonly TempWorkspace thisWorkspace = new();

    public DraftDriftWindowTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static BoardRowChange Row(
        string label,
        BoardRowChangeKind kind = BoardRowChangeKind.Modified,
        params string[] changedFields) => new()
    {
        Section = "Components",
        NaturalKey = BoardDraftNaturalKeys.ForComponent(label),
        DisplayLabel = label,
        Kind = kind,
        ChangedFields = changedFields,
    };

    private static DraftChangeReport Report(
        DraftDriftState state = DraftDriftState.OfficialIsNewer,
        string baseRevision = "2026-May-12",
        string officialRevision = "2026-August-21",
        params BoardRowChange[] rows) => new()
        {
            State = state,
            BaseRevision = baseRevision,
            OfficialRevision = officialRevision,
            Rows = rows,
        };

    // A draft that kept what it was copied from (code review, 2026-10-13): its own changes, the
    // update's, and where both met.
    private static DraftChangeReport ReportWithBase(
        BoardRowChange[]? own = null,
        BoardRowChange[]? update = null,
        DraftRebaseConflict[]? conflicts = null) => new()
        {
            State = DraftDriftState.OfficialIsNewer,
            BaseRevision = "2026-May-12",
            OfficialRevision = "2026-August-21",
            Rows = own ?? [],
            ComparedWithBase = true,
            UpdateRows = update ?? [],
            Conflicts = conflicts ?? [],
            CanBringUpToDate = true,
        };

    private static DraftDriftWindow WindowFor(DraftChangeReport report)
    {
        var window = new DraftDriftWindow();
        window.Initialize("Commodore 64 - 250407", ExcelDataFile, report);
        return window;
    }

    private static string TextOf(DraftDriftWindow window, string name) =>
        window.GetControl<TextBlock>(name).Text ?? string.Empty;

    [Fact]
    public void The_window_constructs_without_throwing()
    {
        UiTest.Run(() => Assert.NotNull(new DraftDriftWindow()));
    }

    // ------------------------------------------------------------------ The revisions

    [Fact]
    public void Both_revisions_are_shown_verbatim()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report());

            Assert.Equal("2026-May-12", TextOf(window, "BaseRevisionText"));
            Assert.Equal("2026-August-21", TextOf(window, "OfficialRevisionText"));
        });
    }

    // A draft written before every save path stamped a base revision genuinely has none. A blank
    // line here would read as a rendering fault rather than as the fact it is.
    [Fact]
    public void An_unrecorded_base_revision_says_so_rather_than_showing_a_blank()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report(state: DraftDriftState.Unknown, baseRevision: string.Empty));

            Assert.Equal("Not recorded", TextOf(window, "BaseRevisionText"));
        });
    }

    // ------------------------------------------------------------------ The summary wording

    // The wording must match what the data supports - "updated" only when an order was actually
    // established. See DraftRevisionComparer.
    [Fact]
    public void A_newer_official_revision_is_described_as_updated()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report(state: DraftDriftState.OfficialIsNewer));

            Assert.Contains("has been updated", TextOf(window, "SummaryText"));
        });
    }

    [Fact]
    public void An_unorderable_difference_is_described_only_as_changed()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report(state: DraftDriftState.Changed));

            string summary = TextOf(window, "SummaryText");
            Assert.Contains("has changed", summary);
            Assert.DoesNotContain("has been updated", summary);
        });
    }

    // ###########################################################################################
    // *** "YOUR EDITS ARE STILL APPLIED ON TOP OF IT" WAS UNTRUE, AND IS GONE (code review,
    // 2026-10-13). *** A draft is a complete copy; it does not get the update by itself, and sent as
    // it was it showed the maintainer the update's rows as removals. A draft without a base (seeded
    // before CRT kept one) is told so plainly; one with a base is offered to be brought up to date.
    // ###########################################################################################
    [Fact]
    public void A_draft_without_a_base_is_told_it_does_not_have_the_update()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report());
            string summary = TextOf(window, "SummaryText");

            Assert.Contains("does not have that update", summary);
            Assert.Contains("changes you made", summary);
            Assert.DoesNotContain("still applied", summary);
            Assert.Equal("I have looked at this", window.RebaseButtonTextForTests);
        });
    }

    [Fact]
    public void A_draft_with_a_base_is_offered_to_be_brought_up_to_date()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(ReportWithBase());

            Assert.Contains("Bring your draft up to date", TextOf(window, "SummaryText"));
            Assert.Equal("Bring my draft up to date", window.RebaseButtonTextForTests);
            Assert.True(window.RebaseButtonVisibleForTests);
        });
    }

    [Fact]
    public void Own_changes_the_update_and_the_rows_both_changed_are_listed_apart()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(ReportWithBase(
                own: [Row("U8", BoardRowChangeKind.Modified, "Description")],
                update: [Row("C1", BoardRowChangeKind.Added), Row("U9", BoardRowChangeKind.Deleted)],
                conflicts: [new DraftRebaseConflict("Components", "U8", "The update and your draft both changed this row. Your version is kept.")]));

            Assert.Equal("1 change of yours", TextOf(window, "OtherHeaderText"));
            Assert.Equal("2 changes in the published update", TextOf(window, "UpdateHeaderText"));
            Assert.Equal("1 row changed by both you and the update", TextOf(window, "AttentionHeaderText"));
            Assert.True(window.GetControl<StackPanel>("AttentionPanel").IsVisible);
            Assert.True(window.GetControl<StackPanel>("UpdatePanel").IsVisible);

            Assert.Equal("Added in the update.", window.UpdateRows[0].Explanation);
            Assert.Equal("Removed in the update.", window.UpdateRows[1].Explanation);
            Assert.Contains("Your version is kept", Assert.Single(window.AttentionRows).Explanation);
        });
    }

    // ###########################################################################################
    // *** NO LONGER HELD BACK BY THE TABLE (CRT 4.0 step 1, 2026-10-25). *** Bringing the draft up to
    // date runs under the draft held in memory: what it holds is saved first, then the rebase writes
    // the file and the draft is read back (DraftEdits.WriteFileDirectlyAsync - DraftSessionTests
    // cover that). When what it holds cannot be saved - the table's change, with the file written by
    // another program meanwhile - nothing is rebased, and the window says why.
    // ###########################################################################################
    [Fact]
    public async Task Bringing_up_to_date_changes_nothing_when_what_the_draft_holds_cannot_be_saved_first()
    {
        string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, ExcelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
        CachedWorkbooks.Write(workbook, new BoardData { RevisionDate = "2026-01-01" });

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, ExcelDataFile),
            new DraftMarker { BoardKey = ExcelDataFile, BaseRevision = "2026-May-12" });

        DraftSession session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, ExcelDataFile)!;
        Assert.True(session.Apply("The table's change", content => content with { Board = content.Board.WithRevisionDate("2026-02-02") }));
        Assert.True(DraftWorkbookStore.Edit(DraftManager.DraftsRoot, ExcelDataFile, board => board.WithRevisionDate("2026-03-03")));

        await UiTest.RunAsync(async () =>
        {
            DraftDriftWindow window = WindowFor(ReportWithBase());

            await window.RebaseAsync();

            Assert.False(window.BroughtUpToDate);
            Assert.True(window.RebaseButtonVisibleForTests);
            Assert.Equal(
                "Nothing was changed. " + DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.ChangedOnDisk),
                window.StatusTextForTests);
        });

        Assert.Equal("The table's change", session.UndoDescription);
        Assert.Equal("2026-03-03", DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, ExcelDataFile)!.RevisionDate);
    }

    // ------------------------------------------------------------------ The rows

    // The two consequential standings go in their own section, ahead of everything else - they are
    // the only ones a contributor may need to act on.
    // ###########################################################################################
    // *** THE "NEEDS ATTENTION" BAND IS GONE, AND THREE TESTS WENT WITH IT (Phase 6,
    // 2026-09-23). ***
    //
    // They pinned the two standings that only a MERGE could produce - a change to a row that
    // had since vanished officially, and an addition the official data had since made too -
    // plus the wording each got. BoardDraftApplier silently dropped the first and failed
    // closed on the second, and the band existed to surface outcomes the contributor could
    // not otherwise see.
    //
    // There is no merge now: the draft workbook IS the board, so neither outcome can occur
    // and neither can be reported. The tests are deleted rather than rewritten because the
    // behaviour they described no longer exists - this note is here so their absence reads as
    // a decision. What the panel showed is recorded in DraftChangeReport's own header,
    // including what was lost.
    // ###########################################################################################
    [Fact]
    public void Every_change_is_listed_in_the_one_panel()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report(rows: new[]
            {
                Row("U8"),
                Row("U9", BoardRowChangeKind.Added),
                Row("U19", BoardRowChangeKind.Deleted),
            }));

            Assert.Equal(3, window.OtherRows.Count);

            // The band for rows both sides changed shows only with such rows (a draft with a base).
            Assert.False(window.GetControl<StackPanel>("AttentionPanel").IsVisible);
        });
    }

    [Fact]
    public void The_heading_counts_the_changes_and_reads_in_the_singular_for_one()
    {
        UiTest.Run(() =>
        {
            // Without a base the rows are differences from the published data AS IT IS NOW - which
            // include the update's, turned round - so they are not called the contributor's.
            var window = WindowFor(Report(rows: new[] { Row("U8") }));

            Assert.True(window.GetControl<StackPanel>("OtherPanel").IsVisible);
            Assert.Equal("1 difference from the published data", TextOf(window, "OtherHeaderText"));

            Assert.Equal("1 change of yours", TextOf(WindowFor(ReportWithBase(own: [Row("U8")])), "OtherHeaderText"));
        });
    }

    // ###########################################################################################
    // A modified row NAMES the fields that differ - "U8 changed" is far less useful than
    // "U8: Description, Part number", and the information is free at the point the comparison
    // is made.
    // ###########################################################################################
    [Fact]
    public void A_modified_row_names_the_fields_that_changed()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report(rows: new[]
            {
                Row("U8", BoardRowChangeKind.Modified, "Description", "PartNumber"),
            }));

            string explanation = Assert.Single(window.OtherRows).Explanation;

            Assert.Contains("Description", explanation);
            Assert.Contains("PartNumber", explanation);
        });
    }

    // ###########################################################################################
    // *** A REMOVED ROW SAYS WHAT SUBMITTING WOULD DO. *** Deletion is the one change whose
    // consequence is not obvious from the row itself: the published board still has it, and
    // the contributor's submission is what would take it away.
    // ###########################################################################################
    [Fact]
    public void A_removed_row_says_the_submission_would_remove_it()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report(rows: new[] { Row("U9", BoardRowChangeKind.Deleted) }));

            Assert.Contains("would remove it", Assert.Single(window.OtherRows).Explanation);
        });
    }

    [Fact]
    public void A_board_with_no_drafted_rows_says_so_rather_than_showing_empty_lists()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report());

            Assert.True(window.GetControl<TextBlock>("NoRowsText").IsVisible);
            Assert.False(window.GetControl<StackPanel>("AttentionPanel").IsVisible);
            Assert.False(window.GetControl<StackPanel>("OtherPanel").IsVisible);
        });
    }

    // ------------------------------------------------------------------ The rebase action

    [Fact]
    public void The_dismiss_action_is_offered_only_when_there_is_a_warning_to_dismiss()
    {
        UiTest.Run(() =>
        {
            Assert.True(WindowFor(Report(state: DraftDriftState.OfficialIsNewer)).RebaseButtonVisibleForTests);
            Assert.False(WindowFor(Report(state: DraftDriftState.InSync)).RebaseButtonVisibleForTests);
            Assert.False(WindowFor(Report(state: DraftDriftState.Unknown)).RebaseButtonVisibleForTests);
        });
    }

    // ###########################################################################################
    // THE guarantee the whole action rests on: it writes ONE STRING and touches no board row.
    // The strategy doc explicitly forbids building a merge-conflict resolver here, so a rebase
    // that quietly dropped or rewrote a row would be both a bug and a broken promise.
    //
    // *** SINCE PHASE 6 THAT STRING LIVES IN THE MARKER, and the board is a real workbook. ***
    // So this writes a draft the way DraftSeeder would, rebases, and then re-reads BOTH - the
    // marker to prove the revision moved, and the workbook to prove the rows did not. Asserted
    // from disk rather than from an in-memory object, so what is pinned is what was persisted.
    // ###########################################################################################
    // ###########################################################################################
    // *** IT NO LONGER MOVES THE BASE REVISION (code review, 2026-10-13). *** The base revision is
    // what the submission tells the server; moving it made a stale draft claim to be made from the
    // newest data. "I have looked at this" now records only that the warning was seen.
    // ###########################################################################################
    [Fact]
    public void Dismissing_the_warning_keeps_the_base_revision_and_leaves_every_row_alone()
    {
        UiTest.Run(() =>
        {
            string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, ExcelDataFile);
            Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);

            var board = new BoardData();
            board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "my correction" });
            CachedWorkbooks.Write(workbook, board);

            DraftMarkerStore.Save(
                DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, ExcelDataFile),
                new DraftMarker { BoardKey = ExcelDataFile, BaseRevision = "2026-May-12" });

            var window = WindowFor(Report(rows: new[] { Row("U8") }));

            // The button is wired by Click, and a window never attached to a visual tree cannot
            // be clicked - so this drives the same handler the button invokes.
            window.RaiseRebaseForTests();

            DraftMarker? marker = DraftMarkerStore.Load(
                DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, ExcelDataFile));

            Assert.NotNull(marker);
            Assert.Equal("2026-May-12", marker!.BaseRevision);
            Assert.Equal("2026-August-21", marker.DriftSeenRevision);

            // And the contributor's own work is untouched - the promise the status line makes.
            BoardData? reloaded = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, ExcelDataFile);

            Assert.NotNull(reloaded);
            Assert.Equal("my correction", Assert.Single(reloaded!.Components).FriendlyName);
        });
    }

    [Fact]
    public void Dismissing_the_warning_hides_the_action_and_confirms_nothing_was_changed()
    {
        UiTest.Run(() =>
        {
            string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, ExcelDataFile);
            Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
            CachedWorkbooks.Write(workbook, new BoardData());

            DraftMarkerStore.Save(
                DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, ExcelDataFile),
                new DraftMarker { BoardKey = ExcelDataFile, BaseRevision = "2026-May-12" });

            var window = WindowFor(Report());
            window.RaiseRebaseForTests();

            Assert.False(window.RebaseButtonVisibleForTests);
            Assert.Contains("unchanged", window.StatusTextForTests);
        });
    }
}
