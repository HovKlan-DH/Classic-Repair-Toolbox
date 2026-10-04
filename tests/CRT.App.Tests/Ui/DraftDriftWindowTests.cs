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

    // Without this, the warning reads as "your work may be lost", which is false.
    [Fact]
    public void The_summary_reassures_that_the_edits_are_still_applied()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report());

            Assert.Contains("still applied", TextOf(window, "SummaryText"));
        });
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

            // The band that held merge consequences stays hidden - nothing can populate it.
            Assert.False(window.GetControl<StackPanel>("AttentionPanel").IsVisible);
        });
    }

    [Fact]
    public void The_heading_counts_the_changes_and_reads_in_the_singular_for_one()
    {
        UiTest.Run(() =>
        {
            var window = WindowFor(Report(rows: new[] { Row("U8") }));

            Assert.True(window.GetControl<StackPanel>("OtherPanel").IsVisible);
            Assert.Equal("1 change of yours", TextOf(window, "OtherHeaderText"));
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
    [Fact]
    public void Dismissing_the_warning_moves_the_base_revision_and_leaves_every_row_alone()
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
                new DraftMarker { SystemKey = ExcelDataFile, BaseRevision = "2026-May-12" });

            var window = WindowFor(Report(rows: new[] { Row("U8") }));

            // The button is wired by Click, and a window never attached to a visual tree cannot
            // be clicked - so this drives the same handler the button invokes.
            window.RaiseRebaseForTests();

            DraftMarker? marker = DraftMarkerStore.Load(
                DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, ExcelDataFile));

            Assert.NotNull(marker);
            Assert.Equal("2026-August-21", marker!.BaseRevision);

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
                new DraftMarker { SystemKey = ExcelDataFile, BaseRevision = "2026-May-12" });

            var window = WindowFor(Report());
            window.RaiseRebaseForTests();

            Assert.False(window.RebaseButtonVisibleForTests);
            Assert.Contains("unchanged", window.StatusTextForTests);
        });
    }
}
