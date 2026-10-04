using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// A SYSTEM'S TABLE AND FILES FOLLOW BETA (owner report, 2026-10-04: "I did push a change through the
// system, fixing the very much highlighted issue in this one ... then it continues to show the error
// in the "Maintainer" tab on "Systems" ... After an app restart ... it did not show as an error any
// more").
//
// The Systems screen read a system's Board data and Files once, and kept them until another system
// was chosen. The system's row now carries BETA's content hash, and every time the row is read again
// (SystemView.ShowSystemAsync -> CatchUpWithBetaAsync) a table or a file list read before BETA moved
// is read again - never a table holding a change not published yet. Drawn without a server: the
// table through OpenTableForTests and ReadTableOverrideForTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SystemViewBetaCatchUpTests
{
    private const string SystemId = "Commodore/C64/250407";

    private static SystemOverviewEntry System(string? hash, bool awaiting = false) =>
        new(SystemId, "Commodore", "C64", "250407", true, true, awaiting, true, "2026-October-3", null, null, 1, BetaContentHash: hash);

    // The owner's board before the fix: a highlight on "Board layout" naming "hest", which is in no
    // row of the Components sheet - the warning above the table.
    private static SubmissionRows Before() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Board layout", SchematicImageFile = "a/board.png" }],
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" }],
        ComponentHighlights =
        [
            new ComponentHighlightEntry { SchematicName = "Board layout", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" },
            new ComponentHighlightEntry { SchematicName = "Board layout", BoardLabel = "hest", X = "9", Y = "9", Width = "5", Height = "5" }
        ]
    };

    // ...and after it: the "hest" highlight gone.
    private static SubmissionRows After() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Board layout", SchematicImageFile = "a/board.png" }],
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" }],
        ComponentHighlights =
        [
            new ComponentHighlightEntry { SchematicName = "Board layout", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" }
        ]
    };

    private static SystemTableAnswer Table(SubmissionRows rows, bool mayEdit = true) =>
        new(SystemId, new string('f', 64), rows, mayEdit, mayEdit ? null : "It waits under \"Queue: Awaiting push from BETA to stable\".");

    private static void Show(SystemView view, SystemOverviewEntry system) =>
        view.ShowDetailForTests(new SystemDetailAnswer(system, [], [], []));

    // The system's row read again - as ShowSystemAsync does, the detail drawn first.
    private static Task ReadAgain(SystemView view, SystemOverviewEntry now)
    {
        Show(view, now);
        return view.CatchUpWithBetaForTests(now);
    }

    private static TextBlock OutsideProblems(SystemView view) =>
        view.SystemTableForTests.FindControl<TextBlock>("OutsideProblemsText")!;

    // The line above the table, under the BETA / Stable switch (owner request, 2026-10-04: it was
    // the table's own line under its search box).
    private static string Status(SystemView view) => view.TableNoteForTests;

    // A view whose table was read at `readAt`, and a count of the tables read since.
    private static (SystemView View, Func<int> Reads) WithTable(SystemOverviewEntry readAt, SubmissionRows rows, SubmissionRows next, bool nextMayEdit = true)
    {
        var view = new SystemView();
        Show(view, readAt);
        view.OpenTableForTests(Table(rows), readAt);

        int reads = 0;
        view.ReadTableOverrideForTests = _ =>
        {
            reads++;
            return Task.FromResult(ReviewApiResult<SystemTableAnswer>.Ok(Table(next, nextMayEdit)));
        };

        return (view, () => reads);
    }

    // The owner's case, end to end: the warning about "hest" goes once BETA no longer has it.
    [Fact]
    public async Task A_table_read_before_BETA_moved_is_read_again_and_a_fixed_warning_goes()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, Func<int> reads) = WithTable(System("aaa"), Before(), After());

            Assert.True(OutsideProblems(view).IsVisible);
            Assert.Contains("hest", OutsideProblems(view).Text, StringComparison.Ordinal);

            await ReadAgain(view, System("bbb"));

            Assert.Equal(1, reads());
            Assert.False(OutsideProblems(view).IsVisible);
        });
    }

    // Read again only when it moved - not on every check of the system.
    [Fact]
    public async Task A_table_BETA_has_not_moved_past_is_not_read_again()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, Func<int> reads) = WithTable(System("aaa"), Before(), After());

            await ReadAgain(view, System("aaa") with { ViewsLast30Days = 40 });

            Assert.Equal(0, reads());
            Assert.True(OutsideProblems(view).IsVisible);
        });
    }

    // Read again on the sheet it was on - the maintainer is looking at it.
    [Fact]
    public async Task A_table_read_again_stays_on_the_sheet_it_was_on()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, _) = WithTable(System("aaa"), Before(), After());
            BoardTableEditor editor = view.SystemTableForTests;

            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetCredits)!);
            await ReadAgain(view, System("bbb"));

            Assert.Equal(BoardWorkbookSchema.SheetCredits, editor.CurrentSheet!.Name);
        });
    }

    // ###########################################################################################
    // A promotion leaves BETA as it is but lets the table be changed again: read while the system
    // waited under BETA > Stable, it was read-only with that reason, and stayed so after it no
    // longer waited.
    // ###########################################################################################
    [Fact]
    public async Task A_table_read_while_the_system_waited_is_read_again_once_it_no_longer_does()
    {
        await UiTest.RunAsync(async () =>
        {
            SystemOverviewEntry waiting = System("aaa", awaiting: true);
            var view = new SystemView();
            Show(view, waiting);
            view.OpenTableForTests(Table(Before(), mayEdit: false), waiting);
            view.SystemTableForTests.IsReadOnly = true;
            view.ReadTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<SystemTableAnswer>.Ok(Table(Before(), mayEdit: true)));

            await ReadAgain(view, System("aaa", awaiting: false));

            Assert.False(view.SystemTableForTests.IsReadOnly);
        });
    }

    // ###########################################################################################
    // *** NEVER UNDER A CHANGE NOT PUBLISHED YET *** - reading it again would throw the change away.
    // It stays, and says it can no longer go (the server would refuse it). Undone, the next look at
    // Board data reads BETA as it is.
    // ###########################################################################################
    [Fact]
    public async Task A_table_holding_a_change_is_not_read_again_but_says_the_change_can_no_longer_go()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, Func<int> reads) = WithTable(System("aaa"), Before(), After());
            BoardTableDocument document = view.SystemTableForTests.CommitAndGetDocument()!;
            BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            components.Rows.Single().Cells[BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName)].Text = "Mine";

            await ReadAgain(view, System("bbb"));

            Assert.Equal(0, reads());
            Assert.True(view.HasUnsavedTableEdits);
            Assert.Equal(SystemSections.BetaMovedUnderChange, Status(view));

            document.History.Undo();
            Assert.False(view.HasUnsavedTableEdits);

            await view.ShowSectionAsync(SystemSection.BoardData);

            Assert.Equal(1, reads());
            Assert.False(OutsideProblems(view).IsVisible);
        });
    }

    // ###########################################################################################
    // A table read just after a publish of its own (thisTableReadAt null) takes the next reading as
    // the state it was read at - reading it again would only replace the line saying what the publish
    // did. A later move still reads it again.
    // ###########################################################################################
    [Fact]
    public async Task A_table_read_after_its_own_publish_is_not_read_again_for_that_publish()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, Func<int> reads) = WithTable(System("aaa"), Before(), After());
            view.OpenTableForTests(Table(After()), readAt: null);

            await ReadAgain(view, System("bbb"));
            Assert.Equal(0, reads());

            await ReadAgain(view, System("ccc"));
            Assert.Equal(1, reads());
        });
    }

    // A board BETA no longer holds (pushed back out of it) is not left on screen as if it were.
    [Fact]
    public async Task A_table_whose_board_left_BETA_is_closed()
    {
        await UiTest.RunAsync(async () =>
        {
            (SystemView view, _) = WithTable(System("aaa"), Before(), After());
            view.ReadTableOverrideForTests = _ =>
                Task.FromResult(ReviewApiResult<SystemTableAnswer>.Failed(ReviewApiFailure.NotFound, "BETA does not hold this system."));

            await ReadAgain(view, System(null) with { InBeta = false });

            Assert.False(view.SystemTableForTests.HasTable);
        });
    }

    // ###########################################################################################
    // The Files view follows the board alone: a list not on screen is dropped when BETA moves, to be
    // read when next shown - and kept through a promotion, which moves nothing in BETA.
    // ###########################################################################################
    [Fact]
    public async Task A_file_list_not_on_screen_is_dropped_when_BETA_moves_and_kept_when_it_does_not()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new SystemView();
            SystemOverviewEntry readAt = System("aaa", awaiting: true);
            Show(view, readAt);
            view.ShowFilesForTests(new SystemFilesAnswer(SystemId, []), readAt);

            await ReadAgain(view, System("aaa", awaiting: false));
            Assert.True(view.HoldsFilesForTests(SystemId));

            await ReadAgain(view, System("bbb"));
            Assert.False(view.HoldsFilesForTests(SystemId));
        });
    }
}
