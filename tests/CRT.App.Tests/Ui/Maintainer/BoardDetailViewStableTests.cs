using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// WHERE A BOARD IS (owner request, 2026-10-04: "I see this system here ... where does this sit now,
// as I do not think it is in BETA nor stable? Shouldn't there be somewhere a possibility to see what
// we actually do have in BETA or stable for this?") - the stage line under its name
// (BoardDetailView.Stages.cs), and the BETA / Stable switch on Board data and Files (BoardDetailView.Stable.cs).
// Drawn without a server: the detail through ShowDetailForTests, tables through their seams.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardDetailViewStableTests
{
    private const string BoardId = "Commodore/C64/250407";

    private static readonly DateTimeOffset Decided = new(2026, 9, 21, 10, 0, 0, TimeSpan.Zero);

    private static BoardOverviewEntry Board(bool inBeta = true, bool? inStable = true, string boardId = BoardDetailViewStableTests.BoardId) =>
        new(boardId, "Commodore", "C64", boardId.Split('/')[2], inBeta, inStable, false, true,
            inBeta ? "2026-October-4" : null, inStable == true ? "2026-September-25" : null, null, 1);

    private static BoardDetailAnswer Detail(BoardOverviewEntry board, params BoardSubmissionEntry[] submissions) =>
        new(board, [], [], submissions);

    private static SubmissionRows Rows(string friendlyName) => new()
    {
        RevisionDate = "2026-September-25",
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = friendlyName, TechnicalNameOrValue = "906114-01" }]
    };

    private static BoardTableAnswer Table(string friendlyName, bool mayEdit, string? reason = null) =>
        new(BoardDetailViewStableTests.BoardId, new string('f', 64), BoardDetailViewStableTests.Rows(friendlyName), mayEdit, reason);

    private static bool Shown(Control root, string name) => root.FindControl<Control>(name)!.IsVisible;

    private static string FriendlyName(BoardTableEditor editor)
    {
        BoardTableSheet components = editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
        return components.Rows.Single().Cells[friendly].Text;
    }

    // ###########################################################################################
    // THE REPORTED BOARD: a new one whose only submission was turned down. The line says so, and
    // that neither BETA nor the stable source holds it - it used to read only "Not published yet".
    // ###########################################################################################
    [Fact]
    public void The_stage_line_says_where_a_board_turned_down_is()
    {
        UiTest.Run(() =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(
                BoardDetailViewStableTests.Board(inBeta: false, inStable: false),
                new BoardSubmissionEntry(5, "c@example.com", "Open128", "rejected", BoardDetailViewStableTests.Decided.AddDays(-1), BoardDetailViewStableTests.Decided, "No.")));

            Assert.True(view.StagesShownForTests);
            Assert.Equal(
                [
                    $"1  Submitted: #5 Not accepted - decided {SubmissionReceiptPresenter.FormatDate(BoardDetailViewStableTests.Decided)}",
                    "2  BETA: Not there",
                    "3  Stable: Not there"
                ],
                view.StagesForTests());
        });
    }

    [Fact]
    public void The_stage_line_names_the_BETA_and_stable_revisions_and_is_gone_with_no_board()
    {
        UiTest.Run(() =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board()));

            Assert.Equal(["1  Submitted: No submissions", "2  BETA: Revision 2026-October-4", "3  Stable: Revision 2026-September-25"], view.StagesForTests());

            view.Clear();
            Assert.False(view.StagesShownForTests);
        });
    }

    // ###########################################################################################
    // The switch belongs to Board data and Files, and only there: the other views carry nothing of
    // either tree.
    // ###########################################################################################
    [Fact]
    public async Task The_switch_shows_with_Board_data_and_Files_only()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board()));

            Assert.True(BoardDetailViewStableTests.Shown(view, "TreeSwitchBar"));
            Assert.Equal("Data source:", view.FindControl<StackPanel>("TreeSwitchBar")!.Children.OfType<TextBlock>().Single().Text);

            await view.ShowSectionAsync(BoardSection.History);
            Assert.False(BoardDetailViewStableTests.Shown(view, "TreeSwitchBar"));

            await view.ShowSectionAsync(BoardSection.Files);
            Assert.True(BoardDetailViewStableTests.Shown(view, "TreeSwitchBar"));
        });
    }

    // ###########################################################################################
    // *** SWITCHING NEVER TOUCHES A CHANGE IN BETA'S TABLE. *** The stable source has a read-only
    // table of its own: chosen, it shows the stable board with the server's reason; back on BETA,
    // the change typed there is still there, still unsaved.
    // ###########################################################################################
    [Fact]
    public async Task The_stable_table_is_its_own_read_only_table_and_BETAs_change_survives_switching()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board()));
            view.OpenTableForTests(BoardDetailViewStableTests.Table("PLA", mayEdit: true));

            BoardTableSheet components = view.BoardTableForTests.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
            components.Rows.Single().Cells[friendly].Text = "PLA (82S100)";

            Assert.True(view.HasUnsavedTableEdits);

            string? asked = null;
            view.ReadStableTableOverrideForTests = boardId =>
            {
                asked = boardId;
                return Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewStableTests.Table("PLA (stable)", mayEdit: false, "Only a publish from BETA changes it.")));
            };

            await view.ShowTreeAsync(stable: true);

            Assert.Equal(BoardDetailViewStableTests.BoardId, asked);
            Assert.True(view.ShowsStableForTests);
            Assert.True(BoardDetailViewStableTests.Shown(view, "StableBoardPart"));
            Assert.False(BoardDetailViewStableTests.Shown(view, "BetaBoardPart"));
            Assert.True(view.StableTableForTests.IsReadOnly);
            Assert.Equal("PLA (stable)", BoardDetailViewStableTests.FriendlyName(view.StableTableForTests));

            // Its reason is under the switch, where BETA's table says what it is - never the table's
            // own line under its search box, which moved the table about between the two (owner
            // request, 2026-10-04).
            Assert.Equal("Only a publish from BETA changes it.", view.StableTableNoteForTests);
            Assert.False(view.StableTableForTests.FindControl<TextBlock>("StatusText")!.IsVisible);
            Assert.Contains("Selected", view.FindControl<Button>("StableTreeButton")!.Classes);

            await view.ShowTreeAsync(stable: false);

            Assert.True(BoardDetailViewStableTests.Shown(view, "BetaBoardPart"));
            Assert.Equal(BoardSections.OpenedMessage(BoardDetailViewStableTests.Table("PLA", mayEdit: true)), view.TableNoteForTests);
            Assert.False(view.BoardTableForTests.FindControl<TextBlock>("StatusText")!.IsVisible);
            Assert.True(view.HasUnsavedTableEdits);
            Assert.Equal("PLA (82S100)", BoardDetailViewStableTests.FriendlyName(view.BoardTableForTests));
        });
    }

    // ###########################################################################################
    // A half the board is not in cannot be chosen; a board only the stable source holds shows it
    // without being asked; and the pick stays for the next board that has it.
    // ###########################################################################################
    [Fact]
    public async Task The_switch_offers_only_where_the_board_is_and_the_pick_stays_for_the_next_board()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            view.ReadStableTableOverrideForTests = _ =>
                Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(BoardDetailViewStableTests.Table("PLA", mayEdit: false)));

            // Not in the stable source: BETA, and Stable cannot be chosen.
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board(inStable: false)));

            Assert.False(view.ShowsStableForTests);
            Assert.False(view.FindControl<Button>("StableTreeButton")!.IsEnabled);
            Assert.True(view.FindControl<Button>("BetaTreeButton")!.IsEnabled);

            // Only in the stable source: shown at once, and BETA cannot be chosen.
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board(inBeta: false)));

            Assert.True(view.ShowsStableForTests);
            Assert.False(view.FindControl<Button>("BetaTreeButton")!.IsEnabled);

            // Stable picked on a board in both stays picked for the next one in both.
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board()));
            await view.ShowTreeAsync(stable: true);
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board(boardId: "Commodore/C128/310378")));

            Assert.True(view.ShowsStableForTests);
        });
    }

    // A server older than 4.5.0 answers a stable request with BETA's board: it is not drawn as the
    // stable source's - the view says the server needs updating.
    [Fact]
    public async Task An_older_servers_answer_is_not_shown_as_the_stable_source()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new BoardDetailView();
            view.ShowDetailForTests(BoardDetailViewStableTests.Detail(BoardDetailViewStableTests.Board()));
            view.ReadStableTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<BoardTableAnswer>.Ok(
                BoardDetailViewStableTests.Table("PLA (BETA)", mayEdit: true) with { BetaDataUrl = "https://example.org/beta" }));

            await view.ShowTreeAsync(stable: true);

            TextBlock message = view.FindControl<TextBlock>("StableTableMessageText")!;
            Assert.True(message.IsVisible);
            Assert.Equal(BoardSections.StableNeedsNewerServer, message.Text);
            Assert.Null(view.StableTableForTests.CommitAndGetDocument());
        });
    }

    // The picked pills are shared with the stable table too - one table, in three places of the
    // Maintainer tab.
    [Fact]
    public void The_stable_table_shares_the_tables_choices()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();

            Assert.Equal(3, tab.TableEditorsForSharedChoices.Count);
            Assert.Same(tab.BoardDetailForTests.StableTableForTests, tab.TableEditorsForSharedChoices[2]);
        });
    }
}
