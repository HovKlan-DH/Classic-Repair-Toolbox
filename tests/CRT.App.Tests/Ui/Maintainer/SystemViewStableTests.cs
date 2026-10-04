using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// WHERE A SYSTEM IS (owner request, 2026-10-04: "I see this system here ... where does this sit now,
// as I do not think it is in BETA nor stable? Shouldn't there be somewhere a possibility to see what
// we actually do have in BETA or stable for this?") - the stage line under its name
// (SystemView.Stages.cs), and the BETA / Stable switch on Board data and Files (SystemView.Stable.cs).
// Drawn without a server: the detail through ShowDetailForTests, tables through their seams.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SystemViewStableTests
{
    private const string SystemId = "Commodore/C64/250407";

    private static readonly DateTimeOffset Decided = new(2026, 9, 21, 10, 0, 0, TimeSpan.Zero);

    private static SystemOverviewEntry System(bool inBeta = true, bool? inStable = true, string systemId = SystemViewStableTests.SystemId) =>
        new(systemId, "Commodore", "C64", systemId.Split('/')[2], inBeta, inStable, false, true,
            inBeta ? "2026-October-4" : null, inStable == true ? "2026-September-25" : null, null, 1);

    private static SystemDetailAnswer Detail(SystemOverviewEntry system, params SystemSubmissionEntry[] submissions) =>
        new(system, [], [], submissions);

    private static SubmissionRows Rows(string friendlyName) => new()
    {
        RevisionDate = "2026-September-25",
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = friendlyName, TechnicalNameOrValue = "906114-01" }]
    };

    private static SystemTableAnswer Table(string friendlyName, bool mayEdit, string? reason = null) =>
        new(SystemViewStableTests.SystemId, new string('f', 64), SystemViewStableTests.Rows(friendlyName), mayEdit, reason);

    private static bool Shown(Control root, string name) => root.FindControl<Control>(name)!.IsVisible;

    private static string FriendlyName(BoardTableEditor editor)
    {
        BoardTableSheet components = editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
        return components.Rows.Single().Cells[friendly].Text;
    }

    // ###########################################################################################
    // THE REPORTED SYSTEM: a new one whose only submission was turned down. The line says so, and
    // that neither BETA nor the stable source holds it - it used to read only "Not published yet".
    // ###########################################################################################
    [Fact]
    public void The_stage_line_says_where_a_system_turned_down_is()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.ShowDetailForTests(SystemViewStableTests.Detail(
                SystemViewStableTests.System(inBeta: false, inStable: false),
                new SystemSubmissionEntry(5, "c@example.com", "Open128", "rejected", SystemViewStableTests.Decided.AddDays(-1), SystemViewStableTests.Decided, "No.")));

            Assert.True(view.StagesShownForTests);
            Assert.Equal(
                [
                    $"Submitted: #5 Not accepted - decided {SubmissionReceiptPresenter.FormatDate(SystemViewStableTests.Decided)}",
                    "BETA: Not there",
                    "Stable: Not there"
                ],
                view.StagesForTests());
        });
    }

    [Fact]
    public void The_stage_line_names_the_BETA_and_stable_revisions_and_is_gone_with_no_system()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System()));

            Assert.Equal(["Submitted: No submissions", "BETA: Revision 2026-October-4", "Stable: Revision 2026-September-25"], view.StagesForTests());

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
            var view = new SystemView();
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System()));

            Assert.True(SystemViewStableTests.Shown(view, "TreeSwitchBar"));
            Assert.Equal("Data source:", view.FindControl<StackPanel>("TreeSwitchBar")!.Children.OfType<TextBlock>().Single().Text);

            await view.ShowSectionAsync(SystemSection.History);
            Assert.False(SystemViewStableTests.Shown(view, "TreeSwitchBar"));

            await view.ShowSectionAsync(SystemSection.Files);
            Assert.True(SystemViewStableTests.Shown(view, "TreeSwitchBar"));
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
            var view = new SystemView();
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System()));
            view.OpenTableForTests(SystemViewStableTests.Table("PLA", mayEdit: true));

            BoardTableSheet components = view.SystemTableForTests.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
            int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
            components.Rows.Single().Cells[friendly].Text = "PLA (82S100)";

            Assert.True(view.HasUnsavedTableEdits);

            string? asked = null;
            view.ReadStableTableOverrideForTests = systemId =>
            {
                asked = systemId;
                return Task.FromResult(ReviewApiResult<SystemTableAnswer>.Ok(SystemViewStableTests.Table("PLA (stable)", mayEdit: false, "Only a publish from BETA changes it.")));
            };

            await view.ShowTreeAsync(stable: true);

            Assert.Equal(SystemViewStableTests.SystemId, asked);
            Assert.True(view.ShowsStableForTests);
            Assert.True(SystemViewStableTests.Shown(view, "StableBoardPart"));
            Assert.False(SystemViewStableTests.Shown(view, "BetaBoardPart"));
            Assert.True(view.StableTableForTests.IsReadOnly);
            Assert.Equal("PLA (stable)", SystemViewStableTests.FriendlyName(view.StableTableForTests));

            // Its reason is under the switch, where BETA's table says what it is - never the table's
            // own line under its search box, which moved the table about between the two (owner
            // request, 2026-10-04).
            Assert.Equal("Only a publish from BETA changes it.", view.StableTableNoteForTests);
            Assert.False(view.StableTableForTests.FindControl<TextBlock>("StatusText")!.IsVisible);
            Assert.Contains("Selected", view.FindControl<Button>("StableTreeButton")!.Classes);

            await view.ShowTreeAsync(stable: false);

            Assert.True(SystemViewStableTests.Shown(view, "BetaBoardPart"));
            Assert.Equal(SystemSections.OpenedMessage(SystemViewStableTests.Table("PLA", mayEdit: true)), view.TableNoteForTests);
            Assert.False(view.SystemTableForTests.FindControl<TextBlock>("StatusText")!.IsVisible);
            Assert.True(view.HasUnsavedTableEdits);
            Assert.Equal("PLA (82S100)", SystemViewStableTests.FriendlyName(view.SystemTableForTests));
        });
    }

    // ###########################################################################################
    // A half the system is not in cannot be chosen; a system only the stable source holds shows it
    // without being asked; and the pick stays for the next system that has it.
    // ###########################################################################################
    [Fact]
    public async Task The_switch_offers_only_where_the_system_is_and_the_pick_stays_for_the_next_system()
    {
        await UiTest.RunAsync(async () =>
        {
            var view = new SystemView();
            view.ReadStableTableOverrideForTests = _ =>
                Task.FromResult(ReviewApiResult<SystemTableAnswer>.Ok(SystemViewStableTests.Table("PLA", mayEdit: false)));

            // Not in the stable source: BETA, and Stable cannot be chosen.
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System(inStable: false)));

            Assert.False(view.ShowsStableForTests);
            Assert.False(view.FindControl<Button>("StableTreeButton")!.IsEnabled);
            Assert.True(view.FindControl<Button>("BetaTreeButton")!.IsEnabled);

            // Only in the stable source: shown at once, and BETA cannot be chosen.
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System(inBeta: false)));

            Assert.True(view.ShowsStableForTests);
            Assert.False(view.FindControl<Button>("BetaTreeButton")!.IsEnabled);

            // Stable picked on a system in both stays picked for the next one in both.
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System()));
            await view.ShowTreeAsync(stable: true);
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System(systemId: "Commodore/C128/310378")));

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
            var view = new SystemView();
            view.ShowDetailForTests(SystemViewStableTests.Detail(SystemViewStableTests.System()));
            view.ReadStableTableOverrideForTests = _ => Task.FromResult(ReviewApiResult<SystemTableAnswer>.Ok(
                SystemViewStableTests.Table("PLA (BETA)", mayEdit: true) with { BetaDataUrl = "https://example.org/beta" }));

            await view.ShowTreeAsync(stable: true);

            TextBlock message = view.FindControl<TextBlock>("StableTableMessageText")!;
            Assert.True(message.IsVisible);
            Assert.Equal(SystemSections.StableNeedsNewerServer, message.Text);
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
            Assert.Same(tab.SystemDetailForTests.StableTableForTests, tab.TableEditorsForSharedChoices[2]);
        });
    }
}
