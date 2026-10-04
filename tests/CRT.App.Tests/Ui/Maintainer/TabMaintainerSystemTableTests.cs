using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// A system's table on the Systems screen as the WHOLE TAB handles it (2026-10-03): choosing another
// system asks about a change not published yet - Cancel keeps the system and its change - the tab
// counts that change as unsaved (quitting CRT and the tab's visibility read it), and a change the
// server made into a submission but could not publish opens it under Contributor Submissions, where
// it waits. (A published change goes straight to BETA - SystemViewSectionsTests.)
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerSystemTableTests
{
    private const string First = "Commodore/C64/250407";
    private const string Second = "Commodore/C128/310378";

    private static SystemOverviewEntry System(string systemId) =>
        new(systemId, "Commodore", systemId.Split('/')[1], systemId.Split('/')[2], true, true, false, true, null, null, null, 1);

    private static ReviewQueueRow Row(long id, string system) =>
        new(id, system, "pending", $"Submission {id}.", "c@example.com", DateTimeOffset.UtcNow.AddDays(-1), false, false, true);

    // Two systems listed, the first on screen with its table open on BETA's board and one change in it.
    private static (TabMaintainer Main, ListBox List) WithChangedTable()
    {
        var main = new TabMaintainer();

        main.ApplySystemsListAsync(
            new SystemOverviewAnswer([TabMaintainerSystemTableTests.System(First), TabMaintainerSystemTableTests.System(Second)]),
            background: true).GetAwaiter().GetResult();

        SystemView view = main.SystemDetailForTests;
        view.ShowDetailForTests(new SystemDetailAnswer(TabMaintainerSystemTableTests.System(First), [], [], []));

        var rows = new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA" }] };
        view.OpenTableForTests(new SystemTableAnswer(First, "f", rows, MayEdit: true));

        BoardTableSheet components = view.SystemTableForTests.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
        components.Rows.Single().Cells[friendly].Text = "CPU 6510";

        return (main, main.FindControl<ListBox>("SystemsList")!);
    }

    private static ListBoxItem ItemOf(ListBox list, string systemId) =>
        ((IEnumerable<ListBoxItem>)list.ItemsSource!).Single(item => ((SystemOverviewEntry)item.Tag!).SystemId == systemId);

    [Fact]
    public void A_change_not_sent_in_a_systems_table_counts_as_unsaved_for_the_whole_tab()
    {
        UiTest.Run(() =>
        {
            (TabMaintainer main, _) = TabMaintainerSystemTableTests.WithChangedTable();

            Assert.True(main.SystemDetailForTests.HasUnsavedTableEdits);
            Assert.True(main.HasUnsavedTableEdits);
        });
    }

    // Cancel: the system stays chosen, on screen, with its change.
    [Fact]
    public async Task Choosing_another_system_asks_first_and_Cancel_keeps_the_system_and_its_change()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer main, ListBox list) = TabMaintainerSystemTableTests.WithChangedTable();
            var asked = new List<UnsavedTableEditsPrompt>();

            main.SystemDetailForTests.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return UnsavedTableEditsChoice.Cancel;
            };

            list.SelectedItem = TabMaintainerSystemTableTests.ItemOf(list, Second);
            Dispatcher.UIThread.RunJobs();
            await Task.Yield();

            Assert.Equal([UnsavedTableEditsPrompt.LeavingSystem], asked);
            Assert.Equal(First, main.SelectedSystemForTests?.SystemId);
            Assert.Equal(First, main.SystemDetailForTests.ShownSystem?.SystemId);
            Assert.True(main.SystemDetailForTests.HasUnsavedTableEdits);
        });
    }

    // Discard: the other system is shown, and nothing of the first one's table stays.
    [Fact]
    public async Task Choosing_another_system_and_discarding_shows_it_with_no_table_left()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer main, ListBox list) = TabMaintainerSystemTableTests.WithChangedTable();

            main.SystemDetailForTests.UnsavedTableEditsAnswerForTests = _ => UnsavedTableEditsChoice.Discard;

            list.SelectedItem = TabMaintainerSystemTableTests.ItemOf(list, Second);
            Dispatcher.UIThread.RunJobs();
            await Task.Yield();

            Assert.Equal(Second, main.SystemDetailForTests.ShownSystem?.SystemId);
            Assert.False(main.SystemDetailForTests.HasUnsavedTableEdits);
            Assert.False(main.HasUnsavedTableEdits);
        });
    }

    // ###########################################################################################
    // A change made into a submission but NOT published: it is opened under Contributor Submissions -
    // even with another submission chosen there already, which the queue's own rule would otherwise keep.
    // ###########################################################################################
    [Fact]
    public async Task A_change_made_but_not_published_opens_its_submission_under_Contributor_Submissions()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();

            main.ApplyQueueResponse(new ReviewQueueResponse(
                CanPublish: true,
                Submissions: [TabMaintainerSystemTableTests.Row(41, Second), TabMaintainerSystemTableTests.Row(57, First)],
                IsAdministrator: true));

            await main.ShowModeAsync(MaintainerMode.Review);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(41, main.SelectedQueueRowForTests?.Id);

            await main.ShowModeAsync(MaintainerMode.Systems);
            await main.OpenSentSubmissionAsync(57);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Equal(57, main.SelectedQueueRowForTests?.Id);
        });
    }
}
