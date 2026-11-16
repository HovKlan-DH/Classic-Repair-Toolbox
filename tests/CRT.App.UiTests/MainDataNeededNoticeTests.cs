using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A board checked again on the "Configuration" tab whose data is not here - a data check leaves an unchecked board out
// (owner request, 2026-11-14: "If then clicked/checked, and it does not have the data, then it should inform about this in
// the bottom-right corner with an info-panel, so the user should take a data refresh"). Main.DataNeededNotice.cs, driven as
// the tab's checkbox drives it: the setting changed, then Main.ApplyCatalogueVisibility with its key.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainDataNeededNoticeTests : IDisposable
{
    private static readonly MainFiles.Board Here = new("Commodore/C64/250407", "Commodore 64", "250407");
    private static readonly MainFiles.Board NotHere = new("Commodore/C64/250469", "Commodore 64", "250469");

    private readonly TempWorkspace thisWorkspace = new();

    public MainDataNeededNoticeTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));

        string dataRoot = this.thisWorkspace.Path_("Data");
        DataManager.LoadFrom(dataRoot, MainFiles.Write(dataRoot, MainDataNeededNoticeTests.Here, MainDataNeededNoticeTests.NotHere));
        StoredFiles.WriteBoard(DraftBoardSource.PublishedPathOf(dataRoot, MainDataNeededNoticeTests.Here.Identity), new BoardData());
    }

    public void Dispose()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.json");
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static async Task CheckAsync(CRT.Main main, string key, bool isChecked)
    {
        UserSettings.SetCatalogueKeyChecked(key, isChecked);
        main.ApplyCatalogueVisibility(key);
        await main.LookForDataNeededForTests;
        Dispatcher.UIThread.RunJobs();
    }

    [Fact]
    public async Task A_board_checked_again_without_its_data_is_said_and_unchecked_again_it_goes()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            Border card = main.GetControl<Border>("DataNeededBanner");

            await CheckAsync(main, "Commodore 64|250469", false);
            Assert.False(card.IsVisible);

            await CheckAsync(main, "Commodore 64|250469", true);

            Assert.True(card.IsVisible);
            Assert.Equal(["Commodore 64", "250469"], BoldRuns.BoldPiecesOf(main.GetControl<TextBlock>("DataNeededBannerText")));

            await CheckAsync(main, "Commodore 64|250469", false);
            Assert.False(card.IsVisible);

            main.Close();
        });
    }

    // ###########################################################################################
    // *** "Check for new data" pressed while a check runs is followed by one more check (code review, 2026-11-16). *** The check
    // running worked out what it takes before the board was checked, so it brings none of it: the card went and nothing more
    // happened. Now one more check runs as the running one ends - and only one, however often it was pressed.
    // ###########################################################################################
    [Fact]
    public async Task Check_for_new_data_while_a_check_runs_runs_one_more_once_it_ends()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            int checksRun = 0;
            main.DataCheckOverrideForTests = () =>
            {
                checksRun++;
                return Task.CompletedTask;
            };

            await CheckAsync(main, "Commodore 64|250469", false);
            await CheckAsync(main, "Commodore 64|250469", true);
            Assert.True(main.GetControl<Border>("DataNeededBanner").IsVisible);

            main.StartDataSyncStatusIconSpin();
            Button check = main.GetControl<Button>("DataNeededBannerCheckButton");
            check.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(0, checksRun);
            Assert.False(main.GetControl<Border>("DataNeededBanner").IsVisible);

            main.StopDataSyncStatusIconSpin();
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(1, checksRun);

            // Nothing asked for meanwhile: the next check ending runs none.
            main.StartDataSyncStatusIconSpin();
            main.StopDataSyncStatusIconSpin();
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(1, checksRun);

            main.Close();
        });
    }

    // ###########################################################################################
    // *** A look that ends after a data check started says nothing (code review, 2026-11-16). *** The board ticked and the data
    // icon clicked at once: the check, begun after the tick, brings the board's data - and the look, ending after it began, put
    // the card back, where it stayed after the data had come.
    // ###########################################################################################
    [Fact]
    public async Task A_look_ending_after_a_data_check_began_puts_no_card_back()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            Border card = main.GetControl<Border>("DataNeededBanner");

            await CheckAsync(main, "Commodore 64|250469", false);

            UserSettings.SetCatalogueKeyChecked("Commodore 64|250469", true);
            main.ApplyCatalogueVisibility("Commodore 64|250469");

            // A data check starts while the look is still under way (CheckForDataUpdatesNowAsync lets the card go first).
            main.HideDataNeededNotice();

            await main.LookForDataNeededForTests;
            Dispatcher.UIThread.RunJobs();

            Assert.False(card.IsVisible);

            main.Close();
        });
    }

    // The hardware checked again: only its boards whose data is not here are named - the one with its data is not.
    [Fact]
    public async Task A_hardware_checked_again_names_only_its_boards_without_data_and_closing_lets_go()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            Border card = main.GetControl<Border>("DataNeededBanner");

            await CheckAsync(main, "Commodore 64", false);
            await CheckAsync(main, "Commodore 64", true);

            Assert.True(card.IsVisible);
            Assert.Contains("250469", NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataNeededBannerText")), StringComparison.Ordinal);
            Assert.DoesNotContain("250407", NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataNeededBannerText")), StringComparison.Ordinal);

            main.GetControl<Button>("DataNeededBannerDismissButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Assert.False(card.IsVisible);

            main.Close();
        });
    }
}
