using System;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DataSyncStatusTip - the sync icon's tooltip, which is where a data check says how far it got
// since it stopped putting a notice up above the tabs (owner report, 2026-10-11: the notice came
// and went around every check and the whole window moved each time).
// ###########################################################################################
public sealed class DataSyncStatusTipTests
{
    private static readonly DateTime Afternoon = new(2026, 10, 11, 14, 5, 0);

    // While a check runs the tooltip is the old notice's two lines under a heading: what it is
    // doing, and the file it is on - read by whoever points at the spinning icon.
    [Fact]
    public void A_running_check_says_what_it_is_doing_and_the_file_it_is_on()
    {
        string tip = DataSyncStatusTip.Build(
            isChecking: true,
            isEnabled: true,
            isHovering: true,
            sourceLabel: "online source",
            progress: "Downloading file [3] of [10] from the online source:",
            file: "Commodore/C64/250407/250407.xlsx",
            lastCheck: "Last check at 13:00: all data files were up to date");

        Assert.Equal(
            "Checking data from the online source - please wait...\n" +
            "Downloading file [3] of [10] from the online source:\n" +
            "Commodore/C64/250407/250407.xlsx",
            tip);
    }

    // Before the first status arrives there is no progress to show - no blank lines either. A
    // status that arrived with a line break of its own stays one line.
    [Fact]
    public void A_running_check_leaves_out_what_it_has_not_said_yet_and_keeps_each_part_to_one_line()
    {
        Assert.Equal(
            "Checking data from the BETA source - please wait...",
            DataSyncStatusTip.Build(true, true, false, "BETA source", "", "   ", null));

        Assert.Equal(
            "Checking data from the online source - please wait...\nFetching online file manifest...",
            DataSyncStatusTip.Build(true, true, false, "online source", "Fetching online\r\nfile manifest...  ", "", null));
    }

    // Not checking: what a click does - the three wordings the icon has always had, by setting
    // and hover.
    [Theory]
    [InlineData(true, false, "Data update is enabled. Click to refresh data now")]
    [InlineData(true, true, "Data update is enabled. Click to refresh data now")]
    [InlineData(false, true, "Data update at launch is disabled. Click to run a manual refresh now")]
    [InlineData(false, false, "Data update at launch is disabled. Hover to show manual refresh")]
    public void An_idle_icon_says_what_a_click_does(bool isEnabled, bool isHovering, string expected)
    {
        Assert.Equal(
            expected,
            DataSyncStatusTip.Build(false, isEnabled, isHovering, "online source", "stale progress", "stale/file.png", null));
    }

    // A check that changed nothing puts nothing up, so the tooltip is where its outcome is said -
    // under the idle text, after a blank line. Leftover progress is never shown once it is over.
    [Fact]
    public void An_idle_icon_says_what_the_last_check_found()
    {
        string tip = DataSyncStatusTip.Build(
            isChecking: false,
            isEnabled: true,
            isHovering: false,
            sourceLabel: "online source",
            progress: "Downloading file [10] of [10] from the online source:",
            file: "Commodore/C64/250407/250407.json",
            lastCheck: DataSyncStatusTip.LastCheck(0, Afternoon));

        Assert.Equal(
            "Data update is enabled. Click to refresh data now\n\nLast check at 14:05: all data files were up to date",
            tip);
    }

    [Theory]
    [InlineData(-1, "Last check at 14:05 failed - see the log")]
    [InlineData(0, "Last check at 14:05: all data files were up to date")]
    [InlineData(1, "Last check at 14:05: [1] file updated")]
    [InlineData(27, "Last check at 14:05: [27] files updated")]
    public void The_last_check_is_said_with_its_time_and_what_it_found(int changedCount, string expected)
    {
        Assert.Equal(expected, DataSyncStatusTip.LastCheck(changedCount, Afternoon));
    }

    // The 24-hour clock with a leading zero, whatever the machine's culture - "9:05 AM" on one
    // machine and "09.05" on another would be the same check said two ways.
    [Fact]
    public void The_time_is_a_24_hour_clock_with_a_leading_zero()
    {
        Assert.StartsWith("Last check at 09:05:", DataSyncStatusTip.LastCheck(0, new DateTime(2026, 10, 11, 9, 5, 0)));
        Assert.StartsWith("Last check at 21:30:", DataSyncStatusTip.LastCheck(0, new DateTime(2026, 10, 11, 21, 30, 0)));
    }

    // The sync's closing summaries are not progress - the outcome line says that, in words a
    // person reads. Both closings count: the manual check ends "Update complete", the launch
    // sync "Sync complete".
    [Theory]
    [InlineData("Fetching online file manifest...", true)]
    [InlineData("Downloading file [3] of [10] from the online source:", true)]
    [InlineData("Sync failed - see log", true)]
    [InlineData("Main Excel file is up to date", false)]
    [InlineData("All files UP TO DATE", false)]
    [InlineData("Update complete (2 files)", false)]
    [InlineData("Sync complete (3 new, 1 updated, 0 failed, 0 invalid)", false)]
    [InlineData("  Sync complete", false)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    [InlineData(null, false)]
    public void Only_progress_is_shown_while_a_check_runs(string? status, bool expected)
    {
        Assert.Equal(expected, DataSyncStatusTip.IsProgress(status));
    }
}
