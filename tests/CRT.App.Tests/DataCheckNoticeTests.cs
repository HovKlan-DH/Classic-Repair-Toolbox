using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DataCheckNotice - what the data-check card says when a check is over, and whether it stays
// (owner report, 2026-10-11: the launch check said nothing anywhere). The rule: it STAYS when
// there is something to do or somebody asked; otherwise it goes by itself.
// ###########################################################################################
public sealed class DataCheckNoticeTests
{
    // Nothing changed: said, then gone - whoever started it, nothing needs doing.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_check_that_changed_nothing_says_so_and_goes_by_itself(bool atLaunch)
    {
        var notice = DataCheckNotice.After(0, atLaunch, "stable source");

        Assert.Equal("All data files are up to date with the stable source", notice.Text);
        Assert.False(notice.OffersRefresh);
        Assert.True(notice.GoesByItself);
    }

    // ###########################################################################################
    // *** FILES THAT COULD NOT BE DOWNLOADED ARE SAID, NEVER "ALL UP TO DATE" (owner request, 2026-11-07: every error and
    // warning in the UI). *** The card said the data was up to date while downloads had been given up on. It stays,
    // with what did arrive offering the refresh; at launch too - the data on this computer is behind.
    // ###########################################################################################
    [Theory]
    [InlineData(0, 1, true, "[1] file could not be downloaded from the stable source - the data on this computer may be behind. \"Show errors and warnings\" in the \"Configuration\" tab says which", false)]
    [InlineData(0, 3, false, "[3] files could not be downloaded from the stable source - the data on this computer may be behind. \"Show errors and warnings\" in the \"Configuration\" tab says which", false)]
    [InlineData(1, 2, true, "[1] file updated in background and [2] files could not be downloaded from the stable source - please refresh board", true)]
    [InlineData(4, 1, false, "[4] files updated and [1] file could not be downloaded from the stable source - please refresh board", true)]
    public void Files_that_could_not_be_downloaded_are_said_and_the_card_stays(int changed, int failed, bool atLaunch, string expected, bool refresh)
    {
        var notice = DataCheckNotice.After(changed, atLaunch, "stable source", failed);

        Assert.Equal(expected, notice.Text);
        Assert.Equal(refresh, notice.OffersRefresh);
        Assert.False(notice.GoesByItself);
    }

    // A check that could not run at all says that, whatever failed before it.
    [Fact]
    public void A_check_that_could_not_run_says_so_whatever_failed_before()
    {
        Assert.StartsWith("Could not check", DataCheckNotice.After(-1, atLaunch: true, "stable source", failedCount: 2).Text, StringComparison.Ordinal);
    }

    // Files changed: the board on screen may be out of date, so the card offers the refresh and
    // waits for it. The launch check ran in the background, and says so - the same words as before.
    [Theory]
    [InlineData(1, true, "[1] file updated in background - please refresh board")]
    [InlineData(4, true, "[4] files updated in background - please refresh board")]
    [InlineData(1, false, "[1] file updated - please refresh board")]
    [InlineData(4, false, "[4] files updated - please refresh board")]
    public void A_check_that_changed_files_offers_the_refresh_and_stays(int changedCount, bool atLaunch, string expected)
    {
        var notice = DataCheckNotice.After(changedCount, atLaunch, "BETA source");

        Assert.Equal(expected, notice.Text);
        Assert.True(notice.OffersRefresh);
        Assert.False(notice.GoesByItself);
    }

    // A check that could not finish names the source. Somebody who asked for it is told until they
    // close it; the launch check's failure goes by itself, or it would greet everybody working
    // offline at every launch.
    [Theory]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public void A_check_that_could_not_finish_stays_only_when_somebody_asked_for_it(bool atLaunch, bool goesByItself)
    {
        var notice = DataCheckNotice.After(-1, atLaunch, "BETA source");

        Assert.Equal("Could not check the BETA source for new data - see the log", notice.Text);
        Assert.False(notice.OffersRefresh);
        Assert.Equal(goesByItself, notice.GoesByItself);
    }
}
