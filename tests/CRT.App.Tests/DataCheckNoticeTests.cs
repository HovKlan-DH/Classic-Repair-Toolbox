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
