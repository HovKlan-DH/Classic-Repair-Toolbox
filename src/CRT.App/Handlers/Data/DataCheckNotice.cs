namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What the data-check notice says once a check is over, and whether it stays (owner report,
    // 2026-10-11: "I am checking for new data updates at app launch, but I do not see any
    // information about that anywhere").
    //
    // The check has a card of its own over the tabs' bottom-right corner (DataCheckBanner in
    // Main.axaml): while it runs, what it is doing (DataSyncStatusTip.Checking); then one of these.
    // The card floats, so coming and going moves nothing - which is why it may come and go at all.
    //
    // A card STAYS when there is something to do or somebody asked: files changed (refresh the
    // board), or a check somebody started could not finish. Otherwise it GOES BY ITSELF after a
    // few seconds - "all up to date" needs nothing done, and neither does a launch check that could
    // not run, which would otherwise greet everybody working offline at every launch.
    //
    // Pure, so every wording and every stay-or-go is tested without a window.
    // ###########################################################################################
    internal sealed record DataCheckNotice(string Text, bool OffersRefresh, bool GoesByItself)
    {
        // ###########################################################################################
        // The outcome of a check that changed `changedCount` files - negative when it could not
        // finish (no data folder, or no manifest). `atLaunch` is the check CRT runs by itself at
        // launch; otherwise somebody started it, from the sync icon or the Configuration tab.
        // ###########################################################################################
        public static DataCheckNotice After(int changedCount, bool atLaunch, string sourceLabel)
        {
            string updated = atLaunch ? "updated in background" : "updated";

            return changedCount switch
            {
                < 0 => new($"Could not check the {sourceLabel} for new data - see the log", OffersRefresh: false, GoesByItself: atLaunch),
                0 => new($"All data files are up to date with the {sourceLabel}", OffersRefresh: false, GoesByItself: true),
                1 => new($"[1] file {updated} - please refresh board", OffersRefresh: true, GoesByItself: false),
                _ => new($"[{changedCount}] files {updated} - please refresh board", OffersRefresh: true, GoesByItself: false),
            };
        }
    }
}
