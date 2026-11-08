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
        //
        // *** A FILE THAT COULD NOT BE DOWNLOADED IS SAID, NEVER "ALL UP TO DATE" (owner request, 2026-11-07: every error
        // and warning in the UI). *** `failedCount` files the check gave up on - a server error, a checksum that did not
        // match, a download that stopped moving: the data on this computer is behind on those, so the card stays, saying
        // how many, beside what did arrive. Which ones, and why, is in "Errors and warnings" (ProblemsLog - each failure
        // is logged as it happens).
        // ###########################################################################################
        public static DataCheckNotice After(int changedCount, bool atLaunch, string sourceLabel, int failedCount = 0)
        {
            string updated = atLaunch ? "updated in background" : "updated";

            if (changedCount >= 0 && failedCount > 0)
            {
                string failed = failedCount == 1
                    ? $"[1] file could not be downloaded from the {sourceLabel}"
                    : $"[{failedCount}] files could not be downloaded from the {sourceLabel}";

                return changedCount switch
                {
                    0 => new($"{failed} - the data on this computer may be behind. \"Show errors and warnings\" in the \"Configuration\" tab says which", OffersRefresh: false, GoesByItself: false),
                    1 => new($"[1] file {updated} and {failed} - please refresh board", OffersRefresh: true, GoesByItself: false),
                    _ => new($"[{changedCount}] files {updated} and {failed} - please refresh board", OffersRefresh: true, GoesByItself: false),
                };
            }

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
