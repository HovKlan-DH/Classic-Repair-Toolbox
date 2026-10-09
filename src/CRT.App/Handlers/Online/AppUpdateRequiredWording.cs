namespace Handlers.Online
{
    // ###########################################################################################
    // WHAT A COVERED DRAFTS OR MAINTAINER TAB SAYS (owner request, 2026-10-09; see
    // AppUpdateRequirement). Shown by UpdateRequiredOverlay, which cannot be closed - so it says
    // plainly what is wrong, what it stops, that nothing is lost, and the one way out.
    //
    // The reason is the server's own sentence (or CRT's copy of it, ClientVersionContract.OutdatedApi)
    // - "This version of CRT [3.0.0] was made for an older version of the server - please update CRT
    // to the newest version." - so the overlay and every other place CRT shows a refusal agree.
    //
    // *** THE BUTTON IS THE WAY OUT. *** When CRT's own update check has already found a newer
    // version, it installs it - the update banner's Install, under the same "please wait". Otherwise
    // it opens the releases page, which lists every version, pre-releases included: a tester on a
    // BETA would find the newest stable release OLDER than what is installed.
    // ###########################################################################################
    public static class AppUpdateRequiredWording
    {
        public const string Heading = "CRT has to be updated";

        public const string RestOfCrt = "Everything else in CRT works as before.";

        public const string InstallButton = "Install update";

        public const string DownloadPageButton = "Open download page";

        // What the covered tab cannot do until CRT is updated - and, for drafts, that they are safe.
        // Installing the update asks first about a table's unsaved edits (Main.ConfirmLeavingTablesAsync),
        // which is what keeps "nothing is lost" true.
        public static string WhatItStops(AppUpdateArea area) => area == AppUpdateArea.Drafts
            ? "Until it is, drafts cannot be submitted and the server cannot be asked how your submissions are getting on. Your drafts stay on this computer, untouched - nothing is lost."
            : "Until it is, the Maintainer tab cannot be used - the server turns this version of CRT away.";

        // ###########################################################################################
        // On the Contribute tab, when "Edit board as draft" or "Add a new board" is pressed while
        // the Drafts tab is covered (code review, 2026-10-09). Both exist to work in that tab, so
        // they make nothing - a draft opened under the cover was a dead end.
        // ###########################################################################################
        public const string NoDraftWhileDraftsTabCovered =
            "Nothing was made: CRT has to be updated before drafts can be worked on - the Drafts tab says why.";

        // ###########################################################################################
        // The whole overlay for one tab. `pendingVersion` is the update CRT's own check found
        // (UpdateService.PendingVersion), or null when it found none or was not asked.
        // ###########################################################################################
        public static AppUpdateRequiredView For(AppUpdateArea area, string reason, string? pendingVersion)
        {
            bool installs = !string.IsNullOrWhiteSpace(pendingVersion);

            return new AppUpdateRequiredView(
                AppUpdateRequiredWording.Heading,
                reason,
                AppUpdateRequiredWording.WhatItStops(area),
                AppUpdateRequiredWording.RestOfCrt,
                installs ? $"Version [{pendingVersion!.Trim()}] is ready to install." : null,
                installs ? AppUpdateRequiredWording.InstallButton : AppUpdateRequiredWording.DownloadPageButton,
                installs);
        }
    }

    // ###########################################################################################
    // One covered tab's overlay. ReadyLine is null when no update has been found to install; then
    // the button opens the download page (ButtonInstalls false).
    // ###########################################################################################
    public sealed record AppUpdateRequiredView(
        string Heading,
        string Reason,
        string WhatItStops,
        string RestOfCrt,
        string? ReadyLine,
        string ButtonLabel,
        bool ButtonInstalls);
}
