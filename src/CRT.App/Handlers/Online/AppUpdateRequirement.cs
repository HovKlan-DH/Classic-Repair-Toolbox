using System;
using Handlers.DataHandling;

namespace Handlers.Online
{
    // ###########################################################################################
    // *** WHEN THE SERVER HAS MOVED ON, THE DRAFTS AND MAINTAINER TABS SAY SO AND STOP (owner
    // request, 2026-10-09: "When an API diff requires for the CRT app to be updated, then for both
    // the tabs, "Draft" and "Maintainer" put a fullpage modal (or alike) in there, that cannot be
    // closed, stating the app needs to be updated"). ***
    //
    // Those two tabs are the parts of CRT that use the GATED areas of the API (CRT.Server's
    // ClientVersionPolicy: /api/submissions for Drafts; /api/review, /api/admin and /api/accounts
    // for Maintainer). Everything else CRT sends - the check-in, feedback, board views, the data
    // sync - is never refused, so the rest of CRT keeps working and only these two tabs are covered
    // (UpdateRequiredOverlay, shown by Main.UpdateRequired.cs).
    //
    // This class decides WHICH of the two is covered, and with what words. Two kinds of evidence:
    //
    //   - THE API REVISION, from GET /api/health (once either tab is in use - AsksServer - and again
    //     after any refusal). A server
    //     whose revision is HIGHER than this CRT's turns it away from BOTH areas - that is the "API
    //     diff" - so both tabs are covered, in the sentence the server itself answers with
    //     (ClientVersionContract.OutdatedApi). Equal or lower is fine: the server serves its own
    //     revision and newer.
    //   - A REFUSAL - HTTP 426, ClientOutdatedAnswer - from one area's route (ApiOutdatedSignal). It
    //     covers THAT area's tab, in the server's own words: the server's other lever, a minimum
    //     CRT version, is set per area, so a Maintainer minimum must not take the Drafts tab with it.
    //     The server's words replace CRT's own - they may name the version to update to.
    //
    // *** NEVER LIFTED. *** Nothing but a newer CRT changes the server's answer - it only ever moves
    // forward - so a covered tab stays covered until CRT restarts as the updated version.
    // ###########################################################################################
    public enum AppUpdateArea
    {
        Drafts,
        Maintainer
    }

    public sealed class AppUpdateRequirement
    {
        private string? thisDraftsReason;
        private string? thisMaintainerReason;

        // ###########################################################################################
        // Whether a server reporting `serverRevision` turns a CRT built for `ownRevision` away. Null
        // is a server older than 4.6.0, which reports no revision - not a "no" from it.
        // ###########################################################################################
        public static bool IsBehind(int? serverRevision, int ownRevision) =>
            serverRevision is int server && server > ownRevision;

        // ###########################################################################################
        // Whether the health question is worth asking: only while one of the two tabs is in use - the
        // Drafts tab shown (drafts, or replies waiting) or the Maintainer tab turned on (code review,
        // 2026-10-09). Somebody using neither is asked nothing: every request is one the server
        // counts, and an answer would cover two tabs nobody sees.
        // ###########################################################################################
        public static bool AsksServer(bool draftsTabShown, bool maintainerTabEnabled) =>
            draftsTabShown || maintainerTabEnabled;

        public bool IsRequiredFor(AppUpdateArea area) => this.ReasonFor(area) is not null;

        // Both tabs covered - nothing more any answer could add.
        public bool IsRequiredEverywhere =>
            this.IsRequiredFor(AppUpdateArea.Drafts) && this.IsRequiredFor(AppUpdateArea.Maintainer);

        // The sentence the tab's overlay leads with, or null while the tab is not covered.
        public string? ReasonFor(AppUpdateArea area) =>
            area == AppUpdateArea.Drafts ? this.thisDraftsReason : this.thisMaintainerReason;

        // ###########################################################################################
        // The health answer. Behind covers both tabs, keeping any server's words already there.
        // `ownVersion` names this CRT in the sentence, as the server would. Returns whether anything
        // changed.
        // ###########################################################################################
        public bool ApplyApiRevision(int? serverRevision, int ownRevision, CrtVersion? ownVersion)
        {
            if (!AppUpdateRequirement.IsBehind(serverRevision, ownRevision))
                return false;

            string reason = ClientVersionContract.OutdatedApi(ownVersion).Message;

            bool drafts = this.SetIfNone(AppUpdateArea.Drafts, reason);
            bool maintainer = this.SetIfNone(AppUpdateArea.Maintainer, reason);

            return drafts || maintainer;
        }

        // ###########################################################################################
        // A refusal from one area's route: its tab is covered, in the server's words - CRT's own
        // sentence when it sent none. Returns whether anything changed.
        // ###########################################################################################
        public bool ApplyRefusal(AppUpdateArea area, string? serversWords, CrtVersion? ownVersion)
        {
            string? current = this.ReasonFor(area);

            if (string.IsNullOrWhiteSpace(serversWords))
                return this.SetIfNone(area, ClientVersionContract.OutdatedApi(ownVersion).Message);

            string reason = serversWords.Trim();

            if (string.Equals(current, reason, StringComparison.Ordinal))
                return false;

            this.Set(area, reason);
            return true;
        }

        private bool SetIfNone(AppUpdateArea area, string reason)
        {
            if (this.IsRequiredFor(area))
                return false;

            this.Set(area, reason);
            return true;
        }

        private void Set(AppUpdateArea area, string reason)
        {
            if (area == AppUpdateArea.Drafts)
                this.thisDraftsReason = reason;
            else
                this.thisMaintainerReason = reason;
        }
    }
}
