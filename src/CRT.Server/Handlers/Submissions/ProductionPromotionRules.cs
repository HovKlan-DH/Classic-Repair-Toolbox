namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The two small decisions the production promotion makes about ROWS rather than files
    // (2026-09-25). Pure, so both are unit tests.
    // ###########################################################################################
    public static class ProductionPromotionRules
    {
        // ###########################################################################################
        // Is this system's BETA state newer than what Production has?
        //
        // Only a system with a BETA content hash can be - that hash is written by every publish,
        // so a shipped board nobody has published through the pipeline is waiting for nothing.
        // Compared on the HASH rather than the revision date, because two publishes on one day
        // share a revision date and the second would otherwise never reach Production.
        // ###########################################################################################
        public static bool IsAwaitingProduction(SystemRecord? system)
        {
            if (system is null || string.IsNullOrWhiteSpace(system.ContentHash))
                return false;

            return !string.Equals(system.ContentHash, system.ProductionContentHash, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The state a CONTRIBUTOR is told about their submission.
        //
        // *** "merged" NOW MEANS "IN BETA". *** A merged submission is in the BETA data and reaches
        // everyone only when its system is promoted. So a merged submission decided at or before
        // the system's last promotion went out with it, and the contributor is told "published" -
        // which CRT shows as "Published to source". Before that they are told "merged", which CRT
        // shows as "Published to BETA source".
        //
        // The database state is NOT changed: "merged" stays the record of what the maintainer did,
        // and this is only how it reads from the outside.
        // ###########################################################################################
        public static string ContributorFacingState(
            string state,
            DateTimeOffset? decidedUtc,
            DateTimeOffset? systemProductionPublishedUtc)
        {
            if (!string.Equals(state, SubmissionState.Merged, StringComparison.Ordinal))
                return state;

            if (decidedUtc is null || systemProductionPublishedUtc is null)
                return state;

            return systemProductionPublishedUtc.Value >= decidedUtc.Value
                ? ProductionPromotionRules.PublishedState
                : state;
        }

        // Not a database state - see ContributorFacingState. CRT.Data's SubmissionReceipt and
        // DraftRetirement both already treat it as "live".
        public const string PublishedState = "published";
    }
}
