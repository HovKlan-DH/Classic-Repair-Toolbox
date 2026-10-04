using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The two data trees as the review API names them: "beta" and "production" (the stable source).
    // The Unused files list has named them so since 2026-09-25; a system's Board data and Files take
    // the same names since 2026-10-04 (owner request: a BETA / Stable switch on both - "Shouldn't
    // there be somewhere a possibility to see what we actually do have in BETA or stable").
    // ###########################################################################################
    public static class DataTreeNames
    {
        public const string Beta = "beta";

        public const string Production = "production";

        // Whether `tree` names the stable source. Anything else - none, blank, unknown - is BETA,
        // which is what every request meant before a tree could be named.
        public static bool IsProduction(string? tree) =>
            string.Equals(tree?.Trim(), DataTreeNames.Production, StringComparison.OrdinalIgnoreCase);
    }
}
