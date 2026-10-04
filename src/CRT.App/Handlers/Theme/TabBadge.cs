using System.Globalization;

namespace Handlers.Theming
{
    // ###########################################################################################
    // THE NUMBER ON A BADGE - the Maintainer tab's and the Drafts tab's in CRT's row of tabs (owner
    // requests, 2026-09-30), and the Maintainer tab's own buttons. One rule, so a badge reads the
    // same wherever it is: nothing waiting is no badge at all, and past BadgeCeiling it says "99+"
    // - a number that large is a backlog to deal with, and more digits would push a tab's title
    // along the row.
    // ###########################################################################################
    public static class TabBadge
    {
        public const int Ceiling = 99;

        // The badge's text, or null to hide it.
        public static string? Text(int count) =>
            count <= 0
                ? null
                : count > TabBadge.Ceiling
                    ? $"{TabBadge.Ceiling.ToString(CultureInfo.InvariantCulture)}+"
                    : count.ToString(CultureInfo.InvariantCulture);
    }
}
