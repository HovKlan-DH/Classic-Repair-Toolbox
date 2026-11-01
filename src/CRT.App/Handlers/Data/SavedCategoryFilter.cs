using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH CATEGORIES THE FILTER BESIDE THE SCHEMATICS STARTS WITH TICKED, once somebody has narrowed
    // it on a board (UserSettings.SelectedCategoriesByBoard).
    //
    // *** A CATEGORY THE FILTER HAS NEVER SEEN IS TICKED (code review, 2026-11-01). *** The list holds what
    // was ticked, so a category that came later - renamed on "Board: Categories", typed for one component,
    // edited in the table, or brought by a data update - read as unticked, and every component under it
    // dropped out of the list beside the schematics with nothing saying why. So what the board had when
    // the filter was chosen is kept beside it (SeenCategoriesByBoard), and only a category it had then
    // can have been unticked. As the catalogue's own keys: what is new shows.
    //
    // A filter saved before that was kept has no such list (`seen` null), and nothing can be told apart:
    // its list alone decides, once more, and Main notes what the board has from then on.
    // ###########################################################################################
    internal static class SavedCategoryFilter
    {
        public static bool IsTicked(string category, IReadOnlyCollection<string> saved, IReadOnlyCollection<string>? seen) =>
            saved.Contains(category, StringComparer.OrdinalIgnoreCase) ||
            (seen is not null && !seen.Contains(category, StringComparer.OrdinalIgnoreCase));
    }
}
