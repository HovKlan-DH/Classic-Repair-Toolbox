using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // A component with no rectangle on the picture, and the category its rectangle is filed under.
    internal sealed record UnmarkedComponent(string Label, string Category);

    // ###########################################################################################
    // THE COMPONENTS NOT MARKED ON A PICTURE YET (CRT 4.0 step 3, slice 3c - the plan: "Photos and
    // scanned schematics use the label editor as today, with the unmarked components listed beside
    // it"): the board's components with no rectangle among those the label editor holds for the
    // picture - once per label whatever its regions, under the board's own spelling, with the first
    // row's category, in the board's order. Labels compare trimmed, capitals ignored.
    // ###########################################################################################
    internal static class UnmarkedComponents
    {
        public static List<UnmarkedComponent> Of(IEnumerable<ComponentEntry> components, IEnumerable<string?> marked)
        {
            ArgumentNullException.ThrowIfNull(components);
            ArgumentNullException.ThrowIfNull(marked);

            var done = new HashSet<string>(marked.Select(label => label?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);
            var unmarked = new List<UnmarkedComponent>();

            foreach (ComponentEntry row in components)
            {
                string label = row.BoardLabel?.Trim() ?? string.Empty;

                if (label.Length > 0 && done.Add(label))
                {
                    unmarked.Add(new UnmarkedComponent(label, row.Category?.Trim() ?? string.Empty));
                }
            }

            return unmarked;
        }
    }
}
