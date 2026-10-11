using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One entry in a schematic image's "CAD name" drop-down: the value written to the workbook's
    // "CAD name" column, and what the drop-down says.
    // ###########################################################################################
    public sealed record SchematicCadNameChoice(string CadName, string Label);

    // ###########################################################################################
    // WHICH CAD NAMES A SCHEMATIC IMAGE CAN PICK (owner request, 2026-10-10: "have that mapping
    // available, so I can select a KiCad name. Once selected, it cannot be selected again").
    //
    // A CAD name is the KiCad VIEW an image's traces are drawn from - the Schematics tab looks for
    // the view whose display name equals it, case-insensitively (ResolveKiCadViewForSchematicName).
    // So the choices are the project's view names, in the project's own order (the order "CAD names
    // you can use" lists them in).
    //
    // *** A NAME ANOTHER IMAGE HAS IS LEFT OUT ENTIRELY (owner request, 2026-10-10: "remove the CAD
    // name entirely from the drop-down when it has been selected once, so it will be easier and
    // easier to select the correct one"). *** It was first listed greyed out with "(used by ...)";
    // on a board with two dozen views that buried the few still free. It comes back as soon as that
    // image picks another name or "No CAD name".
    //
    // *** THE IMAGE'S OWN NAME IS ALWAYS LISTED. *** Two images sharing one name (typed into Excel)
    // each keep theirs - nothing here takes a name away. A name that is no view of this KiCad data is
    // still listed, as it is: the Schematics tab also resolves "Top (replica)", "Bottom (replica)"
    // and page numbers, which no view is called, so this does not call such a name wrong.
    // ###########################################################################################
    public static class SchematicCadNameChoices
    {
        public const string NoCadNameLabel = "No CAD name";

        // The choices for the image at imageIndex, given every image's schematic name and CAD name
        // in the window's order and the KiCad project's view names.
        public static IReadOnlyList<SchematicCadNameChoice> For(
            IReadOnlyList<(string SchematicName, string CadName)> images,
            int imageIndex,
            IEnumerable<string> viewNames)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(imageIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(imageIndex, images.Count);

            string own = images[imageIndex].CadName?.Trim() ?? string.Empty;

            var heldByOthers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < images.Count; i++)
            {
                string cadName = images[i].CadName?.Trim() ?? string.Empty;

                if (i != imageIndex && cadName.Length > 0)
                {
                    heldByOthers.Add(cadName);
                }
            }

            var choices = new List<SchematicCadNameChoice> { new(string.Empty, SchematicCadNameChoices.NoCadNameLabel) };
            var listed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string viewName in viewNames.Select(name => name?.Trim() ?? string.Empty))
            {
                if (viewName.Length == 0 || !listed.Add(viewName))
                {
                    continue;
                }

                bool isOwn = string.Equals(viewName, own, StringComparison.OrdinalIgnoreCase);

                if (isOwn || !heldByOthers.Contains(viewName))
                {
                    choices.Add(new SchematicCadNameChoice(viewName, viewName));
                }
            }

            if (own.Length > 0 && !listed.Contains(own))
            {
                choices.Add(new SchematicCadNameChoice(own, own));
            }

            return choices;
        }

        // The choice showing this CAD name - the first that matches case-insensitively, or "No CAD
        // name" for a blank one. For takes care that the image's own name is always among them.
        public static SchematicCadNameChoice Selected(IReadOnlyList<SchematicCadNameChoice> choices, string? cadName)
        {
            string wanted = cadName?.Trim() ?? string.Empty;

            return choices.FirstOrDefault(choice => string.Equals(choice.CadName, wanted, StringComparison.OrdinalIgnoreCase))
                ?? choices[0];
        }
    }
}
