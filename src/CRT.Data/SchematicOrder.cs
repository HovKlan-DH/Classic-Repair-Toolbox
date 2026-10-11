using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // PUTS A BOARD'S SCHEMATICS IN THE ORDER A CONTRIBUTOR ARRANGED THEM (owner request,
    // 2026-09-27) - by dragging the rows in the "Schematic images" window.
    //
    // *** THE ORDER IS GIVEN AS NAMES, THE WHOLE LIST AS THE WINDOW SHOWS IT, never as "move row
    // 3 to 7". *** The workbook is re-read when it is saved, and it can have changed since the
    // window read it (Excel is outside the application's control). An index would then move a
    // different row than the one dragged; names say what the contributor saw.
    //
    // Nothing is ever lost or invented: a schematic the list does not name keeps its place after
    // the named ones, in the order it already had, and a name that matches no schematic is
    // skipped. Names match case-insensitively, like every schematic-name comparison in the app,
    // and a name listed twice takes the next schematic of that name - so two rows sharing one in a
    // hand-edited workbook both survive.
    //
    // *** EACH NAME COMES WITH ITS IMAGE FILE (code review, 2026-10-10). *** By name alone, two
    // rows sharing a name could not change places: whichever was dragged, the first of the name
    // took the first place. A file given must match too (SchematicRows.SameImageFile); a null file
    // matches by the name alone.
    // ###########################################################################################
    public static class SchematicOrder
    {
        public static List<BoardSchematicEntry> Apply(
            IReadOnlyList<BoardSchematicEntry> schematics,
            IReadOnlyList<(string SchematicName, string? ImageFile)> order)
        {
            ArgumentNullException.ThrowIfNull(schematics);

            var remaining = new List<BoardSchematicEntry>(schematics);
            var ordered = new List<BoardSchematicEntry>(schematics.Count);

            foreach ((string name, string? imageFile) in order ?? [])
            {
                int index = remaining.FindIndex(entry =>
                    SchematicRename.SameName(entry.SchematicName, name) &&
                    (imageFile is null || SchematicRows.SameImageFile(entry.SchematicImageFile, imageFile)));

                if (index < 0)
                {
                    continue;
                }

                ordered.Add(remaining[index]);
                remaining.RemoveAt(index);
            }

            ordered.AddRange(remaining);

            return ordered;
        }
    }
}
