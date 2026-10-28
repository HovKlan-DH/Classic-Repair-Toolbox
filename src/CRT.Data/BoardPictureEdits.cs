using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE BOARD'S PICTURES, CHANGED ON THE CONTRIBUTE TAB (CRT 4.0 step 3 - the plan: "Pictures:
    // today's Schematic images content, moved inline"; owner, 2026-10-28: "do what is best for having
    // the best UX for the user. It should be as easy as possible"). The "Schematic images" window is
    // gone: pictures are added, named, ordered and removed on their cards in the Pictures part.
    //
    // What a picture's NAME carries (highlights, KiCad calibration) follows a rename by
    // SchematicRename and goes with a removal by DraftEditRules, as it did in the window; this holds
    // only what the window kept to itself:
    //
    //   - NEW PICTURES come after the last, each under a name no picture has - its file's name
    //     without its type, "Top (2)" when "Top" is taken (FreeName). A name is the picture's key, and
    //     a second "Top" would be a duplicate the checks warn about and a rename could not tell apart.
    //   - A PICTURE MOVES one place up or down - the board's own order, the default every user of
    //     the data sees the pictures in (a user's own order on the Schematics tab stays theirs).
    //
    // PURE: the board in, the board out - the very board when nothing changes, so no step is made.
    // ###########################################################################################
    public static class BoardPictureEdits
    {
        public const string FallbackName = "Board image";

        // The board with `pictures` after its last picture - each a new row by its name and image file.
        public static BoardData WithAdded(BoardData board, IReadOnlyList<(string Name, string File)> pictures)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(pictures);

            if (pictures.Count == 0)
            {
                return board;
            }

            var rows = new List<BoardSchematicEntry>(board.Schematics);
            rows.AddRange(pictures.Select(picture => new BoardSchematicEntry
            {
                SchematicName = picture.Name.Trim(),
                SchematicImageFile = picture.File.Trim(),
            }));

            return board.WithSchematics(rows);
        }

        // ###########################################################################################
        // A picture's name no other picture has: `desired`, trimmed - else "<desired> (2)", "(3)"...
        // A blank one is "Board image": a file called ".png" still needs a name, or its row would have
        // a blank key that could never be found or removed again. Capitals are no difference.
        // ###########################################################################################
        public static string FreeName(string? desired, IEnumerable<string> taken)
        {
            var used = new HashSet<string>((taken ?? []).Select(name => name?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);

            string name = string.IsNullOrWhiteSpace(desired) ? BoardPictureEdits.FallbackName : desired.Trim();
            string candidate = name;

            for (int number = 2; used.Contains(candidate); number++)
            {
                candidate = $"{name} ({number})";
            }

            return candidate;
        }

        // ###########################################################################################
        // The names pictures added from these files take: each file's name without its type, free
        // among the board's pictures AND the ones added before it - two files called "top.png" from
        // two folders become "top" and "top (2)".
        // ###########################################################################################
        public static List<string> NamesFor(IReadOnlyList<string> files, IEnumerable<string> taken)
        {
            ArgumentNullException.ThrowIfNull(files);

            var used = new List<string>(taken ?? []);
            var names = new List<string>();

            foreach (string file in files)
            {
                string stem = Path.GetFileNameWithoutExtension((file ?? string.Empty).Replace('\\', '/').Split('/')[^1]);
                string name = BoardPictureEdits.FreeName(stem, used);

                used.Add(name);
                names.Add(name);
            }

            return names;
        }

        // ###########################################################################################
        // The board with the picture `row` moved `by` places - -1 up, +1 down - or the very board at
        // an end, or when the row is not there (the draft moved meanwhile).
        // ###########################################################################################
        public static BoardData Moved(BoardData board, SchematicRowRef row, int by)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(row);

            int from = SchematicRows.IndexOf(board.Schematics, row);
            int to = from + by;

            if (from < 0 || by == 0 || to < 0 || to >= board.Schematics.Count)
            {
                return board;
            }

            var rows = new List<BoardSchematicEntry>(board.Schematics);
            BoardSchematicEntry moving = rows[from];
            rows.RemoveAt(from);
            rows.Insert(to, moving);

            return board.WithSchematics(rows);
        }
    }
}
