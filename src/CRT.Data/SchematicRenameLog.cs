using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One schematic image renamed in published data: from which name, to which, the image FILE it
    // was renamed on, and the revision (publish) date it happened in - "2026-October-10", as the
    // workbook's revision date reads.
    // ###########################################################################################
    public sealed record SchematicRenameRecord(string From, string To, string File, string Date);

    // One schematic image of a board: its name and its image file (data-root relative).
    public sealed record SchematicImage(string Name, string File);

    // ###########################################################################################
    // THE BOARD'S LIST OF RENAMED SCHEMATIC IMAGES, AND HOW CRT FOLLOWS A RENAME (owner decisions,
    // 2026-10-10: "If the schematic is renamed and this renaming is sent as a contribution, and
    // accepted, then it must also change for everyone updating to this data. It should probably
    // show that in the board JSON file, as I do not want this data in Excel" - and, with the image
    // file in it, "to be unambiguous" when a name is later given to another image).
    //
    // Every CRT keeps things of its user's own under a board's image NAMES - worklog areas, image
    // orders, a hidden image, the image last looked at (LocalSchematicRename). When an image is
    // renamed, those must move to the new name.
    //
    // *** CRT FOLLOWS THE IMAGE, NOT THE NAME. *** It remembers, per board, which image file each
    // name was on the last time it showed the board (UserSettings.SchematicImagesByBoard). The next
    // time, a name is followed to wherever ITS image file is now (Resolver) - which is right in every
    // direction: a rename published, the user's own draft rename, the draft discarded or rejected,
    // "view boards as officially published", a rename pushed back out of BETA, two images swapping
    // names, and a name later given to ANOTHER image (the old name's things go with the image they
    // were made on, not with the name).
    //
    // *** THE PUBLISHED LIST IS FOR WHEN THERE IS NOTHING REMEMBERED. *** The board's .json carries
    // every rename ("Schematic renames", beside the highlights and the KiCad calibrations -
    // BoardSidecarWriter writes it, BoardComponentHighlightStorage reads it), oldest first, as
    // {From, To, File, Date}. The SERVER writes it at each publish (ForPublish), with the rule it
    // already follows highlights with (SubmissionRowsBoard.FollowSchematics: a removed name and an
    // added one on the same image file are one image renamed) - nothing extra is sent with a
    // submission, and the push to stable copies the file. A CRT seeing a board for the first time -
    // freshly updated, with worklogs older than its memory - follows the list instead. Then a name
    // the board HAS stays where it is: whether an old worklog on it meant the image renamed away or
    // the one that has the name now cannot be known, and nothing is moved on a guess.
    // ###########################################################################################
    public static class SchematicRenameLog
    {
        public const string Root = "Schematic renames";

        // ###########################################################################################
        // The list a publish writes: what the board's .json already lists, then every rename this
        // publish makes (`published` -> `board`), with the image file and dated `date`. A new board
        // (`published` null) has nothing to rename.
        // ###########################################################################################
        public static List<SchematicRenameRecord> ForPublish(
            IReadOnlyList<SchematicRenameRecord> existing,
            BoardData? published,
            BoardData board,
            string date)
        {
            var result = new List<SchematicRenameRecord>(existing);

            if (published is null)
            {
                return result;
            }

            IReadOnlyDictionary<string, string?> follow = SubmissionRowsBoard.FollowSchematics(published.Schematics, board.Schematics);

            foreach ((string oldName, string newName) in SchematicRename.RenamesIn(follow))
            {
                string file = board.Schematics
                    .FirstOrDefault(entry => SchematicRename.SameName(entry.SchematicName, newName))?
                    .SchematicImageFile?.Trim() ?? string.Empty;

                result.Add(new SchematicRenameRecord(oldName.Trim(), newName.Trim(), file, date?.Trim() ?? string.Empty));
            }

            return result;
        }

        // A board's images as name -> file, the shape the resolver and the user's memory use.
        public static List<SchematicImage> ImagesOf(BoardData board) =>
            board.Schematics
                .Where(entry => !string.IsNullOrWhiteSpace(entry.SchematicName))
                .Select(entry => new SchematicImage(entry.SchematicName.Trim(), entry.SchematicImageFile?.Trim() ?? string.Empty))
                .ToList();

        // ###########################################################################################
        // Where each name goes on the board as it is now (`board`) - the board's own spelling of the
        // new name - or null when it stays where it is.
        //
        //   1. A name CRT remembers (`remembered`, the board as it was last shown) goes where its
        //      image FILE is now - when exactly one image has that file. On several, it stays: which
        //      one it meant cannot be told.
        //   2. Otherwise a name the board has stays.
        //   3. Otherwise the board's published list (`published`) says where it went: the image file
        //      its latest rename names, where that is now; failing that, the renames followed name by
        //      name to the first name the board has.
        // ###########################################################################################
        public static Func<string, string?> Resolver(
            IReadOnlyList<SchematicImage> board,
            IReadOnlyList<SchematicImage>? remembered,
            IReadOnlyList<SchematicRenameRecord> published)
        {
            // An image without a name is nowhere a name can go.
            List<SchematicImage> named = board.Where(image => SchematicRenameLog.Clean(image.Name).Length > 0).ToList();

            return name =>
            {
                string wanted = name?.Trim() ?? string.Empty;

                if (wanted.Length == 0)
                {
                    return null;
                }

                SchematicImage? was = remembered?.FirstOrDefault(image => SchematicRename.SameName(image.Name, wanted));

                if (was is not null && !string.IsNullOrWhiteSpace(was.File))
                {
                    List<SchematicImage> now = named.Where(image => SchematicRenameLog.SameFile(image.File, was.File)).ToList();

                    if (now.Count == 1)
                    {
                        return SchematicRename.SameName(now[0].Name, wanted) ? null : SchematicRenameLog.Clean(now[0].Name);
                    }

                    if (now.Count > 1)
                    {
                        return null;
                    }

                    // The image left the board - the name rules below decide.
                }

                if (named.Any(image => SchematicRename.SameName(image.Name, wanted)))
                {
                    return null;
                }

                return SchematicRenameLog.ByPublishedList(wanted, named, published);
            };
        }

        // ###########################################################################################
        // Where the published list says a name the board does not have went: the image file of its
        // LATEST rename, where that file is now (one image only) - or, without a file to go by, the
        // renames followed name by name, nearest first, to the first name the board has.
        // ###########################################################################################
        private static string? ByPublishedList(string name, IReadOnlyList<SchematicImage> board, IReadOnlyList<SchematicRenameRecord> published)
        {
            SchematicRenameRecord? latest = published.LastOrDefault(rename => SchematicRename.SameName(rename.From, name));

            if (latest is not null && !string.IsNullOrWhiteSpace(latest.File))
            {
                List<SchematicImage> now = board.Where(image => SchematicRenameLog.SameFile(image.File, latest.File)).ToList();

                if (now.Count == 1)
                {
                    return SchematicRenameLog.Clean(now[0].Name);
                }
            }

            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { name };
            var queue = new Queue<string>();
            queue.Enqueue(name);

            while (queue.Count > 0)
            {
                string current = queue.Dequeue();

                foreach (SchematicRenameRecord rename in published.Where(rename => SchematicRename.SameName(rename.From, current)))
                {
                    string next = SchematicRenameLog.Clean(rename.To);

                    if (next.Length == 0 || !visited.Add(next))
                    {
                        continue;
                    }

                    SchematicImage? onBoard = board.FirstOrDefault(image => SchematicRename.SameName(image.Name, next));

                    if (onBoard is not null)
                    {
                        return SchematicRenameLog.Clean(onBoard.Name);
                    }

                    queue.Enqueue(next);
                }
            }

            return null;
        }

        // Two image paths name one file: trimmed, either slash, without regard to capitals.
        public static bool SameFile(string? first, string? second)
        {
            string a = (first ?? string.Empty).Trim().Replace('\\', '/');
            string b = (second ?? string.Empty).Trim().Replace('\\', '/');

            return a.Length > 0 && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        // Whether two lists name the same images under the same names, in any order.
        public static bool SameImages(IReadOnlyList<SchematicImage> first, IReadOnlyList<SchematicImage> second) =>
            first.Count == second.Count &&
            first.All(image => second.Any(other =>
                string.Equals(SchematicRenameLog.Clean(other.Name), SchematicRenameLog.Clean(image.Name), StringComparison.Ordinal) &&
                string.Equals(SchematicRenameLog.Clean(other.File), SchematicRenameLog.Clean(image.File), StringComparison.OrdinalIgnoreCase)));

        // ###########################################################################################
        // A name or file, trimmed - blank for null (code review, 2026-10-11). The images CRT remembers
        // come from the user's settings file, where a hand edit or a half-written file can leave
        // {"name": null}: that threw in the middle of the board load, which then stopped before its
        // thumbnails and worklogs, every time the board was shown.
        // ###########################################################################################
        private static string Clean(string? text) => text?.Trim() ?? string.Empty;
    }
}
