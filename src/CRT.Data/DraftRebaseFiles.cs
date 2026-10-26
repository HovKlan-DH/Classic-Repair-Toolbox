using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE FILES' HALF OF BRINGING A CRT 4.0 DRAFT UP TO DATE (2026-10-24) - beside DraftRebase's
    // rows. A 4.0 draft names its files in its board file, each with the checksum of its bytes, so
    // the files merge as the rows do, by name, three ways (BASE as copied, DRAFT now, PUBLISHED now):
    //
    //   - a file the contributor left as copied follows the update - replaced, added or removed;
    //   - a file the contributor replaced (other bytes than the base's) stays theirs;
    //   - a file the update added comes in; one the contributor removed stays removed.
    //
    // *** A REPLACED FILE IS A CHANGE TO THE ROWS THAT CITE IT. *** The rows alone do not show it -
    // the contributor's U8 picture row still reads "Component images/U8.png" - so an update that
    // deleted the row would delete it as untouched, and the contributor's picture with it, the
    // list following the rows. MarkReplaced makes the base's citing rows differ, so the merge keeps
    // the contributor's rows (and lists a conflict where the update changed them too).
    //
    // Pure: nothing here reads or writes a file (DraftRebaser does).
    // ###########################################################################################
    public static class DraftRebaseFiles
    {
        // Appended to a file name in the base's rows to make them differ - never written anywhere.
        private const string ReplacedMark = "\u001Freplaced";

        // ###########################################################################################
        // The names whose bytes the contributor replaced: listed in the base and in the draft with
        // other checksums (ignoring capitals), or lying in the draft folder as a new file under a
        // name the base lists (`arrivedInDraft`).
        // ###########################################################################################
        public static HashSet<string> Replaced(
            IReadOnlyList<BoardDocument.StoredFileRow> based,
            IReadOnlyList<BoardDocument.StoredFileRow> draft,
            Func<string, bool> arrivedInDraft)
        {
            Dictionary<string, BoardDocument.StoredFileRow> draftByName = DraftRebaseFiles.ByName(draft);
            var replaced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in based)
            {
                if (arrivedInDraft(file.Name) ||
                    (draftByName.TryGetValue(file.Name, out BoardDocument.StoredFileRow? mine) && !DraftRebaseFiles.SameBytes(mine, file)))
                {
                    replaced.Add(file.Name);
                }
            }

            return replaced;
        }

        // The base board with every row citing a `replaced` file told apart from the draft's - see
        // the header. The lists are new; the rows not citing one are shared.
        public static BoardData MarkReplaced(BoardData based, IReadOnlySet<string> replaced)
        {
            ArgumentNullException.ThrowIfNull(based);

            if (replaced.Count == 0)
                return based;

            bool Cites(string? file) => !string.IsNullOrWhiteSpace(file) && replaced.Contains(BoardFileNames.Normalise(file));

            return new BoardData
            {
                RevisionDate = based.RevisionDate,
                HardwareName = based.HardwareName,
                BoardName = based.BoardName,
                Schematics = based.Schematics
                    .Select(row => Cites(row.SchematicImageFile) ? DraftRebaseFiles.WithMark(row, nameof(BoardSchematicEntry.SchematicImageFile), row.SchematicImageFile) : row)
                    .ToList(),
                Components = based.Components,
                ComponentImages = based.ComponentImages
                    .Select(row => Cites(row.File) ? DraftRebaseFiles.WithMark(row, nameof(ComponentImageEntry.File), row.File) : row)
                    .ToList(),
                ComponentHighlights = based.ComponentHighlights,
                ComponentLocalFiles = based.ComponentLocalFiles
                    .Select(row => Cites(row.File) ? DraftRebaseFiles.WithMark(row, nameof(ComponentLocalFileEntry.File), row.File) : row)
                    .ToList(),
                ComponentLinks = based.ComponentLinks,
                BoardLocalFiles = based.BoardLocalFiles
                    .Select(row => Cites(row.File) ? DraftRebaseFiles.WithMark(row, nameof(BoardLocalFileEntry.File), row.File) : row)
                    .ToList(),
                BoardLinks = based.BoardLinks,
                Credits = based.Credits,
                KiCadImportantSignals = based.KiCadImportantSignals,

                // Cite no file.
                ComponentPins = based.ComponentPins,
                ComponentReplacements = based.ComponentReplacements,
            };
        }

        // ###########################################################################################
        // The draft's file list after the merge - see the header - and how many entries followed
        // the update. The draft's write then keeps the entries its merged rows cite (DraftFiles).
        // A name in `replaced` is the contributor's whatever the update did.
        // ###########################################################################################
        public static List<BoardDocument.StoredFileRow> Merge(
            IReadOnlyList<BoardDocument.StoredFileRow> based,
            IReadOnlyList<BoardDocument.StoredFileRow> draft,
            IReadOnlyList<BoardDocument.StoredFileRow> published,
            IReadOnlySet<string> replaced,
            out int followed)
        {
            Dictionary<string, BoardDocument.StoredFileRow> baseByName = DraftRebaseFiles.ByName(based);
            Dictionary<string, BoardDocument.StoredFileRow> publishedByName = DraftRebaseFiles.ByName(published);
            Dictionary<string, BoardDocument.StoredFileRow> merged = DraftRebaseFiles.ByName(draft);

            followed = 0;

            foreach (BoardDocument.StoredFileRow theirs in published)
            {
                if (replaced.Contains(theirs.Name))
                    continue;

                if (merged.TryGetValue(theirs.Name, out BoardDocument.StoredFileRow? mine))
                {
                    // Added on both sides: the contributor's, as for a row both changed.
                    if (!baseByName.ContainsKey(theirs.Name))
                        continue;

                    // Left as copied, and replaced by the update: the update's.
                    if (!DraftRebaseFiles.SameBytes(mine, theirs))
                    {
                        merged[theirs.Name] = theirs;
                        followed++;
                    }
                }
                else if (!baseByName.ContainsKey(theirs.Name))
                {
                    // Added by the update. (In the base and not in the draft: the contributor removed it.)
                    merged[theirs.Name] = theirs;
                    followed++;
                }
            }

            foreach (BoardDocument.StoredFileRow mine in draft)
            {
                // Left as copied, and removed by the update.
                if (!replaced.Contains(mine.Name) &&
                    baseByName.ContainsKey(mine.Name) &&
                    !publishedByName.ContainsKey(mine.Name))
                {
                    merged.Remove(mine.Name);
                    followed++;
                }
            }

            return merged.Values
                .OrderBy(file => file.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(file => file.Name, StringComparer.Ordinal)
                .ToList();
        }

        private static Dictionary<string, BoardDocument.StoredFileRow> ByName(IReadOnlyList<BoardDocument.StoredFileRow> files)
        {
            var byName = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in files ?? [])
                byName.TryAdd(file.Name, file);

            return byName;
        }

        private static bool SameBytes(BoardDocument.StoredFileRow a, BoardDocument.StoredFileRow b) =>
            string.Equals(a.Sha256, b.Sha256, StringComparison.OrdinalIgnoreCase);

        // A copy of the row with its file field marked. The row types are plain init-only classes,
        // so the copy is a member-wise one with that one property set again.
        private static readonly MethodInfo CloneMethod =
            typeof(object).GetMethod("MemberwiseClone", BindingFlags.Instance | BindingFlags.NonPublic)!;

        private static T WithMark<T>(T row, string property, string file) where T : class
        {
            var copy = (T)DraftRebaseFiles.CloneMethod.Invoke(row, null)!;
            typeof(T).GetProperty(property)!.SetValue(copy, file + DraftRebaseFiles.ReplacedMark);
            return copy;
        }
    }
}
