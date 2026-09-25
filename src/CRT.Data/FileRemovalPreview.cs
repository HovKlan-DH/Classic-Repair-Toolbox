using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE FILES A PUBLISH WILL REMOVE, shown to the maintainer BEFORE they approve (owner,
    // 2026-09-25: "Is this list of file deletions visible BEFORE the maintainer/admin approves it to
    // either BETA or real? It must be, so it is clear what will happen").
    //
    // One record on both ends of the wire - the server builds it, the maintainer application shows it
    // and sends the list it showed back with the approval - so what is on screen and what is
    // removed cannot drift apart. The server refuses an approval whose list no longer matches
    // (Matches): another publish can change whether a shared file is still used, and the maintainer
    // then looks again rather than having files removed they were never shown.
    //
    // A file is removed only when the board stops citing it AND nothing else in that tree uses it
    // (DataTreeUsage). BlockedBecause is set when the tree cannot be read completely - fail closed,
    // so nothing is removed and the maintainer is told why.
    // ###########################################################################################
    public sealed record FileRemovalPreview(IReadOnlyList<string> Files, string? BlockedBecause)
    {
        public static FileRemovalPreview Nothing { get; } = new([], null);

        public bool IsBlocked => !string.IsNullOrWhiteSpace(this.BlockedBecause);

        // ###########################################################################################
        // The preview for one publish. `candidates` are the files the board stops citing
        // (DataTreeUsage.NoLongerCited); `citationsAfter` is what the written workbook(s) will cite.
        // With no candidates nothing can be removed and the tree is not read at all - the common
        // case, a rows-only edit, costs nothing. `cache` is WorkbookReadCache, for a preview that
        // is shown on every click.
        // ###########################################################################################
        public static FileRemovalPreview Compute(
            string dataRoot,
            IReadOnlyCollection<string> candidates,
            IReadOnlyDictionary<string, IReadOnlyCollection<string>>? citationsAfter,
            WorkbookReadCache? cache = null)
        {
            ArgumentNullException.ThrowIfNull(candidates);

            if (candidates.Count == 0)
                return FileRemovalPreview.Nothing;

            return FileRemovalPreview.From(DataTreeUsage.Compute(dataRoot, citationsAfter, cache), candidates);
        }

        public static FileRemovalPreview From(DataTreeUsageResult usage, IEnumerable<string> candidates)
        {
            ArgumentNullException.ThrowIfNull(usage);

            if (!usage.IsComplete)
            {
                return new FileRemovalPreview(
                    [],
                    "Nothing is removed, because the data could not be read completely: " + string.Join(" ", usage.Problems));
            }

            return new FileRemovalPreview(usage.RemovableFrom(candidates), null);
        }

        // ###########################################################################################
        // Is this the list the maintainer was shown? Order-free, exact spelling. A client that sends
        // nothing was shown nothing, which matches only an empty list.
        // ###########################################################################################
        public bool Matches(IEnumerable<string>? shown)
        {
            var mine = new HashSet<string>(this.Files, StringComparer.Ordinal);
            var theirs = new HashSet<string>((shown ?? []).Where(path => !string.IsNullOrWhiteSpace(path)), StringComparer.Ordinal);

            return mine.SetEquals(theirs);
        }
    }
}
