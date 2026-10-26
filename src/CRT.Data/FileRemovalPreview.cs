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
    // One record on both ends of the wire - the server builds it, the Maintainer tab shows it
    // and sends the list it showed back with the approval - so what is on screen and what is
    // removed cannot drift apart. The server refuses an approval whose list no longer matches
    // (Matches): another publish can change whether a shared file is still used, and the maintainer
    // then looks again rather than having files removed they were never shown.
    //
    // A file is removed only when the board stops citing it AND nothing else in that tree uses it
    // (DataTreeUsage). BlockedBecause is set when the tree cannot be read completely - fail closed,
    // so nothing is removed and the maintainer is told why.
    // ###########################################################################################
    //
    // *** IN THE CRT 4.0 TREE A REMOVED FILE IS A STORED FILE (2026-10-25). *** "Files/3f/3f7a...png"
    // says nothing to a person - `Names` gives each the name the board knew it by
    // ("Component images/7406.png"), for showing. Files stays what is removed and what is matched.
    // Optional: an older server sends none, and an older CRT ignores it.
    public sealed record FileRemovalPreview(IReadOnlyList<string> Files, string? BlockedBecause, IReadOnlyDictionary<string, string>? Names = null)
    {
        public static FileRemovalPreview Nothing { get; } = new([], null);

        // The name to show for one of Files: its board name when Names has one, else the path itself.
        public string NameOf(string path) =>
            this.Names is not null && this.Names.TryGetValue(path, out string? name) && !string.IsNullOrWhiteSpace(name) ? name : path;

        // Every file by the name to show - what an answer carrying only names says (the Boards
        // screen's edit check, BoardEditCheckAnswer). A method, so it is no field on the wire.
        public IReadOnlyList<string> FileNames() => this.Files.Select(this.NameOf).ToList();

        // Is this the list shown BY NAME? The same rule as Matches, on FileNames.
        public bool MatchesNames(IEnumerable<string>? shown) =>
            new FileRemovalPreview(this.FileNames(), this.BlockedBecause).Matches(shown);

        // The same preview with each of its files' board names, from `names` (tree path -> name).
        public FileRemovalPreview Named(IReadOnlyDictionary<string, string>? names)
        {
            if (names is null || this.Files.Count == 0)
                return this;

            var named = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (string path in this.Files)
            {
                if (names.TryGetValue(path, out string? name) && !string.IsNullOrWhiteSpace(name))
                    named[path] = name;
            }

            return named.Count == 0 ? this : this with { Names = named };
        }

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
