using CRT.Server.Configuration;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Usage
{
    // A board's names as CRT's drop-down lists show them: "Commodore 64", "250407".
    public sealed record BoardViewNames(string HardwareName, string BoardName);

    // ###########################################################################################
    // WHICH BOARDS A VIEW MAY NAME, AND WHAT THEY ARE CALLED: the rows of the PUBLISHED main Excel
    // data file - production's first, then BETA's (a new board placed there is viewed by those
    // downloading from BETA before production has it).
    //
    // *** A VIEW OF A BOARD NEITHER LISTS IS NOT COUNTED. *** That is what keeps made-up text out
    // of crt_board_views and off the public Fun facts page: the names stored are the listing's own,
    // never the sender's, and an id no listing has names nothing. A contributor's draft-only board
    // is never sent by CRT in the first place.
    //
    // Reading a workbook is costly, so each listing is read once per file VERSION (its time and
    // size), and the file is looked at again at most every RecheckAfter - a view sent a minute
    // after a new master arrives may still be judged by the old one, which costs nothing.
    // ###########################################################################################
    public sealed class BoardViewNameDirectory
    {
        public static readonly TimeSpan RecheckAfter = TimeSpan.FromMinutes(1);

        public delegate bool ListingReader(string masterPath, out IReadOnlyList<MasterListingRow> rows);

        private readonly IReadOnlyList<string> thisRoots;
        private readonly ListingReader thisReader;
        private readonly Func<DateTimeOffset> thisClock;
        private readonly object thisGate = new();

        private readonly Dictionary<string, Listing> thisListings = new(StringComparer.Ordinal);

        public BoardViewNameDirectory(ServerOptions options)
            : this(BoardViewNameDirectory.RootsOf(options), BoardViewNameDirectory.ReadListing, () => DateTimeOffset.UtcNow)
        {
        }

        // The test seam: the roots, how a listing is read, and the clock.
        internal BoardViewNameDirectory(IReadOnlyList<string> roots, ListingReader reader, Func<DateTimeOffset> clock)
        {
            this.thisRoots = roots;
            this.thisReader = reader;
            this.thisClock = clock;
        }

        // Production's data root first (the promotion's own, else the older setting), then BETA's.
        public static IReadOnlyList<string> RootsOf(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);

            string? production = !string.IsNullOrWhiteSpace(options.ProductionDataTreeRoot)
                ? options.ProductionDataTreeRoot
                : options.ProductionTreeRoot;

            return new[] { production, options.DataTreeRoot }
                .Where(root => !string.IsNullOrWhiteSpace(root))
                .Select(root => root!)
                .Distinct(StringComparer.Ordinal)
                .ToList();
        }

        // The names for a board id, or null when no published listing has it.
        public BoardViewNames? Find(string? boardId)
        {
            string id = boardId?.Trim() ?? string.Empty;

            if (id.Length == 0)
                return null;

            lock (this.thisGate)
            {
                foreach (string root in this.thisRoots)
                {
                    if (this.ListingFor(root).Names.TryGetValue(id, out BoardViewNames? names))
                        return names;
                }
            }

            return null;
        }

        private Listing ListingFor(string root)
        {
            DateTimeOffset now = this.thisClock();

            if (this.thisListings.TryGetValue(root, out Listing? known) && now - known.CheckedUtc < BoardViewNameDirectory.RecheckAfter)
                return known;

            string? master = MasterListing.NewestMasterPath(root);
            (DateTime, long)? version = BoardViewNameDirectory.VersionOf(master);

            // The same file as before, read whole: only the time it was looked at moves.
            if (known is not null && known.IsComplete && known.MasterPath == master && known.Version == version)
            {
                known.CheckedUtc = now;
                return known;
            }

            IReadOnlyList<MasterListingRow> rows = [];
            bool read = master is not null && this.thisReader(master, out rows);

            // A file that cannot be read just now (being replaced, say) keeps the names last read
            // from it, and is tried again after RecheckAfter - never an empty listing that would
            // refuse every view until the file next changes.
            if (!read && known is not null)
            {
                known.CheckedUtc = now;
                return known;
            }

            var names = new Dictionary<string, BoardViewNames>(StringComparer.Ordinal);

            if (read)
            {
                foreach (MasterListingRow row in rows)
                {
                    string id = row.BoardId;

                    // The first row for a board wins - CRT shows the same one.
                    if (id.Length > 0 && !names.ContainsKey(id))
                    {
                        names[id] = new BoardViewNames(
                            BoardViewRules.Clip(row.HardwareName, BoardViewRules.NameLength) ?? string.Empty,
                            BoardViewRules.Clip(row.BoardName, BoardViewRules.NameLength) ?? string.Empty);
                    }
                }
            }

            var listing = new Listing(master, version, names, read) { CheckedUtc = now };
            this.thisListings[root] = listing;
            return listing;
        }

        private static (DateTime, long)? VersionOf(string? path)
        {
            if (path is null)
                return null;

            try
            {
                var file = new FileInfo(path);
                return file.Exists ? (file.LastWriteTimeUtc, file.Length) : null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }

        private static bool ReadListing(string masterPath, out IReadOnlyList<MasterListingRow> rows) =>
            MasterListing.TryRead(masterPath, out rows, out _);

        private sealed class Listing(string? masterPath, (DateTime, long)? version, IReadOnlyDictionary<string, BoardViewNames> names, bool isComplete)
        {
            public bool IsComplete { get; } = isComplete;

            public string? MasterPath { get; } = masterPath;

            public (DateTime, long)? Version { get; } = version;

            public IReadOnlyDictionary<string, BoardViewNames> Names { get; } = names;

            public DateTimeOffset CheckedUtc { get; set; }
        }
    }
}
