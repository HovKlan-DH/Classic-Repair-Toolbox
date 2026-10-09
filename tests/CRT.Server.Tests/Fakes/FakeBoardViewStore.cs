using System.Data.Common;
using System.Net;
using CRT.Server.Handlers.Usage;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // An in-memory IBoardViewStore, so the board-view flows test with no database. It behaves like
    // MySqlBoardViewStore where that matters: a batch id is stored ONCE, and its second copy writes
    // nothing at all; counts leave BETA-source views out.
    //
    // FailWith makes every read throw, as a database that cannot be reached does - the Boards
    // screen must still answer then.
    // ###########################################################################################
    public sealed class FakeBoardViewStore : IBoardViewStore
    {
        public List<BoardViewRow> Rows { get; } = [];

        public HashSet<Guid> Batches { get; } = [];

        public bool FailReads { get; set; }

        public Task<bool> RecordAsync(Guid batchId, IReadOnlyList<BoardViewRow> rows, DateTimeOffset receivedUtc, CancellationToken cancellationToken = default)
        {
            if (!this.Batches.Add(batchId))
                return Task.FromResult(false);

            this.Rows.AddRange(rows);
            return Task.FromResult(true);
        }

        public Task<IReadOnlyDictionary<string, int>> CountSinceByBoardAsync(DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            if (this.FailReads)
                throw new UnreachableDatabase();

            IReadOnlyDictionary<string, int> counts = this.Rows
                .Where(row => !row.FromBeta && row.ViewedUtc >= sinceUtc)
                .GroupBy(row => row.BoardId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

            return Task.FromResult(counts);
        }

        public Task<IReadOnlyList<BoardViewFact>> FactsForBoardAsync(string boardId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            if (this.FailReads)
                throw new UnreachableDatabase();

            IReadOnlyList<BoardViewFact> facts = this.Rows
                .Where(row => string.Equals(row.BoardId, boardId, StringComparison.Ordinal) && row.ViewedUtc >= sinceUtc)
                .GroupBy(row => (Day: DateOnly.FromDateTime(row.ViewedUtc.UtcDateTime), row.FromBeta, row.CountryCode, row.CountryName))
                .Select(group => new BoardViewFact(group.Key.Day, group.Key.FromBeta, group.Key.CountryCode, group.Key.CountryName, group.Count()))
                .ToList();

            return Task.FromResult(facts);
        }

        public bool FailDeletes { get; set; }

        public Task<int> DeleteForBoardAsync(string boardId, CancellationToken cancellationToken = default)
        {
            if (this.FailDeletes)
                throw new UnreachableDatabase();

            return Task.FromResult(this.Rows.RemoveAll(row => string.Equals(row.BoardId, boardId, StringComparison.Ordinal)));
        }

        // A view as the flow stores one - for tests that fill the store directly.
        public static BoardViewRow Row(string boardId, DateTimeOffset at, string? country = "DK", bool fromBeta = false) =>
            new(at, boardId, "Commodore 64", "250407", "CRT 2026.10.0", "Windows", "Microsoft Windows 10.0.19045", "64-bit",
                country, country switch { "DK" => "Denmark", "DE" => "Germany", "US" => "United States", "SE" => "Sweden", _ => country },
                fromBeta, FromLocalNetwork: false);

        private sealed class UnreachableDatabase : DbException
        {
            public UnreachableDatabase()
                : base("The database cannot be reached.")
            {
            }
        }
    }

    // A country lookup that answers one country for every address and another for the server's own
    // (Sweden unless told), and remembers what it was asked.
    public sealed class FakeCountryLookup(CountryAnswer? answer = null, CountryAnswer? own = null) : ICountryLookup
    {
        public static readonly CountryAnswer Sweden = new("SE", "Sweden");

        public List<IPAddress?> Asked { get; } = [];

        public int AskedOwn { get; private set; }

        public Task<CountryAnswer?> LookupAsync(IPAddress? address, CancellationToken cancellationToken = default)
        {
            this.Asked.Add(address);
            return Task.FromResult(answer);
        }

        public Task<CountryAnswer?> LookupOwnAsync(CancellationToken cancellationToken = default)
        {
            this.AskedOwn++;
            return Task.FromResult<CountryAnswer?>(own ?? FakeCountryLookup.Sweden);
        }
    }
}
