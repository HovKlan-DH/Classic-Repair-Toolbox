using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Text.Json;

namespace CRT.Server.Handlers.Usage
{
    // A country, as a board view stores it: "DK", "Denmark".
    public sealed record CountryAnswer(string Code, string Name);

    // ###########################################################################################
    // Which country an address is in - the seam BoardViewFlows is tested through.
    // ###########################################################################################
    public interface ICountryLookup
    {
        // Null when it cannot be told: a private address, no answer, or an answer that is not one.
        Task<CountryAnswer?> LookupAsync(IPAddress? address, CancellationToken cancellationToken = default);

        // The country of the SERVER's own public address - which is where a sender on its own
        // network is too (ServerOptions.CountLocalNetworkBoardViews). Null when it cannot be told.
        Task<CountryAnswer?> LookupOwnAsync(CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // The country of a board view's sender, from ip-api.com - THE SAME SERVICE THE LAUNCH CHECK-IN
    // HAS ALWAYS USED (Assets/Webserver/app-checkin), so the two tables name countries alike and the
    // Fun facts maps can put them side by side.
    //
    // *** THE ADDRESS IS NEVER STORED. *** It is sent to ip-api.com (as the check-in sends it) and
    // kept in memory for at most CacheFor, so a machine sending several batches is looked up once;
    // only the country reaches crt_board_views. The server's own country (LookupOwnAsync, for views
    // from its own network) is remembered the same way. ip-api.com's free service is plain HTTP and
    // allows 45 lookups a minute, far above what CRT's users send; asking only for three fields keeps
    // the answer small and says nothing more about the address than the country.
    //
    // A lookup that fails costs the COUNTRY, never the view: it is stored with no country.
    //
    // *** A FAILED LOOKUP PAUSES LOOKING UP, FOR EVERY ADDRESS (code review, 2026-09-29). *** A
    // failure was never remembered, so while ip-api.com was slow or down EVERY board-view request
    // waited the full 3-second timeout inside the request before its batch was stored - and a CRT
    // sending an offline backlog waited 3 seconds per batch. After a failure, lookups are skipped
    // (no country) for FailureBackoff; the next one after that asks again. A lookup cancelled
    // because the SENDER went away says nothing about ip-api.com, and pauses nothing.
    // ###########################################################################################
    public sealed class IpApiCountryLookup : ICountryLookup
    {
        private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(3) };

        public static readonly TimeSpan CacheFor = TimeSpan.FromHours(1);

        // Beyond this many addresses the memory is simply emptied - it only saves lookups.
        private const int CacheLimit = 10_000;

        // How long lookups are skipped after one fails - see the header.
        public static readonly TimeSpan FailureBackoff = TimeSpan.FromMinutes(5);

        private readonly ConcurrentDictionary<string, (CountryAnswer? Answer, DateTimeOffset Until)> thisCache = new();

        private readonly ILogger<IpApiCountryLookup> thisLogger;

        // The request and the clock - ip-api.com and the real time, unless a test supplies its own.
        private readonly Func<string, CancellationToken, Task<string>> thisFetch;
        private readonly Func<DateTimeOffset> thisNow;

        // Until when lookups are skipped after a failure (ticks, so it can be read and written from
        // many requests at once without a lock). Zero: not paused.
        private long thisPausedUntilTicks;

        public IpApiCountryLookup(ILogger<IpApiCountryLookup> logger)
            : this(logger, (url, token) => IpApiCountryLookup.Http.GetStringAsync(url, token), () => DateTimeOffset.UtcNow)
        {
        }

        // The test seam: a fetch that never reaches the network, and a clock the test moves.
        // Internal, so dependency injection only ever sees the constructor above.
        internal IpApiCountryLookup(
            ILogger<IpApiCountryLookup> logger,
            Func<string, CancellationToken, Task<string>> fetch,
            Func<DateTimeOffset> now)
        {
            this.thisLogger = logger;
            this.thisFetch = fetch;
            this.thisNow = now;
        }

        public Task<CountryAnswer?> LookupAsync(IPAddress? address, CancellationToken cancellationToken = default)
        {
            IPAddress? looked = SenderAddress.PublicOrNull(address);

            return looked is null
                ? Task.FromResult<CountryAnswer?>(null)
                : this.AskAsync(looked, cancellationToken);
        }

        // ip-api.com answers a question with no address about the address ASKING - the server's own.
        public Task<CountryAnswer?> LookupOwnAsync(CancellationToken cancellationToken = default) =>
            this.AskAsync(null, cancellationToken);

        // ###########################################################################################
        // The question put to ip-api.com: about `address`, or with none, about the server itself.
        // Only the three fields a board view stores are asked for.
        // ###########################################################################################
        public static string QueryUrl(IPAddress? address) =>
            $"http://ip-api.com/json/{address}?fields=status,countryCode,country";

        private async Task<CountryAnswer?> AskAsync(IPAddress? address, CancellationToken cancellationToken)
        {
            // The server's own address is remembered under a key no address can have.
            string key = address?.ToString() ?? "(own)";
            DateTimeOffset now = this.thisNow();

            if (this.thisCache.TryGetValue(key, out var cached) && cached.Until > now)
                return cached.Answer;

            // Paused after a failure: no country, and no 3-second wait (see the header).
            if (now.UtcTicks < Interlocked.Read(ref this.thisPausedUntilTicks))
                return null;

            CountryAnswer? answer = null;

            try
            {
                string json = await this.thisFetch(IpApiCountryLookup.QueryUrl(address), cancellationToken);

                answer = IpApiCountryLookup.ReadAnswer(json);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
            {
                // The sender went away: that says nothing about ip-api.com.
                if (cancellationToken.IsCancellationRequested)
                    return null;

                Interlocked.Exchange(ref this.thisPausedUntilTicks, (now + IpApiCountryLookup.FailureBackoff).UtcTicks);

                // The view is still stored, without a country. Said once per failure, without the
                // address - it is not to be written anywhere, the journal included.
                this.thisLogger.LogWarning(
                    "Looking up a board view's country failed: {Reason} - skipping lookups for {Minutes} minutes.",
                    ex.Message, IpApiCountryLookup.FailureBackoff.TotalMinutes);

                return null;
            }

            if (this.thisCache.Count >= IpApiCountryLookup.CacheLimit)
                this.thisCache.Clear();

            this.thisCache[key] = (answer, now + IpApiCountryLookup.CacheFor);
            return answer;
        }

        // ###########################################################################################
        // ip-api.com's answer read: a country only when it says "success" with a two-letter code
        // and a name. The code is kept in capitals, the name cut to its column (50).
        // ###########################################################################################
        public static CountryAnswer? ReadAnswer(string? json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return null;

            try
            {
                using JsonDocument document = JsonDocument.Parse(json);
                JsonElement root = document.RootElement;

                if (root.ValueKind != JsonValueKind.Object ||
                    !root.TryGetProperty("status", out JsonElement status) ||
                    status.ValueKind != JsonValueKind.String ||
                    !string.Equals(status.GetString(), "success", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                string? code = root.TryGetProperty("countryCode", out JsonElement c) && c.ValueKind == JsonValueKind.String
                    ? c.GetString()?.Trim().ToUpperInvariant()
                    : null;

                string? name = root.TryGetProperty("country", out JsonElement n) && n.ValueKind == JsonValueKind.String
                    ? n.GetString()?.Trim()
                    : null;

                if (code is null || code.Length != 2 || !code.All(character => character is >= 'A' and <= 'Z') ||
                    string.IsNullOrWhiteSpace(name))
                {
                    return null;
                }

                return new CountryAnswer(code, name.Length > 50 ? name[..50].TrimEnd() : name);
            }
            catch (JsonException)
            {
                return null;
            }
        }
    }

    // ###########################################################################################
    // A SENDER'S ADDRESS - PUBLIC, OR FROM A LOCAL NETWORK.
    //
    // A local or private address places nobody, so it is never looked up. Behind Apache, a real CRT
    // user always arrives with their public address, so a private one is the server's own network -
    // the project owner's machines. The launch check-in has always left those out (app-checkin
    // stores nothing for "192.168.*", and every Fun facts query filters them). Board views store them
    // only while ServerOptions.CountLocalNetworkBoardViews is on (owner request, 2026-09-27: "For now
    // I would like my own home usage also to count"), marked, with the country of the server's own
    // public address (ICountryLookup.LookupOwnAsync) - BoardViewFlows.
    // ###########################################################################################
    public static class SenderAddress
    {
        // ###########################################################################################
        // The address to look up, or null for one no lookup service can place: loopback, private
        // ranges (10/8, 172.16/12, 192.168/16), link-local, carrier-grade NAT (100.64/10), IPv6
        // unique-local and link-local. An IPv4 address carried in IPv6 is looked up as IPv4.
        // ###########################################################################################
        public static IPAddress? PublicOrNull(IPAddress? address)
        {
            if (address is null)
                return null;

            if (address.IsIPv4MappedToIPv6)
                address = address.MapToIPv4();

            if (IPAddress.IsLoopback(address) || address.Equals(IPAddress.Any) || address.Equals(IPAddress.IPv6Any))
                return null;

            if (address.AddressFamily == AddressFamily.InterNetwork)
            {
                byte[] bytes = address.GetAddressBytes();

                bool isPrivate =
                    bytes[0] == 10 ||
                    (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31) ||
                    (bytes[0] == 192 && bytes[1] == 168) ||
                    (bytes[0] == 169 && bytes[1] == 254) ||
                    (bytes[0] == 100 && bytes[1] >= 64 && bytes[1] <= 127) ||
                    bytes[0] == 0 ||
                    bytes[0] >= 224;

                return isPrivate ? null : address;
            }

            if (address.AddressFamily == AddressFamily.InterNetworkV6)
            {
                byte first = address.GetAddressBytes()[0];

                bool isLocal = address.IsIPv6LinkLocal || address.IsIPv6SiteLocal || address.IsIPv6Multicast || (first & 0xFE) == 0xFC;

                return isLocal ? null : address;
            }

            return null;
        }

    }
}
