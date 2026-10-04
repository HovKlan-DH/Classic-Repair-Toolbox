using Handlers.DataHandling;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // The calls counted since the last write, in memory (owner request, 2026-10-04 - see migration
    // 0018 and ApiUsageFlow). Every request that reaches a route counts here (ApiUsageRules.Count,
    // from Program's pipeline), and ApiUsageFlusher takes the lot every few minutes and adds it to
    // crt_api_calls - so a busy minute is one database write, not one per request.
    //
    // A plain lock, not a concurrent dictionary: a count is a dictionary lookup and an addition, the
    // service sees a few requests a second at most, and TakeAll must swap the whole set atomically.
    //
    // *** BOUNDED. *** Routes are the server's own patterns and versions are cut to
    // ApiUsageRules.MaximumVersionLength, but a User-Agent is the sender's to write - so once
    // MaximumKeys different day/route/version keys are waiting, a new version counts as
    // ApiUsageVersion.Other instead of growing the set further.
    //
    // *** AND BOUNDED PER DAY, NOT ONLY PER WRITE (code review, 2026-10-04). *** The key cap alone
    // emptied with every write, so "CRT 9.9.1", "CRT 9.9.2" ... sent all day became up to 10,000
    // new rows of crt_api_calls every five minutes. So at most MaximumVersionsPerDay versions are
    // counted apart in one UTC day, remembered across writes; any further one that day counts as
    // ApiUsageVersion.Other. Real use is a few dozen versions at most, so a day only reaches it
    // when somebody makes versions up - which then costs that day's figures their detail, never the
    // table its size. NotCrt and Other themselves are never held to it.
    //
    // *** A WRITE AND A RESET NEVER OVERLAP (code review, 2026-10-04). *** HoldWritesAsync is held by
    // ApiUsageFlusher from taking the tallies to putting back a failed write, and by a reset of the
    // contribution data around its transaction and the Clear after it. Without it, tallies taken
    // just before a reset were written - or put back and written later - after it, so crt_api_calls
    // held calls from before the reset that was meant to empty it.
    // ###########################################################################################
    public sealed class ApiUsageCounter
    {
        public const int MaximumKeys = 10_000;

        public const int MaximumVersionsPerDay = 100;

        private readonly object thisGate = new();
        private Dictionary<ApiUsageKey, ApiUsageTally> thisTallies = [];

        // One write into crt_api_calls at a time, a reset's included - see the header.
        private readonly SemaphoreSlim thisWriteGate = new(1, 1);

        // The versions counted apart on thisVersionsDay - kept across TakeAll, unlike the tallies.
        private readonly HashSet<string> thisVersionsToday = new(StringComparer.Ordinal);
        private DateOnly thisVersionsDay;

        public void Count(string method, string route, string version, DateTimeOffset nowUtc)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(method);
            ArgumentException.ThrowIfNullOrWhiteSpace(route);
            ArgumentException.ThrowIfNullOrWhiteSpace(version);

            DateOnly day = DateOnly.FromDateTime(nowUtc.UtcDateTime);

            lock (this.thisGate)
            {
                var key = new ApiUsageKey(day, method, route, this.Admit(version, day));

                if (!this.thisTallies.ContainsKey(key) && this.thisTallies.Count >= ApiUsageCounter.MaximumKeys)
                    key = key with { Version = ApiUsageVersion.Other };

                this.Add(key, 1, nowUtc);
            }
        }

        // Everything counted so far, and the counter empty again.
        public IReadOnlyList<ApiUsageTally> TakeAll()
        {
            lock (this.thisGate)
            {
                Dictionary<ApiUsageKey, ApiUsageTally> taken = this.thisTallies;
                this.thisTallies = [];

                return [.. taken.Values];
            }
        }

        // ###########################################################################################
        // Tallies a write could not store, added back - to whatever has been counted since - so the
        // next write tries them again and nothing is lost to a database that was briefly away. The
        // key cap does not apply: these were already counted once.
        // ###########################################################################################
        public void PutBack(IEnumerable<ApiUsageTally> tallies)
        {
            ArgumentNullException.ThrowIfNull(tallies);

            lock (this.thisGate)
            {
                foreach (ApiUsageTally tally in tallies)
                    this.Add(new ApiUsageKey(tally.Day, tally.Method, tally.Route, tally.Version), tally.Calls, tally.LastUtc);
            }
        }

        // ###########################################################################################
        // Holds every write of the tallies off until the returned value is disposed - a write that
        // has already begun is waited for first. Counting goes on meanwhile; only writing waits.
        // ###########################################################################################
        public async Task<IDisposable> HoldWritesAsync(CancellationToken cancellationToken = default)
        {
            await this.thisWriteGate.WaitAsync(cancellationToken).ConfigureAwait(false);

            return new Releaser(this.thisWriteGate);
        }

        // Forgets everything not yet written - a reset of the contribution data does this once its
        // transaction is done, holding the writes (HoldWritesAsync), so the calls from before it
        // are not written after it. The day's versions start again with it: the reset deleted
        // their rows.
        public void Clear()
        {
            lock (this.thisGate)
            {
                this.thisTallies = [];
                this.thisVersionsToday.Clear();
            }
        }

        // ###########################################################################################
        // The version a call counts under: itself while today has room for it or already counts it,
        // else ApiUsageVersion.Other. A new UTC day starts the set again; a call stamped with the
        // day before, arriving just after midnight, is held to the new day's set.
        // ###########################################################################################
        private string Admit(string version, DateOnly day)
        {
            if (string.Equals(version, ApiUsageVersion.NotCrt, StringComparison.Ordinal) ||
                string.Equals(version, ApiUsageVersion.Other, StringComparison.Ordinal))
            {
                return version;
            }

            if (day > this.thisVersionsDay)
            {
                this.thisVersionsDay = day;
                this.thisVersionsToday.Clear();
            }

            if (this.thisVersionsToday.Contains(version))
                return version;

            if (this.thisVersionsToday.Count >= ApiUsageCounter.MaximumVersionsPerDay)
                return ApiUsageVersion.Other;

            this.thisVersionsToday.Add(version);
            return version;
        }

        private void Add(ApiUsageKey key, long calls, DateTimeOffset lastUtc)
        {
            this.thisTallies[key] = this.thisTallies.TryGetValue(key, out ApiUsageTally? existing)
                ? existing with
                {
                    Calls = existing.Calls + calls,
                    LastUtc = lastUtc > existing.LastUtc ? lastUtc : existing.LastUtc
                }
                : new ApiUsageTally(key.Day, key.Method, key.Route, key.Version, calls, lastUtc);
        }
    }

    file sealed class Releaser(SemaphoreSlim gate) : IDisposable
    {
        private SemaphoreSlim? thisGate = gate;

        public void Dispose() => Interlocked.Exchange(ref this.thisGate, null)?.Release();
    }

    public readonly record struct ApiUsageKey(DateOnly Day, string Method, string Route, string Version);

    // One day's calls of one route by one CRT version.
    public sealed record ApiUsageTally(DateOnly Day, string Method, string Route, string Version, long Calls, DateTimeOffset LastUtc);
}
