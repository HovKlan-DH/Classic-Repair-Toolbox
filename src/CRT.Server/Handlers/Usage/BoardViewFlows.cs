using System.Net;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // STORING A BATCH OF BOARD VIEWS (owner request, 2026-09-27) - POST /api/usage/board-views.
    // See CRT.Data's BoardViewContract for what a view is and why nothing stored identifies a user.
    //
    // The shape every flow here keeps: the endpoint is a rim, and every decision is below, taking its
    // store, the name listing and the country lookup as arguments so it tests with fakes.
    //
    // WHAT IS REFUSED (400), so CRT drops the report rather than sending it again: no report, a
    // batch id that is not a GUID, no version, no list of views, or more views than one report may
    // carry. Everything a report can be right about is then judged VIEW BY VIEW, and a view that
    // cannot be counted is IGNORED, never the report: a board no published listing has (a made-up
    // id, or a board since removed), a view older than BoardViewRules.MaxAge or too far ahead.
    //
    // The COUNTRY is looked up once per report, and only when something is to be stored. The
    // sender's address is handed to the lookup and to nothing else.
    //
    // *** A REPORT FROM A LOCAL NETWORK IS THE PROJECT OWNER'S OWN *** (SenderAddress says why).
    // With `countLocalNetwork` (ServerOptions.CountLocalNetworkBoardViews - owner request,
    // 2026-09-27: "For now I would like my own home usage also to count") its views are stored like
    // anybody's, each marked FromLocalNetwork, with the country of the server's own address - a
    // private address places nobody, and the sender is on the server's network. Without it the
    // report is accepted, so CRT stops sending it, and nothing is stored. Either way the outcome says
    // it came from a local network, so the endpoint can log it - the way the project owner sees their
    // own test reach the server.
    // ###########################################################################################
    public static class BoardViewFlows
    {
        public static async Task<BoardViewOutcome> RecordAsync(
            BoardViewReport? report,
            IPAddress? sender,
            Func<string, BoardViewNames?> namesOf,
            ICountryLookup countries,
            IBoardViewStore store,
            bool countLocalNetwork,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(namesOf);
            ArgumentNullException.ThrowIfNull(countries);
            ArgumentNullException.ThrowIfNull(store);

            if (report is null)
                return BoardViewOutcome.Refused("No report was sent.");

            if (!Guid.TryParse(report.BatchId, out Guid batchId) || batchId == Guid.Empty)
                return BoardViewOutcome.Refused("The report has no usable batch id.");

            string? version = BoardViewRules.Clip(report.Version, BoardViewRules.VersionLength);

            if (version is null)
                return BoardViewOutcome.Refused("The report does not say which version of CRT sent it.");

            if (report.Views is null)
                return BoardViewOutcome.Refused("The report carries no list of views.");

            if (report.Views.Count > BoardViewRules.MaxViewsPerReport)
                return BoardViewOutcome.Refused($"A report may carry at most {BoardViewRules.MaxViewsPerReport} views.");

            bool fromLocalNetwork = SenderAddress.PublicOrNull(sender) is null;

            if (fromLocalNetwork && !countLocalNetwork)
                return BoardViewOutcome.NotCountedFromLocalNetwork(report.Views.Count);

            var kept = new List<(BoardView View, string BoardId, BoardViewNames Names)>();

            foreach (BoardView? view in report.Views)
            {
                string? boardId = BoardViewRules.Clip(BoardView.IdOf(view), BoardViewRules.BoardIdLength);

                if (view is null || boardId is null || !BoardViewRules.IsCountable(view.ViewedUtc, now))
                    continue;

                if (namesOf(boardId) is BoardViewNames names)
                    kept.Add((view, boardId, names));
            }

            int ignored = report.Views.Count - kept.Count;

            // Nothing to store: nothing to look up or remember either.
            if (kept.Count == 0)
                return BoardViewOutcome.Stored(0, ignored, fromLocalNetwork);

            CountryAnswer? country = fromLocalNetwork
                ? await countries.LookupOwnAsync(cancellationToken)
                : await countries.LookupAsync(sender, cancellationToken);

            string? os = BoardViewRules.Clip(report.OsHighlevel, BoardViewRules.OsHighlevelLength);
            string? osVersion = BoardViewRules.Clip(report.OsVersion, BoardViewRules.OsVersionLength);
            string? cpu = BoardViewRules.Clip(report.Cpu, BoardViewRules.CpuLength);

            List<BoardViewRow> rows = kept
                .Select(item => new BoardViewRow(
                    BoardViewRules.StoredTime(item.View.ViewedUtc, now),
                    item.BoardId,
                    item.Names.HardwareName,
                    item.Names.BoardName,
                    version,
                    os,
                    osVersion,
                    cpu,
                    country?.Code,
                    country?.Name,
                    item.View.FromBeta,
                    fromLocalNetwork))
                .ToList();

            bool isNew = await store.RecordAsync(batchId, rows, now, cancellationToken);

            return isNew
                ? BoardViewOutcome.Stored(rows.Count, ignored, fromLocalNetwork)
                : BoardViewOutcome.Repeated(fromLocalNetwork);
        }
    }

    // ###########################################################################################
    // What became of a report. Refused carries why; Stored says how many views were stored and how
    // many ignored; IsRepeat is a batch stored before (CRT never heard the first answer) - accepted,
    // and nothing counted twice; IsFromLocalNetwork is a report from the server's own network,
    // stored or - NotCountedFromLocalNetwork - accepted with all its views ignored.
    // ###########################################################################################
    public sealed record BoardViewOutcome(bool IsRefused, string? Reason, int StoredViews, int IgnoredViews, bool IsRepeat, bool IsFromLocalNetwork = false)
    {
        public static BoardViewOutcome Refused(string reason) => new(true, reason, 0, 0, false);

        public static BoardViewOutcome Stored(int stored, int ignored, bool fromLocalNetwork = false) =>
            new(false, null, stored, ignored, false, fromLocalNetwork);

        public static BoardViewOutcome Repeated(bool fromLocalNetwork = false) => new(false, null, 0, 0, true, fromLocalNetwork);

        public static BoardViewOutcome NotCountedFromLocalNetwork(int views) => new(false, null, 0, views, false, IsFromLocalNetwork: true);
    }

    // ###########################################################################################
    // The Boards screen's numbers from a board's facts (IBoardViewStore.FactsForBoardAsync).
    // Pure, so the windows are tested.
    //
    // *** WHOLE UTC DAYS, TODAY INCLUDED. *** "The last 7 days" is today and the six before it -
    // the list's "views in 30 days" (CountSinceByBoardAsync from WindowStart) counts the same days
    // as the detail's, so the two never disagree on screen. BETA-source views count only in
    // FromBetaLast30Days; views with no country count everywhere but in TopCountries.
    // ###########################################################################################
    public static class BoardViewStatisticsRules
    {
        public const int TopCountryCount = 5;

        // The first moment of the `days`-day window ending today (UTC).
        public static DateTimeOffset WindowStart(DateTimeOffset now, int days)
        {
            DateTime today = now.UtcDateTime.Date;

            return new DateTimeOffset(today.AddDays(-(days - 1)), TimeSpan.Zero);
        }

        public static BoardViewStatistics Build(IEnumerable<BoardViewFact> facts, DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(facts);

            List<BoardViewFact> all = facts.ToList();

            DateOnly From(int days) => DateOnly.FromDateTime(BoardViewStatisticsRules.WindowStart(now, days).UtcDateTime);

            int Count(int days, bool fromBeta) =>
                all.Where(fact => fact.FromBeta == fromBeta && fact.Day >= From(days)).Sum(fact => fact.Views);

            List<BoardViewCountry> countries = all
                .Where(fact => !fact.FromBeta && fact.Day >= From(365) && !string.IsNullOrWhiteSpace(fact.CountryCode))
                .GroupBy(fact => fact.CountryCode!.Trim().ToUpperInvariant(), StringComparer.Ordinal)
                .Select(group => new BoardViewCountry(
                    group.Key,
                    group.Select(fact => fact.CountryName?.Trim()).FirstOrDefault(name => !string.IsNullOrWhiteSpace(name)) ?? group.Key,
                    group.Sum(fact => fact.Views)))
                .OrderByDescending(country => country.Views)
                .ThenBy(country => country.CountryName, StringComparer.OrdinalIgnoreCase)
                .Take(BoardViewStatisticsRules.TopCountryCount)
                .ToList();

            // Each day of the last 365 with a view, oldest first (2026-10-09) - the Statistics view's
            // graph. The BETA source's views are left out, as from every count above but its own.
            List<BoardViewDay> daily = all
                .Where(fact => !fact.FromBeta && fact.Day >= From(365))
                .GroupBy(fact => fact.Day)
                .Select(group => new BoardViewDay(group.Key, group.Sum(fact => fact.Views)))
                .Where(day => day.Views > 0)
                .OrderBy(day => day.Day)
                .ToList();

            return new BoardViewStatistics(
                Last7Days: Count(7, fromBeta: false),
                Last30Days: Count(30, fromBeta: false),
                Last365Days: Count(365, fromBeta: false),
                FromBetaLast30Days: Count(30, fromBeta: true),
                TopCountries: countries,
                Daily: daily);
        }
    }
}
