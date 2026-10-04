using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ReviewApiParser, part two: the Systems screen - the overview, one system's detail, its place in
    // CRT's drop-down lists, and accepting an invitation. Split out of ReviewApiParser.cs when it passed
    // the project's ~1,500 lines (code review, 2026-09-27); the rules in that file's header hold here.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        // ###########################################################################################
        // The "Systems" screen (2026-09-27). Read into CRT.Data's own records - they are plain facts
        // with nothing computed, so a view type of their own would only be a second copy - but field
        // by field and forgivingly, like everything here: a row with no system id is skipped, a list
        // that is missing is empty, and an unreadable answer is null rather than an empty screen.
        // ###########################################################################################
        public static SystemOverviewAnswer? ParseSystemOverview(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("systems", out JsonElement systems) ||
                systems.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<SystemOverviewEntry>();

            foreach (JsonElement element in systems.EnumerateArray())
            {
                if (ReviewApiParser.ParseSystemEntry(element) is SystemOverviewEntry entry)
                    rows.Add(entry);
            }

            return new SystemOverviewAnswer(rows);
        }

        // ###########################################################################################
        // A system's Board data and Files views (2026-10-03). The table is read into CRT.Data's own
        // record, as a submission's table is (ParseTable) - its rows are the board's - and refused
        // without a fingerprint, which a change could not be sent back with. The files are the
        // submission tree's own entries (ParseSubmissionFiles' rule: entries with no path dropped).
        // ###########################################################################################
        public static SystemTableAnswer? ParseSystemTable(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                SystemTableAnswer? table = root.Deserialize<SystemTableAnswer>(ReviewApiParser.FactOptions);

                return table is null || table.Rows is null ||
                       string.IsNullOrWhiteSpace(table.SystemId) || string.IsNullOrWhiteSpace(table.Fingerprint)
                    ? null
                    : table;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        // ###########################################################################################
        // What checking a change answered: the files publishing it would remove. Refused without the
        // list - an answer that does not say is not "nothing", and the maintainer must be shown it.
        // ###########################################################################################
        public static SystemEditCheckAnswer? ParseSystemEditCheck(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("removals", out JsonElement removals) ||
                removals.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            return new SystemEditCheckAnswer(ReviewApiParser.ParseStrings(root, "removals"));
        }

        // What sending a change answered: the submission it became, its warnings, and whether it was
        // published to BETA - at which revision and removing what - or why not.
        public static SystemEditResult? ParseSystemEdit(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object || ReviewApiParser.Long(root, "submissionId") is not long id || id < 1)
                return null;

            return new SystemEditResult(
                id,
                ReviewApiParser.ParseFindings(root),
                ReviewApiParser.Bool(root, "published") == true,
                ReviewApiParser.String(root, "revision"),
                ReviewApiParser.ParseStrings(root, "removedFiles"),
                ReviewApiParser.String(root, "notPublishedReason"));
        }

        public static SystemFilesAnswer? ParseSystemFiles(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                SystemFilesAnswer? answer = root.Deserialize<SystemFilesAnswer>(ReviewApiParser.FactOptions);

                if (answer is null || string.IsNullOrWhiteSpace(answer.SystemId))
                    return null;

                return answer with
                {
                    Files = (answer.Files ?? []).Where(file => file is not null && !string.IsNullOrWhiteSpace(file.Path)).ToList()
                };
            }
            catch (JsonException)
            {
                return null;
            }
        }

        public static SystemDetailAnswer? ParseSystemDetail(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("system", out JsonElement system) ||
                ReviewApiParser.ParseSystemEntry(system) is not SystemOverviewEntry entry)
            {
                return null;
            }

            var maintainers = new List<PoolMaintainerEntry>();

            foreach (JsonElement maintainer in ReviewApiParser.Objects(root, "maintainers"))
            {
                if (ReviewApiParser.Long(maintainer, "accountId") is long accountId)
                {
                    maintainers.Add(new PoolMaintainerEntry(
                        accountId,
                        ReviewApiParser.String(maintainer, "displayName") ?? string.Empty,
                        ReviewApiParser.String(maintainer, "email") ?? string.Empty));
                }
            }

            var contributors = new List<SystemContributorEntry>();

            foreach (JsonElement contributor in ReviewApiParser.Objects(root, "contributors"))
            {
                contributors.Add(new SystemContributorEntry(
                    ReviewApiParser.String(contributor, "email"),
                    ReviewApiParser.String(contributor, "name"),
                    (int)(ReviewApiParser.Long(contributor, "accepted") ?? 0),
                    (int)(ReviewApiParser.Long(contributor, "waiting") ?? 0),
                    (int)(ReviewApiParser.Long(contributor, "changesRequested") ?? 0),
                    (int)(ReviewApiParser.Long(contributor, "rejected") ?? 0),
                    ReviewApiParser.Time(contributor, "lastSubmittedUtc")));
            }

            var submissions = new List<SystemSubmissionEntry>();

            foreach (JsonElement submission in ReviewApiParser.Objects(root, "submissions"))
            {
                // Without an id or a state it cannot be told apart or described - not a usable row.
                if (ReviewApiParser.Long(submission, "id") is not long id ||
                    ReviewApiParser.String(submission, "state") is not string state)
                {
                    continue;
                }

                submissions.Add(new SystemSubmissionEntry(
                    id,
                    ReviewApiParser.String(submission, "contactEmail"),
                    ReviewApiParser.String(submission, "summary"),
                    state,
                    ReviewApiParser.Time(submission, "createdUtc") ?? default,
                    ReviewApiParser.Time(submission, "decidedUtc"),
                    ReviewApiParser.String(submission, "decisionComment"),

                    // Its contributor discarded their own draft since (2026-09-28).
                    ReviewApiParser.Time(submission, "draftDiscardedUtc"),

                    // What it changed as it went into BETA (2026-10-04) - none when not recorded.
                    ReviewApiParser.ParseChanges(submission)));
            }

            // The invitations nobody has accepted - sent to an administrator only; absent is none.
            List<MaintainerInvitationEntry>? invitations = null;

            if (root.TryGetProperty("invitations", out JsonElement invited) && invited.ValueKind == JsonValueKind.Array)
            {
                invitations = [];

                foreach (JsonElement invitation in ReviewApiParser.Objects(root, "invitations"))
                {
                    if (ReviewApiParser.Long(invitation, "id") is long id &&
                        ReviewApiParser.String(invitation, "email") is string email &&
                        !string.IsNullOrWhiteSpace(email))
                    {
                        invitations.Add(new MaintainerInvitationEntry(
                            id,
                            email,
                            ReviewApiParser.Time(invitation, "invitedUtc") ?? default,
                            ReviewApiParser.Time(invitation, "expiresUtc") ?? default));
                    }
                }
            }

            // What has happened to it (2026-09-27); absent from an older server, which the panel
            // answers by showing the submissions alone, as before.
            List<SystemHistoryEntry>? history = null;

            if (root.TryGetProperty("history", out JsonElement happened) && happened.ValueKind == JsonValueKind.Array)
            {
                history = [];

                foreach (JsonElement item in ReviewApiParser.Objects(root, "history"))
                {
                    // Without a date or a kind it cannot be placed or said - not a usable line.
                    if (ReviewApiParser.Time(item, "atUtc") is not DateTimeOffset at ||
                        ReviewApiParser.String(item, "event") is not string kind ||
                        string.IsNullOrWhiteSpace(kind))
                    {
                        continue;
                    }

                    history.Add(new SystemHistoryEntry(
                        at,
                        kind,
                        ReviewApiParser.String(item, "who"),
                        ReviewApiParser.Long(item, "submissionId"),
                        ReviewApiParser.String(item, "detail"),
                        ReviewApiParser.String(item, "note")));
                }
            }

            return new SystemDetailAnswer(entry, maintainers, contributors, submissions, invitations, history, ReviewApiParser.ParseViewStatistics(root));
        }

        // ###########################################################################################
        // What a submission changed as it went into BETA (2026-10-04) - CRT.Data's SubmissionChanges,
        // read whole. One that does not read is no summary rather than a failed screen; the lists a
        // missing field leaves null are read as empty, so the History view never meets a null.
        // ###########################################################################################
        private static SubmissionChanges? ParseChanges(JsonElement submission)
        {
            if (!submission.TryGetProperty("changes", out JsonElement element) || element.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                SubmissionChanges? changes = element.Deserialize<SubmissionChanges>(ReviewApiParser.FactOptions);

                if (changes is null)
                    return null;

                FileChanges files = changes.Files ?? FileChanges.None;

                return changes with
                {
                    Sections = (changes.Sections ?? [])
                        .Where(section => section is not null && !string.IsNullOrWhiteSpace(section.Section))
                        .Select(section => section with
                        {
                            Added = section.Added ?? [],
                            Changed = (section.Changed ?? []).Where(row => row is not null).Select(row => row with { Fields = row.Fields ?? [] }).ToList(),
                            Removed = section.Removed ?? [],
                            Renamed = (section.Renamed ?? []).Where(row => row is not null).ToList()
                        })
                        .ToList(),
                    Files = files with
                    {
                        Added = files.Added ?? [],
                        Replaced = files.Replaced ?? [],
                        Removed = files.Removed ?? []
                    }
                };
            }
            catch (JsonException)
            {
                return null;
            }
        }

        // What ordering the drop-down lists did (2026-10-04). Without BETA's answer it is not one.
        public static SystemOrderAnswer? ParseSystemOrder(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object || ReviewApiParser.Bool(root, "betaChanged") is not bool betaChanged)
                return null;

            return new SystemOrderAnswer(
                betaChanged,
                ReviewApiParser.Bool(root, "stableChanged"),
                ReviewApiParser.String(root, "problem"));
        }

        // What accepting an invitation did. Without the address it cannot fill the sign-in box and
        // is not the answer it claims to be.
        public static AcceptInvitationAnswer? ParseAcceptInvitation(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                ReviewApiParser.String(root, "email") is not string email ||
                string.IsNullOrWhiteSpace(email))
            {
                return null;
            }

            List<string> systems = root.TryGetProperty("systemIds", out JsonElement ids) && ids.ValueKind == JsonValueKind.Array
                ? ids.EnumerateArray().Where(id => id.ValueKind == JsonValueKind.String).Select(id => id.GetString()!).ToList()
                : [];

            return new AcceptInvitationAnswer(email, systems, ReviewApiParser.String(root, "message") ?? string.Empty);
        }

        // ###########################################################################################
        // Where a new system goes in the drop-down lists (2026-09-27). Forgiving like the rest: a row
        // with no system id or workbook is skipped, and a placement with no hardware name is no
        // placement. An answer without its "rows" list is unreadable (null), never an empty list -
        // the screen would otherwise say nothing is listed on the strength of a bad answer.
        // ###########################################################################################
        public static SystemListingAnswer? ParseSystemListing(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("rows", out JsonElement listed) ||
                listed.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<SystemListingRow>();

            foreach (JsonElement row in ReviewApiParser.Objects(root, "rows"))
            {
                string? systemId = ReviewApiParser.String(row, "systemId");
                string? workbook = ReviewApiParser.String(row, "excelDataFile");

                if (string.IsNullOrWhiteSpace(systemId) || string.IsNullOrWhiteSpace(workbook))
                    continue;

                rows.Add(new SystemListingRow(
                    systemId,
                    ReviewApiParser.String(row, "hardwareName") ?? string.Empty,
                    ReviewApiParser.String(row, "boardName") ?? string.Empty,
                    workbook));
            }

            var unlisted = new List<UnlistedSystemEntry>();

            foreach (JsonElement entry in ReviewApiParser.Objects(root, "unlisted"))
            {
                string? systemId = ReviewApiParser.String(entry, "systemId");

                if (string.IsNullOrWhiteSpace(systemId))
                    continue;

                string hardware = ReviewApiParser.String(entry, "hardware") ?? string.Empty;
                string board = ReviewApiParser.String(entry, "board") ?? string.Empty;

                unlisted.Add(new UnlistedSystemEntry(
                    systemId,
                    ReviewApiParser.String(entry, "manufacturer") ?? string.Empty,
                    hardware,
                    board,
                    ReviewApiParser.Bool(entry, "inBeta") ?? false,
                    ReviewApiParser.Bool(entry, "canPlace") ?? false,
                    ReviewApiParser.ParsePlacement(entry, "placement"),
                    ReviewApiParser.ParsePlacement(entry, "suggested") ?? new SystemPlacement(hardware, board, string.Empty, null)));
            }

            return new SystemListingAnswer(ReviewApiParser.Bool(root, "hasList") ?? false, rows, unlisted);
        }

        public static SetPlacementAnswer? ParseSetPlacement(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object || ReviewApiParser.ParsePlacement(root, "placement") is not SystemPlacement placement)
                return null;

            return new SetPlacementAnswer(
                placement,
                ReviewApiParser.Bool(root, "listedInBeta") ?? false,
                ReviewApiParser.String(root, "message") ?? string.Empty);
        }

        private static SystemPlacement? ParsePlacement(JsonElement parent, string name)
        {
            if (!parent.TryGetProperty(name, out JsonElement element) || element.ValueKind != JsonValueKind.Object)
                return null;

            string? hardware = ReviewApiParser.String(element, "hardwareName");

            if (string.IsNullOrWhiteSpace(hardware))
                return null;

            string? after = ReviewApiParser.String(element, "afterExcelDataFile");

            return new SystemPlacement(
                hardware,
                ReviewApiParser.String(element, "boardName") ?? string.Empty,
                ReviewApiParser.String(element, "notes") ?? string.Empty,
                string.IsNullOrWhiteSpace(after) ? null : after);
        }

        private static SystemOverviewEntry? ParseSystemEntry(JsonElement element)
        {
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            string? systemId = ReviewApiParser.String(element, "systemId");

            if (string.IsNullOrWhiteSpace(systemId))
                return null;

            return new SystemOverviewEntry(
                systemId,
                ReviewApiParser.String(element, "manufacturer") ?? string.Empty,
                ReviewApiParser.String(element, "hardware") ?? string.Empty,
                ReviewApiParser.String(element, "board") ?? string.Empty,
                ReviewApiParser.Bool(element, "inBeta") ?? false,

                // Null is the server saying it has no production tree to look in - kept as null.
                ReviewApiParser.Bool(element, "inProduction"),
                ReviewApiParser.Bool(element, "isAwaitingProduction") ?? false,
                ReviewApiParser.Bool(element, "isAccepting") ?? true,
                ReviewApiParser.String(element, "betaRevision"),
                ReviewApiParser.String(element, "productionRevision"),
                ReviewApiParser.Time(element, "productionPublishedUtc"),
                (int)(ReviewApiParser.Long(element, "maintainerCount") ?? 0),

                // How often CRT users looked at it in 30 days (2026-09-27); absent from an older
                // server, or when the server could not count - kept as null, never shown as 0.
                ReviewApiParser.Long(element, "viewsLast30Days") is long views ? (int)views : null,

                // What BETA holds (2026-10-04) - the Systems screen's table is read again when it
                // moves. Absent from an older server, or for a board nothing has published.
                ReviewApiParser.String(element, "betaContentHash"),

                // Whether each source's drop-down list names it (2026-10-04) - absent (null) when
                // the server could not read that list, which is never said to be "not listed".
                ReviewApiParser.Bool(element, "listedInBeta"),
                ReviewApiParser.Bool(element, "listedInStable"));
        }

        // ###########################################################################################
        // How often CRT users look at a system (2026-09-27). Null when the answer has none - an older
        // server, or counts it could not read - so the panel leaves the section out rather than
        // claiming nobody looks at the board.
        // ###########################################################################################
        private static BoardViewStatistics? ParseViewStatistics(JsonElement root)
        {
            if (!root.TryGetProperty("views", out JsonElement views) || views.ValueKind != JsonValueKind.Object)
                return null;

            var countries = new List<BoardViewCountry>();

            foreach (JsonElement country in ReviewApiParser.Objects(views, "topCountries"))
            {
                if (ReviewApiParser.String(country, "countryCode") is string code && !string.IsNullOrWhiteSpace(code))
                {
                    countries.Add(new BoardViewCountry(
                        code,
                        ReviewApiParser.String(country, "countryName") ?? code,
                        (int)(ReviewApiParser.Long(country, "views") ?? 0)));
                }
            }

            return new BoardViewStatistics(
                (int)(ReviewApiParser.Long(views, "last7Days") ?? 0),
                (int)(ReviewApiParser.Long(views, "last30Days") ?? 0),
                (int)(ReviewApiParser.Long(views, "last365Days") ?? 0),
                (int)(ReviewApiParser.Long(views, "fromBetaLast30Days") ?? 0),
                countries);
        }

        // The objects in an array property - none when it is missing or not an array.
        private static IEnumerable<JsonElement> Objects(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement raw) && raw.ValueKind == JsonValueKind.Array
                ? raw.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.Object).ToList()
                : [];
    }
}
