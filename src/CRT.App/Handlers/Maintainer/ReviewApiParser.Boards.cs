using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ReviewApiParser, part two: the Boards screen - the overview, one board's detail, its place in
    // CRT's drop-down lists, and accepting an invitation. Split out of ReviewApiParser.cs when it passed
    // the project's ~1,500 lines (code review, 2026-09-27); the rules in that file's header hold here.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        // ###########################################################################################
        // The "Boards" screen (2026-09-27). Read into CRT.Data's own records - they are plain facts
        // with nothing computed, so a view type of their own would only be a second copy - but field
        // by field and forgivingly, like everything here: a row with no board id is skipped, a list
        // that is missing is empty, and an unreadable answer is null rather than an empty screen.
        // ###########################################################################################
        public static BoardOverviewAnswer? ParseBoardOverview(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("boards", out JsonElement boards) ||
                boards.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<BoardOverviewEntry>();

            foreach (JsonElement element in boards.EnumerateArray())
            {
                if (ReviewApiParser.ParseBoardEntry(element) is BoardOverviewEntry entry)
                    rows.Add(entry);
            }

            return new BoardOverviewAnswer(rows);
        }

        // ###########################################################################################
        // A board's Board data and Files views (2026-10-03). The table is read into CRT.Data's own
        // record, as a submission's table is (ParseTable) - its rows are the board's - and refused
        // without a fingerprint, which a change could not be sent back with. The files are the
        // submission tree's own entries (ParseSubmissionFiles' rule: entries with no path dropped).
        // ###########################################################################################
        public static BoardTableAnswer? ParseBoardTable(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                BoardTableAnswer? table = root.Deserialize<BoardTableAnswer>(ReviewApiParser.FactOptions);

                return table is null || table.Rows is null ||
                       string.IsNullOrWhiteSpace(table.BoardId) || string.IsNullOrWhiteSpace(table.Fingerprint)
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
        public static BoardEditCheckAnswer? ParseBoardEditCheck(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("removals", out JsonElement removals) ||
                removals.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            return new BoardEditCheckAnswer(ReviewApiParser.ParseStrings(root, "removals"));
        }

        // What sending a change answered: the submission it became, its warnings, and whether it was
        // published to BETA - at which revision and removing what - or why not.
        public static BoardEditResult? ParseBoardEdit(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object || ReviewApiParser.Long(root, "submissionId") is not long id || id < 1)
                return null;

            return new BoardEditResult(
                id,
                ReviewApiParser.ParseFindings(root),
                ReviewApiParser.Bool(root, "published") == true,
                ReviewApiParser.String(root, "revision"),
                ReviewApiParser.ParseStrings(root, "removedFiles"),
                ReviewApiParser.String(root, "notPublishedReason"));
        }

        public static BoardFilesAnswer? ParseBoardFiles(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                BoardFilesAnswer? answer = root.Deserialize<BoardFilesAnswer>(ReviewApiParser.FactOptions);

                if (answer is null || string.IsNullOrWhiteSpace(answer.BoardId))
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

        public static BoardDetailAnswer? ParseBoardDetail(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("board", out JsonElement board) ||
                ReviewApiParser.ParseBoardEntry(board) is not BoardOverviewEntry entry)
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

            var contributors = new List<BoardContributorEntry>();

            foreach (JsonElement contributor in ReviewApiParser.Objects(root, "contributors"))
            {
                contributors.Add(new BoardContributorEntry(
                    ReviewApiParser.String(contributor, "email"),
                    ReviewApiParser.String(contributor, "name"),
                    (int)(ReviewApiParser.Long(contributor, "accepted") ?? 0),
                    (int)(ReviewApiParser.Long(contributor, "waiting") ?? 0),
                    (int)(ReviewApiParser.Long(contributor, "changesRequested") ?? 0),
                    (int)(ReviewApiParser.Long(contributor, "rejected") ?? 0),
                    ReviewApiParser.Time(contributor, "lastSubmittedUtc")));
            }

            var submissions = new List<BoardSubmissionEntry>();

            foreach (JsonElement submission in ReviewApiParser.Objects(root, "submissions"))
            {
                // Without an id or a state it cannot be told apart or described - not a usable row.
                if (ReviewApiParser.Long(submission, "id") is not long id ||
                    ReviewApiParser.String(submission, "state") is not string state)
                {
                    continue;
                }

                submissions.Add(new BoardSubmissionEntry(
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
            List<BoardHistoryEntry>? history = null;

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

                    history.Add(new BoardHistoryEntry(
                        at,
                        kind,
                        ReviewApiParser.String(item, "who"),
                        ReviewApiParser.Long(item, "submissionId"),
                        ReviewApiParser.String(item, "detail"),
                        ReviewApiParser.String(item, "note")));
                }
            }

            // No address sent, as the account does not maintain the board (2026-10-05) - absent
            // from an older server, which sent them all.
            bool addressesHidden =
                root.TryGetProperty("addressesHidden", out JsonElement hidden) && hidden.ValueKind == JsonValueKind.True;

            // Whether BETA's data may be changed now, and why not (code review, 2026-10-09) - absent
            // from an older server and for a board BETA does not hold: then not known.
            return new BoardDetailAnswer(
                entry, maintainers, contributors, submissions, invitations, history, ReviewApiParser.ParseViewStatistics(root),
                addressesHidden,
                ReviewApiParser.Bool(root, "mayEdit"),
                ReviewApiParser.String(root, "mayNotEditReason"));
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
        public static BoardOrderAnswer? ParseBoardOrder(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object || ReviewApiParser.Bool(root, "betaChanged") is not bool betaChanged)
                return null;

            return new BoardOrderAnswer(
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

            List<string> boards = root.TryGetProperty("boardIds", out JsonElement ids) && ids.ValueKind == JsonValueKind.Array
                ? ids.EnumerateArray().Where(id => id.ValueKind == JsonValueKind.String).Select(id => id.GetString()!).ToList()
                : [];

            return new AcceptInvitationAnswer(email, boards, ReviewApiParser.String(root, "message") ?? string.Empty);
        }

        // ###########################################################################################
        // Where a new board goes in the drop-down lists (2026-09-27). Forgiving like the rest: a row
        // with no board id or workbook is skipped, and a placement with no hardware name is no
        // placement. An answer without its "rows" list is unreadable (null), never an empty list -
        // the screen would otherwise say nothing is listed on the strength of a bad answer.
        // ###########################################################################################
        public static BoardListingAnswer? ParseBoardListing(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("rows", out JsonElement listed) ||
                listed.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<BoardListingRow>();

            foreach (JsonElement row in ReviewApiParser.Objects(root, "rows"))
            {
                string? boardId = ReviewApiParser.String(row, "boardId");
                string? workbook = ReviewApiParser.String(row, "excelDataFile");

                if (string.IsNullOrWhiteSpace(boardId) || string.IsNullOrWhiteSpace(workbook))
                    continue;

                rows.Add(new BoardListingRow(
                    boardId,
                    ReviewApiParser.String(row, "hardwareName") ?? string.Empty,
                    ReviewApiParser.String(row, "boardName") ?? string.Empty,
                    workbook));
            }

            var unlisted = new List<UnlistedBoardEntry>();

            foreach (JsonElement entry in ReviewApiParser.Objects(root, "unlisted"))
            {
                string? boardId = ReviewApiParser.String(entry, "boardId");

                if (string.IsNullOrWhiteSpace(boardId))
                    continue;

                string hardware = ReviewApiParser.String(entry, "hardware") ?? string.Empty;
                string board = ReviewApiParser.String(entry, "board") ?? string.Empty;

                unlisted.Add(new UnlistedBoardEntry(
                    boardId,
                    ReviewApiParser.String(entry, "manufacturer") ?? string.Empty,
                    hardware,
                    board,
                    ReviewApiParser.Bool(entry, "inBeta") ?? false,
                    ReviewApiParser.Bool(entry, "canPlace") ?? false,
                    ReviewApiParser.ParsePlacement(entry, "placement"),
                    ReviewApiParser.ParsePlacement(entry, "suggested") ?? new BoardPlacement(hardware, board, string.Empty, null)));
            }

            return new BoardListingAnswer(ReviewApiParser.Bool(root, "hasList") ?? false, rows, unlisted);
        }

        public static SetPlacementAnswer? ParseSetPlacement(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object || ReviewApiParser.ParsePlacement(root, "placement") is not BoardPlacement placement)
                return null;

            return new SetPlacementAnswer(
                placement,
                ReviewApiParser.Bool(root, "listedInBeta") ?? false,
                ReviewApiParser.String(root, "message") ?? string.Empty);
        }

        private static BoardPlacement? ParsePlacement(JsonElement parent, string name)
        {
            if (!parent.TryGetProperty(name, out JsonElement element) || element.ValueKind != JsonValueKind.Object)
                return null;

            string? hardware = ReviewApiParser.String(element, "hardwareName");

            if (string.IsNullOrWhiteSpace(hardware))
                return null;

            string? after = ReviewApiParser.String(element, "afterExcelDataFile");

            return new BoardPlacement(
                hardware,
                ReviewApiParser.String(element, "boardName") ?? string.Empty,
                ReviewApiParser.String(element, "notes") ?? string.Empty,
                string.IsNullOrWhiteSpace(after) ? null : after);
        }

        private static BoardOverviewEntry? ParseBoardEntry(JsonElement element)
        {
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            string? boardId = ReviewApiParser.String(element, "boardId");

            if (string.IsNullOrWhiteSpace(boardId))
                return null;

            return new BoardOverviewEntry(
                boardId,
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

                // What BETA holds (2026-10-04) - the Boards screen's table is read again when it
                // moves. Absent from an older server, or for a board nothing has published.
                ReviewApiParser.String(element, "betaContentHash"),

                // Whether each source's drop-down list names it (2026-10-04) - absent (null) when
                // the server could not read that list, which is never said to be "not listed".
                ReviewApiParser.Bool(element, "listedInBeta"),
                ReviewApiParser.Bool(element, "listedInStable"),

                // How its submissions went (2026-10-09) - absent from an older server, and then
                // its line says nothing about them.
                ReviewApiParser.ParseSubmissionCounts(element));
        }

        private static BoardSubmissionCounts? ParseSubmissionCounts(JsonElement element)
        {
            if (!element.TryGetProperty("submissionCounts", out JsonElement counts) || counts.ValueKind != JsonValueKind.Object)
                return null;

            int Count(string name) => (int)Math.Max(0, ReviewApiParser.Long(counts, name) ?? 0);

            return new BoardSubmissionCounts(Count("total"), Count("waiting"), Count("rejected"), Count("inBeta"), Count("inStable"));
        }

        // ###########################################################################################
        // How often CRT users look at a board (2026-09-27). Null when the answer has none - an older
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

            // Each day's views (2026-10-09), for the graph - absent from an older server (null:
            // no graph), a day that does not read as a date left out.
            List<BoardViewDay>? daily = null;

            if (views.TryGetProperty("daily", out JsonElement days) && days.ValueKind == JsonValueKind.Array)
            {
                daily = [];

                foreach (JsonElement day in ReviewApiParser.Objects(views, "daily"))
                {
                    if (ReviewApiParser.String(day, "day") is string text &&
                        DateOnly.TryParseExact(text, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateOnly date))
                    {
                        daily.Add(new BoardViewDay(date, (int)Math.Max(0, ReviewApiParser.Long(day, "views") ?? 0)));
                    }
                }
            }

            return new BoardViewStatistics(
                (int)(ReviewApiParser.Long(views, "last7Days") ?? 0),
                (int)(ReviewApiParser.Long(views, "last30Days") ?? 0),
                (int)(ReviewApiParser.Long(views, "last365Days") ?? 0),
                (int)(ReviewApiParser.Long(views, "fromBetaLast30Days") ?? 0),
                countries,
                daily);
        }

        // The objects in an array property - none when it is missing or not an array.
        private static IEnumerable<JsonElement> Objects(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement raw) && raw.ValueKind == JsonValueKind.Array
                ? raw.EnumerateArray().Where(item => item.ValueKind == JsonValueKind.Object).ToList()
                : [];
    }
}
