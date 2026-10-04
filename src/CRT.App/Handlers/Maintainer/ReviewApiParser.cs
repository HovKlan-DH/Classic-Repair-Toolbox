using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // Reads the server's JSON answers (NewContributeStrategy.md Phase 5, task 2).
    //
    // *** SEPARATE FROM THE HTTP CLIENT FOR THE SAME REASON ReviewApiRoutes IS. *** The socket
    // half is an untested I/O boundary by this project's rules; parsing is pure and is where a
    // client actually breaks. Every method here takes a string and returns a value, so the whole
    // contract with the server is covered by tests that need no server.
    //
    // *** EVERY ANSWER IS TREATED AS UNTRUSTED, even though it comes from our own server. *** Not
    // because the server is suspected, but because a partial response, a proxy error page, or a
    // version skew between this CRT and a server it was not built against all arrive here as
    // "JSON that is not what was expected". Throwing a JsonException out of a UI event handler
    // crashes the app; answering null lets the tab say the server sent something it did not
    // understand, which is both true and actionable.
    //
    // CAMEL CASE, because that is what the endpoints emit (ASP.NET Core's default). Matching it
    // explicitly rather than relying on a case-insensitive fallback means a field renamed on the
    // server fails a test here rather than silently reading as absent.
    //
    // FILE MAP (split 2026-09-27, past ~1,500 lines): this file - sign-in, the queue, one
    // submission, decisions, messages and the shared readers; ReviewApiParser.Systems.cs - the
    // Systems screen; ReviewApiParser.Production.cs - "Beta > Prod" and unused files;
    // ReviewApiParser.Account.cs - the signed-in maintainer's own account. The records
    // they read into are ReviewApiViews.cs.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        // ###########################################################################################
        // The session from a login response, or null when the answer is not one.
        //
        // The bearer token comes from the `refreshToken` field - see ReviewSession's header for
        // why that name is a trap rather than a mistake.
        // ###########################################################################################
        public static ReviewSession? ParseLogin(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? token = ReviewApiParser.String(root, "refreshToken");

            if (string.IsNullOrWhiteSpace(token))
                return null;

            if (!root.TryGetProperty("account", out JsonElement account) ||
                account.ValueKind != JsonValueKind.Object)
            {
                return null;
            }

            // An expiry that cannot be read is treated as ALREADY EXPIRED rather than as
            // never-expiring. Being wrong the other way would have the app keep presenting a dead
            // token and reporting 401s it could have explained.
            DateTimeOffset expires = ReviewApiParser.Time(root, "expiresUtc") ?? DateTimeOffset.MinValue;

            return new ReviewSession(
                token,
                expires,
                ReviewApiParser.Long(account, "id") ?? 0,
                ReviewApiParser.String(account, "email") ?? string.Empty,
                ReviewApiParser.String(account, "displayName") ?? string.Empty);
        }

        // ###########################################################################################
        // The queue. An answer that cannot be read yields null - distinct from an EMPTY queue,
        // which is a perfectly good answer and yields an empty list.
        //
        // That distinction is the whole point: "no submissions are waiting" and "I could not ask"
        // must never look the same to a maintainer, or a broken connection reads as a clear queue
        // and the backlog goes unnoticed.
        // ###########################################################################################
        public static ReviewQueueResponse? ParseQueue(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            if (!root.TryGetProperty("submissions", out JsonElement submissions) ||
                submissions.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<ReviewQueueRow>();

            foreach (JsonElement element in submissions.EnumerateArray())
            {
                ReviewQueueRow? row = ReviewApiParser.ParseRow(element);

                // A single unreadable row is SKIPPED rather than failing the whole queue: one
                // malformed record must not hide every other contribution waiting behind it.
                if (row is not null)
                    rows.Add(row);
            }

            return new ReviewQueueResponse(
                ReviewApiParser.Bool(root, "canPublish") ?? false,
                rows,
                ReviewApiParser.Bool(root, "isAdministrator") ?? false);
        }

        // ###########################################################################################
        // The administrator's two lists (Phase 6 roles): every system with its maintainers, and
        // every account. Each is a plain array under one property; a row that cannot be read is
        // skipped rather than failing the list, for the same reason a bad queue row is.
        // ###########################################################################################
        public static ReviewSystemsResponse? ParseSystems(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("systems", out JsonElement systems) ||
                systems.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<ReviewSystemRow>();

            foreach (JsonElement element in systems.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.Object)
                    continue;

                string? systemId = ReviewApiParser.String(element, "systemId");

                if (string.IsNullOrWhiteSpace(systemId))
                    continue;

                var maintainers = new List<MaintainerRow>();

                if (element.TryGetProperty("maintainers", out JsonElement list) && list.ValueKind == JsonValueKind.Array)
                {
                    foreach (JsonElement maintainer in list.EnumerateArray())
                    {
                        long? accountId = maintainer.ValueKind == JsonValueKind.Object
                            ? ReviewApiParser.Long(maintainer, "accountId")
                            : null;

                        if (accountId is null)
                            continue;

                        maintainers.Add(new MaintainerRow(
                            accountId.Value,
                            ReviewApiParser.String(maintainer, "displayName") ?? string.Empty,
                            ReviewApiParser.String(maintainer, "email") ?? string.Empty));
                    }
                }

                rows.Add(new ReviewSystemRow(
                    systemId,
                    ReviewApiParser.String(element, "manufacturer") ?? string.Empty,
                    ReviewApiParser.String(element, "hardware") ?? string.Empty,
                    ReviewApiParser.String(element, "board") ?? string.Empty,
                    ReviewApiParser.String(element, "currentRevision"),
                    ReviewApiParser.Bool(element, "isAccepting") ?? true,
                    maintainers));
            }

            return new ReviewSystemsResponse(rows);
        }

        // ###########################################################################################
        // BETA to production (2026-09-25).
        //
        // The list says whether the server can do it at all ("configured"), so "nothing waiting"
        // and "this server cannot publish to production" read differently. The plan's files are
        // CRT.Data's PromotionFile, deserialised as the same record the server wrote.
        // ###########################################################################################
        public static ProductionListResponse? ParseProductionList(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("systems", out JsonElement systems) ||
                systems.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<ProductionSystemRow>();

            foreach (JsonElement element in systems.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.Object)
                    continue;

                string? systemId = ReviewApiParser.String(element, "systemId");

                if (string.IsNullOrWhiteSpace(systemId))
                    continue;

                rows.Add(new ProductionSystemRow(
                    systemId,
                    ReviewApiParser.String(element, "manufacturer") ?? string.Empty,
                    ReviewApiParser.String(element, "hardware") ?? string.Empty,
                    ReviewApiParser.String(element, "board") ?? string.Empty,
                    ReviewApiParser.String(element, "betaRevision"),
                    ReviewApiParser.String(element, "betaContentHash") ?? string.Empty,
                    ReviewApiParser.String(element, "productionRevision"),
                    ReviewApiParser.Time(element, "productionPublishedUtc"),

                    // The BETA badge (2026-09-27). Absent from an older server: null, counted as yours.
                    ReviewApiParser.Bool(element, "awaitsYou"),

                    // A contributor whose work this carries discarded their draft (2026-09-28).
                    ReviewApiParser.Bool(element, "carriesDiscardedDraft")));
            }

            return new ProductionListResponse(ReviewApiParser.Bool(root, "configured") ?? false, rows);
        }

        public static ReviewAccountsResponse? ParseAccounts(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                !root.TryGetProperty("accounts", out JsonElement accounts) ||
                accounts.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            var rows = new List<ReviewAccountRow>();

            foreach (JsonElement element in accounts.EnumerateArray())
            {
                if (element.ValueKind != JsonValueKind.Object)
                    continue;

                long? id = ReviewApiParser.Long(element, "id");

                if (id is null)
                    continue;

                rows.Add(new ReviewAccountRow(
                    id.Value,
                    ReviewApiParser.String(element, "email") ?? string.Empty,
                    ReviewApiParser.String(element, "displayName") ?? string.Empty,
                    ReviewApiParser.Bool(element, "isAdministrator") ?? false,
                    ReviewApiParser.Bool(element, "isVerified") ?? false,
                    ReviewApiParser.Bool(element, "isLocked") ?? false));
            }

            return new ReviewAccountsResponse(rows);
        }

        // ###########################################################################################
        // One submission's detail, as GET /api/review/submissions/{id} answers it.
        //
        // *** THE CHANGE SUMMARY IS PARSED SEPARATELY FROM THE REST, and a missing one is not a
        // failure. *** The server answers `changes: null` when a submission's payload could not be
        // loaded, which is a real state a maintainer must be able to open and read the findings for.
        // Refusing the whole response because the summary is absent would hide the very
        // explanation they came for.
        //
        // The shapes below were verified by serialising a real ReviewChangeSummary through
        // System.Text.Json with ASP.NET Core's camelCase policy, not by reading the record
        // declarations and guessing - the records carry computed properties (`hasChanges`,
        // `totalChanges`) that also appear on the wire, which a guess would have missed.
        // ###########################################################################################
        public static ReviewSubmissionDetail? ParseSubmission(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            if (!root.TryGetProperty("submission", out JsonElement submission))
                return null;

            ReviewQueueRow? row = ReviewApiParser.ParseRow(submission);

            if (row is null)
                return null;

            return new ReviewSubmissionDetail(
                row,
                ReviewApiParser.Bool(root, "canPublish") ?? false,
                root.TryGetProperty("changes", out JsonElement changes)
                    ? ReviewApiParser.ParseChangeSummary(changes)
                    : null,
                ReviewApiParser.ParseFindings(root),
                ReviewApiParser.ParseAssets(root),
                ReviewApiParser.ParsePublishedFiles(root),
                ReviewApiParser.ParsePublishedHashes(root),
                ReviewApiParser.ParseSchematicImages(root),
                ReviewApiParser.ParseSubmittedFiles(root),
                ReviewApiParser.ParseApproval(root),
                ReviewApiParser.ParseRemovals(root),
                ReviewApiParser.ParseAmendment(root),
                ReviewApiParser.ParseContributor(root));
        }

        // ###########################################################################################
        // The maintainer's table (2026-09-25): CRT.Data's ReviewTableData, as the server wrote it.
        // Null when the answer carries no submitted rows - nothing to open.
        // ###########################################################################################
        public static ReviewTableData? ParseTable(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                ReviewTableData? table = root.Deserialize<ReviewTableData>(ReviewApiParser.FactOptions);

                return table?.Submitted is null ? null : table;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        // What saving an amendment answered: the new version, and any warnings it raised.
        public static ReviewAmendResult? ParseAmend(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            long? version = ReviewApiParser.Long(root, "version");

            return version is null or < 1
                ? null
                : new ReviewAmendResult((int)version.Value, ReviewApiParser.ParseFindings(root));
        }

        // Who last changed the submission in the Maintainer tab's table, or null.
        private static ReviewAmendmentView? ParseAmendment(JsonElement root)
        {
            if (!root.TryGetProperty("amendment", out JsonElement raw) || raw.ValueKind != JsonValueKind.Object)
                return null;

            long? version = ReviewApiParser.Long(raw, "version");

            return version is null
                ? null
                : new ReviewAmendmentView(
                    (int)version.Value,
                    ReviewApiParser.String(raw, "by") ?? string.Empty,
                    ReviewApiParser.Time(raw, "atUtc"));
        }

        // ###########################################################################################
        // One fact per submitted file: whose it is, whether a row uses it, and what is published
        // at its path now (security review, 2026-09-25).
        //
        // *** DESERIALISED INTO THE SERVER'S OWN TYPE. *** SubmittedFileFact lives in CRT.Data and
        // the server serialises exactly that record, so a renamed property fails to compile rather
        // than silently arriving blank. The scope travels as its NAME (the enum carries its own
        // converter), so a member added later cannot shift every value by one.
        //
        // A missing or malformed field yields an EMPTY list, never a failure - an older server that
        // does not send it degrades to the image-only comparison it had before, and one bad entry is
        // dropped rather than costing the maintainer the whole screen.
        // ###########################################################################################
        private static IReadOnlyList<SubmittedFileFact> ParseSubmittedFiles(JsonElement root)
        {
            if (!root.TryGetProperty("submittedFiles", out JsonElement raw) ||
                raw.ValueKind != JsonValueKind.Array)
            {
                return [];
            }

            var facts = new List<SubmittedFileFact>();

            foreach (JsonElement element in raw.EnumerateArray())
            {
                try
                {
                    SubmittedFileFact? fact = element.Deserialize<SubmittedFileFact>(ReviewApiParser.FactOptions);

                    if (fact is not null &&
                        !string.IsNullOrWhiteSpace(fact.Path) &&
                        !string.IsNullOrWhiteSpace(fact.Sha256))
                    {
                        facts.Add(fact);
                    }
                }
                catch (JsonException)
                {
                    // One unreadable entry is dropped; the rest still describe the submission.
                }
            }

            return facts;
        }

        private static readonly JsonSerializerOptions FactOptions = new()
        {
            PropertyNameCaseInsensitive = true
        };

        // ###########################################################################################
        // Path -> SHA-256 for the published files the submission also carries (added 2026-09-23).
        //
        // *** THIS IS WHAT TURNS "1178 images to compare" INTO THE ONE THAT ACTUALLY CHANGED. ***
        // ReviewImageComparison.Plan drops a pair only when it is handed the published hash and it
        // matches the submitted one. Until the server sent this, nothing handed it anything, so
        // every file present on both sides was drawn as replaced - all of them identical.
        //
        // A missing or malformed field yields an EMPTY map, never a failure: the planner then falls
        // back to treating every shared path as replaced, which shows a maintainer too much rather
        // than too little. An older server that does not send the field degrades to exactly the
        // behaviour it had before.
        //
        // Ordinal keys: the paths are matched against a file list the server built on a
        // case-sensitive tree.
        // ###########################################################################################
        private static IReadOnlyDictionary<string, string> ParsePublishedHashes(JsonElement root)
        {
            var hashes = new Dictionary<string, string>(StringComparer.Ordinal);

            if (!root.TryGetProperty("publishedHashes", out JsonElement raw) ||
                raw.ValueKind != JsonValueKind.Object)
            {
                return hashes;
            }

            foreach (JsonProperty entry in raw.EnumerateObject())
            {
                if (entry.Value.ValueKind != JsonValueKind.String)
                    continue;

                string? hash = entry.Value.GetString();

                if (!string.IsNullOrWhiteSpace(entry.Name) && !string.IsNullOrWhiteSpace(hash))
                    hashes[entry.Name] = hash;
            }

            return hashes;
        }

        // ###########################################################################################
        // Which image file each schematic is drawn from (task 4's moved highlight).
        //
        // *** A HIGHLIGHT NAMES A SCHEMATIC, NOT A FILE. *** Its row key is
        // SchematicName|BoardLabel, so without this mapping the app knows a rectangle moved on
        // "Sheet 1" and cannot find the picture of Sheet 1 to draw it on.
        //
        // Ordinal keys: schematic names are matched against a row key the server built, and
        // folding case here would pair a highlight with the wrong board on a tree that is
        // case-sensitive everywhere else.
        // ###########################################################################################
        private static IReadOnlyDictionary<string, string> ParseSchematicImages(JsonElement root)
        {
            var images = new Dictionary<string, string>(StringComparer.Ordinal);

            if (!root.TryGetProperty("schematicImages", out JsonElement raw) ||
                raw.ValueKind != JsonValueKind.Object)
            {
                return images;
            }

            foreach (JsonProperty entry in raw.EnumerateObject())
            {
                if (entry.Value.ValueKind != JsonValueKind.String)
                    continue;

                string? file = entry.Value.GetString();

                if (!string.IsNullOrWhiteSpace(file))
                    images[entry.Name] = file;
            }

            return images;
        }

        // ###########################################################################################
        // The files the PUBLISHED board references, which the image comparison pairs against.
        //
        // *** THIS COMES FROM THE SERVER AND IS NOT DERIVED FROM `changes`. *** A summary's row
        // keys are natural keys (for an image, BoardLabel|Region|Pin|Name), not file paths, so
        // comparing them against the manifest's paths matches nothing - every image would read as
        // an addition and no deletion would ever appear, with the panel drawing perfectly either
        // way. The first version of the review window made exactly that mistake.
        // ###########################################################################################
        private static IReadOnlyList<string> ParsePublishedFiles(JsonElement root)
        {
            if (!root.TryGetProperty("publishedFiles", out JsonElement raw) ||
                raw.ValueKind != JsonValueKind.Array)
            {
                return [];
            }

            var files = new List<string>();

            foreach (JsonElement file in raw.EnumerateArray())
            {
                if (file.ValueKind == JsonValueKind.String)
                {
                    string? value = file.GetString();

                    if (!string.IsNullOrWhiteSpace(value))
                        files.Add(value);
                }
            }

            return files;
        }

        // ###########################################################################################
        // The answer to a DECISION (task 5).
        //
        // The state is what the submission became, which the window uses to say what happened
        // rather than assuming the button it drew did what it said. An approval also carries the
        // published revision, which is what the contributor's next submission will diff against.
        // ###########################################################################################
        public static ReviewDecisionResult? ParseDecision(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? state = ReviewApiParser.String(root, "state");

            // Without a state there is nothing to report and no way to know what happened, so this
            // is not a usable answer - the same rule ParseSubmission applies to a missing id.
            if (string.IsNullOrWhiteSpace(state))
                return null;

            return new ReviewDecisionResult(
                state,
                ReviewApiParser.String(root, "revision") ?? string.Empty,
                ReviewApiParser.ParseRoles(root, "waitingFor"),
                ReviewApiParser.ParseStrings(root, "removedFiles"));
        }

        // ###########################################################################################
        // The files a publish would remove (2026-09-25) - CRT.Data's FileRemovalPreview, the same
        // record the server wrote. Null from an older server, which removed nothing.
        // ###########################################################################################
        private static FileRemovalPreview? ParseRemovals(JsonElement root)
        {
            if (!root.TryGetProperty("removals", out JsonElement raw) || raw.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                FileRemovalPreview? preview = raw.Deserialize<FileRemovalPreview>(ReviewApiParser.FactOptions);

                return preview is null ? null : preview with { Files = preview.Files ?? [] };
            }
            catch (JsonException)
            {
                return null;
            }
        }

        // An array of strings; anything else in it is skipped.
        private static IReadOnlyList<string> ParseStrings(JsonElement root, string name)
        {
            if (!root.TryGetProperty(name, out JsonElement raw) || raw.ValueKind != JsonValueKind.Array)
                return [];

            return raw.EnumerateArray()
                .Where(item => item.ValueKind == JsonValueKind.String)
                .Select(item => item.GetString()!)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList();
        }

        // ###########################################################################################
        // The two-person approval (2026-09-25). CRT.Data's ApprovalStatus, deserialised as the same
        // record the server wrote; null when absent or unreadable, never a failure of the answer
        // around it.
        // ###########################################################################################
        // Who sent the submission and how their other submissions went - CRT.Data's
        // ReviewContributorFacts, as the server wrote it. Null from an older server.
        private static ReviewContributorFacts? ParseContributor(JsonElement root)
        {
            if (!root.TryGetProperty("contributor", out JsonElement raw) || raw.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                return raw.Deserialize<ReviewContributorFacts>(ReviewApiParser.FactOptions);
            }
            catch (JsonException)
            {
                return null;
            }
        }

        private static ApprovalStatus? ParseApproval(JsonElement root)
        {
            if (!root.TryGetProperty("approval", out JsonElement raw) || raw.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                return raw.Deserialize<ApprovalStatus>(ReviewApiParser.FactOptions);
            }
            catch (JsonException)
            {
                return null;
            }
        }

        // An array of role NAMES ("Maintainer", "Administrator"); unknown names are skipped.
        private static IReadOnlyList<ApproverRole> ParseRoles(JsonElement root, string name)
        {
            if (!root.TryGetProperty(name, out JsonElement raw) || raw.ValueKind != JsonValueKind.Array)
                return [];

            var roles = new List<ApproverRole>();

            foreach (JsonElement item in raw.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String &&
                    Enum.TryParse(item.GetString(), ignoreCase: true, out ApproverRole role))
                {
                    roles.Add(role);
                }
            }

            return roles;
        }

        // ###########################################################################################
        // The server's own `message` out of a success body.
        //
        // Used for the password-reset answer, whose whole content is a sentence written for the
        // person reading it. Taking the server's wording rather than writing a second copy here
        // means the two cannot drift into saying different things about the same event.
        // ###########################################################################################
        public static string? ParseMessage(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? message = ReviewApiParser.String(root, "message");

            return string.IsNullOrWhiteSpace(message) ? null : message;
        }

        // ###########################################################################################
        // The server's own error sentence out of a refusal body.
        //
        // Returned rather than replaced with a generic message because the server's wording is
        // written for the person reading it - "write at least 10 characters" is actionable where
        // "the request was refused" is not.
        // ###########################################################################################
        public static string? ParseError(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? error = ReviewApiParser.String(root, "error");

            return string.IsNullOrWhiteSpace(error) ? null : error;
        }

        // ###########################################################################################
        // The first sentence out of an `errors` ARRAY.
        //
        // *** A THIRD REFUSAL SHAPE, AND IT IS NOT INTERCHANGEABLE WITH THE OTHER TWO. *** Most
        // refusals answer `{"message":...}` or `{"error":...}`, but a rejected PASSWORD answers
        // `{"errors":[...]}` - AccountRules.ValidatePassword returns a list, because a password
        // can fail several rules at once. A caller reading only `message` sees nothing at all
        // there and falls back to a generic sentence, which is the one case where the server
        // actually had something worth saying ("at least 12 characters").
        //
        // ALL of them are joined rather than only the first, despite the name: being told to fix
        // one rule, fixing it, and then being told about the next is a loop with no visible end.
        // ###########################################################################################
        public static string? ParseFirstError(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            if (!root.TryGetProperty("errors", out JsonElement errors) ||
                errors.ValueKind != JsonValueKind.Array)
            {
                return null;
            }

            string joined = string.Join(
                " ",
                errors.EnumerateArray()
                    .Where(entry => entry.ValueKind == JsonValueKind.String)
                    .Select(entry => entry.GetString())
                    .Where(text => !string.IsNullOrWhiteSpace(text)));

            return string.IsNullOrWhiteSpace(joined) ? null : joined;
        }

        // ###########################################################################################
        // The files a submission carries, so the Maintainer tab can fetch their bytes (task 4).
        //
        // Only the PATH and the HASH are taken. The manifest carries a great deal more, and
        // lifting all of it into a view type would make every future field on the contract look
        // like something this screen depends on.
        //
        // *** A FILE MISSING EITHER VALUE IS SKIPPED, NOT DEFAULTED. *** A blank hash would build
        // a fetch URL that the server refuses, and the panel would report a failure that is really
        // a malformed manifest. Skipping it means the rest of the submission still reviews - the
        // same "one bad row must not conceal the others" rule the queue parser follows.
        // ###########################################################################################
        private static ReviewSubmissionAssets ParseAssets(JsonElement root)
        {
            var files = new List<ReviewSubmittedFile>();

            if (root.TryGetProperty("manifest", out JsonElement manifest) &&
                manifest.ValueKind == JsonValueKind.Object &&
                manifest.TryGetProperty("files", out JsonElement raw) &&
                raw.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement file in raw.EnumerateArray())
                {
                    if (file.ValueKind != JsonValueKind.Object)
                        continue;

                    string path = ReviewApiParser.String(file, "path") ?? string.Empty;
                    string hash = ReviewApiParser.String(file, "sha256") ?? string.Empty;

                    if (path.Length == 0 || hash.Length == 0)
                        continue;

                    files.Add(new ReviewSubmittedFile(path, hash));
                }
            }

            return new ReviewSubmissionAssets(files);
        }

        // ###########################################################################################
        // The change summary. Null when the server sent none, or sent something unreadable.
        //
        // *** SECTIONS WITH NO CHANGES ARE DROPPED HERE. *** The server sends all ten every time
        // (they are cheap - the whole summary is about 1.4 KB), but the maintainer's screen only
        // ever shows the ones that changed. Filtering at the edge means nothing downstream has to
        // remember to.
        // ###########################################################################################
        private static ReviewChangeSummaryView? ParseChangeSummary(JsonElement element)
        {
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            var sections = new List<ReviewSectionView>();

            if (element.TryGetProperty("sections", out JsonElement raw) &&
                raw.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement section in raw.EnumerateArray())
                {
                    ReviewSectionView? parsed = ReviewApiParser.ParseSection(section);

                    if (parsed is not null && parsed.TotalChanges > 0)
                        sections.Add(parsed);
                }
            }

            return new ReviewChangeSummaryView(
                ReviewApiParser.Bool(element, "isNewSystem") ?? false,
                sections);
        }

        private static ReviewSectionView? ParseSection(JsonElement element)
        {
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            string? name = ReviewApiParser.String(element, "section");

            // A section with no name cannot be shown as anything, so it is not a usable row.
            if (string.IsNullOrWhiteSpace(name))
                return null;

            return new ReviewSectionView(
                name,
                ReviewApiParser.Strings(element, "added"),
                ReviewApiParser.Strings(element, "removed"),
                ReviewApiParser.Strings(element, "changed"),
                ReviewApiParser.ParseRenames(element),
                ReviewApiParser.ParseFieldChanges(element));
        }

        // ###########################################################################################
        // The per-row field diffs (task 4's "text rows as a field-level diff").
        //
        // Keyed by row, mirroring the server's own dictionary. A row whose fields could not be read
        // simply has no entry, and the window falls back to naming the row without detail - which
        // is what the screen did before this existed, so the degradation is to the previous
        // behaviour rather than to a blank.
        // ###########################################################################################
        private static IReadOnlyDictionary<string, IReadOnlyList<ReviewFieldChangeView>> ParseFieldChanges(
            JsonElement element)
        {
            var result = new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>(StringComparer.Ordinal);

            if (!element.TryGetProperty("fieldChanges", out JsonElement raw) ||
                raw.ValueKind != JsonValueKind.Object)
            {
                return result;
            }

            foreach (JsonProperty row in raw.EnumerateObject())
            {
                if (row.Value.ValueKind != JsonValueKind.Array)
                    continue;

                var fields = new List<ReviewFieldChangeView>();

                foreach (JsonElement field in row.Value.EnumerateArray())
                {
                    if (field.ValueKind != JsonValueKind.Object)
                        continue;

                    string? name = ReviewApiParser.String(field, "field");

                    // Without a name there is nothing to label the change with, and "something
                    // changed to 251715-01" is not a reviewable statement.
                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    fields.Add(new ReviewFieldChangeView(
                        name,

                        // A CLEARED field arrives as an empty string, which is exactly what should
                        // be shown - deleting information is a real edit and the hardest to spot.
                        ReviewApiParser.String(field, "before") ?? string.Empty,
                        ReviewApiParser.String(field, "after") ?? string.Empty));
                }

                if (fields.Count > 0)
                    result[row.Name] = fields;
            }

            return result;
        }

        private static IReadOnlyList<ReviewRenameView> ParseRenames(JsonElement element)
        {
            if (!element.TryGetProperty("renamed", out JsonElement raw) ||
                raw.ValueKind != JsonValueKind.Array)
            {
                return [];
            }

            var renames = new List<ReviewRenameView>();

            foreach (JsonElement item in raw.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                    continue;

                string? from = ReviewApiParser.String(item, "from");
                string? to = ReviewApiParser.String(item, "to");

                if (string.IsNullOrWhiteSpace(from) || string.IsNullOrWhiteSpace(to))
                    continue;

                renames.Add(new ReviewRenameView(
                    from,
                    to,
                    ReviewApiParser.Bool(item, "alsoChanged") ?? false));
            }

            return renames;
        }

        // ###########################################################################################
        // The validation findings.
        //
        // *** A FINDING WITH NO MESSAGE IS STILL KEPT, carrying its code. *** These are the reasons
        // a submission was rejected or flagged, and dropping one because its prose is missing
        // would leave a maintainer looking at a shorter list than the truth - on the one screen
        // whose job is to say what is wrong.
        // ###########################################################################################
        private static IReadOnlyList<ReviewFindingView> ParseFindings(JsonElement root)
        {
            if (!root.TryGetProperty("findings", out JsonElement raw) ||
                raw.ValueKind != JsonValueKind.Array)
            {
                return [];
            }

            var findings = new List<ReviewFindingView>();

            foreach (JsonElement item in raw.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                    continue;

                findings.Add(new ReviewFindingView(
                    ReviewApiParser.String(item, "code") ?? string.Empty,
                    ReviewApiParser.String(item, "subject") ?? string.Empty,
                    ReviewApiParser.String(item, "message") ?? string.Empty,

                    // Severity is serialised as a NUMBER by default (the enum's value), so both
                    // spellings are accepted: a future JsonStringEnumConverter on the server would
                    // otherwise silently turn every finding into a warning.
                    ReviewApiParser.IsError(item)));
            }

            return findings;
        }

        private static bool IsError(JsonElement element)
        {
            if (!element.TryGetProperty("severity", out JsonElement severity))
                return false;

            // ValidationSeverity: Warning = 0, Error = 1.
            if (severity.ValueKind == JsonValueKind.Number)
                return severity.TryGetInt32(out int value) && value == 1;

            return severity.ValueKind == JsonValueKind.String &&
                   string.Equals(severity.GetString(), "Error", StringComparison.OrdinalIgnoreCase);
        }

        private static IReadOnlyList<string> Strings(JsonElement element, string name)
        {
            if (!element.TryGetProperty(name, out JsonElement raw) || raw.ValueKind != JsonValueKind.Array)
                return [];

            var values = new List<string>();

            foreach (JsonElement item in raw.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String && item.GetString() is string text)
                    values.Add(text);
            }

            return values;
        }

        private static ReviewQueueRow? ParseRow(JsonElement element)
        {
            if (element.ValueKind != JsonValueKind.Object)
                return null;

            long? id = ReviewApiParser.Long(element, "id");

            // Without an id the row cannot be opened, so it is not a usable queue entry however
            // much else it carries.
            if (id is null)
                return null;

            return new ReviewQueueRow(
                id.Value,
                ReviewApiParser.String(element, "systemId") ?? string.Empty,
                ReviewApiParser.String(element, "state") ?? string.Empty,
                ReviewApiParser.String(element, "summary") ?? string.Empty,
                ReviewApiParser.String(element, "contactEmail") ?? string.Empty,
                ReviewApiParser.Time(element, "createdUtc"),
                ReviewApiParser.Bool(element, "touchesSharedFiles") ?? false,

                // The queue list's two badges (2026-09-26). Absent from an older server - null, so
                // no badge is shown rather than a wrong one.
                ReviewApiParser.Bool(element, "isNewSystem"),
                ReviewApiParser.Bool(element, "awaitsYou"),

                // The contributor discarded their own draft after sending it (2026-09-28).
                ReviewApiParser.Time(element, "draftDiscardedUtc"));
        }

        // -----------------------------------------------------------------------------------
        // Reading one value, never throwing
        // -----------------------------------------------------------------------------------

        private static JsonElement Root(string? json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return default;

            try
            {
                using JsonDocument document = JsonDocument.Parse(json);

                // Cloned because the document is disposed on the way out of this method, and an
                // element borrowed from a disposed document throws when it is next read - which
                // would be a crash at a random later line rather than here.
                return document.RootElement.Clone();
            }
            catch (JsonException)
            {
                // A proxy error page, a truncated body, or a server this build does not
                // understand. All of them are "not the answer expected", and none is worth
                // crashing a UI event handler over.
                return default;
            }
        }

        private static string? String(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.String
                ? value.GetString()
                : null;

        private static long? Long(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement value) &&
            value.ValueKind == JsonValueKind.Number &&
            value.TryGetInt64(out long parsed)
                ? parsed
                : null;

        private static bool? Bool(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement value) &&
            value.ValueKind is JsonValueKind.True or JsonValueKind.False
                ? value.GetBoolean()
                : null;

        private static DateTimeOffset? Time(JsonElement element, string name) =>
            element.TryGetProperty(name, out JsonElement value) &&
            value.ValueKind == JsonValueKind.String &&
            value.TryGetDateTimeOffset(out DateTimeOffset parsed)
                ? parsed
                : null;
    }
}
