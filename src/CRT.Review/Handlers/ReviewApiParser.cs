using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;

namespace CRT.Review.Handlers
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
    // version skew between this app and a server it was not built against all arrive here as
    // "JSON that is not what was expected". Throwing a JsonException out of a UI event handler
    // crashes the app; answering null lets the window say the server sent something it did not
    // understand, which is both true and actionable.
    //
    // CAMEL CASE, because that is what the endpoints emit (ASP.NET Core's default). Matching it
    // explicitly rather than relying on a case-insensitive fallback means a field renamed on the
    // server fails a test here rather than silently reading as absent.
    // ###########################################################################################
    public static class ReviewApiParser
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
                token!,
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
        // must never look the same to a reviewer, or a broken connection reads as a clear queue
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
                rows);
        }

        // ###########################################################################################
        // One submission's detail, as GET /api/review/submissions/{id} answers it.
        //
        // *** THE CHANGE SUMMARY IS PARSED SEPARATELY FROM THE REST, and a missing one is not a
        // failure. *** The server answers `changes: null` when a submission's payload could not be
        // loaded, which is a real state a reviewer must be able to open and read the findings for.
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
                ReviewApiParser.ParseSchematicImages(root));
        }

        // ###########################################################################################
        // Path -> SHA-256 for the published files the submission also carries (added 2026-09-23).
        //
        // *** THIS IS WHAT TURNS "1178 images to compare" INTO THE ONE THAT ACTUALLY CHANGED. ***
        // ReviewImageComparison.Plan drops a pair only when it is handed the published hash and it
        // matches the submitted one. Until the server sent this, nothing handed it anything, so
        // every file present on both sides was drawn as replaced - all of them identical.
        //
        // A missing or malformed field yields an EMPTY map, never a failure: the planner then falls
        // back to treating every shared path as replaced, which shows a reviewer too much rather
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
                ReviewApiParser.String(root, "revision") ?? string.Empty);
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
        // The files a submission carries, so the review app can fetch their bytes (task 4).
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
        // (they are cheap - the whole summary is about 1.4 KB), but the reviewer's screen only
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
                name!,
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
                        name!,

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
                    from!,
                    to!,
                    ReviewApiParser.Bool(item, "alsoChanged") ?? false));
            }

            return renames;
        }

        // ###########################################################################################
        // The validation findings.
        //
        // *** A FINDING WITH NO MESSAGE IS STILL KEPT, carrying its code. *** These are the reasons
        // a submission was rejected or flagged, and dropping one because its prose is missing
        // would leave a reviewer looking at a shorter list than the truth - on the one screen
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
                ReviewApiParser.Time(element, "createdUtc"));
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

    public sealed record ReviewQueueResponse(bool CanPublish, IReadOnlyList<ReviewQueueRow> Submissions);

    // ###########################################################################################
    // One submission as the review window holds it.
    //
    // *** THESE ARE VIEW TYPES, NOT THE SERVER'S OWN RECORDS. *** CRT.Data's ReviewChangeSummary
    // could have been deserialised directly and that was considered; it was rejected because the
    // server's record carries computed properties and a shape that exists to be COMPUTED, whereas
    // what the window needs is a shape that exists to be DRAWN - sections already filtered to the
    // ones that changed, renames already paired. Sharing the type would also mean any future
    // change to the comparison's internals became a wire-format change by accident.
    //
    // Changes is NULL when the server could not build a summary (an unloadable payload). That is
    // distinct from a summary with no changes, and the window says something different for each.
    // ###########################################################################################
    public sealed record ReviewSubmissionDetail(
        ReviewQueueRow Submission,
        bool CanPublish,
        ReviewChangeSummaryView? Changes,
        IReadOnlyList<ReviewFindingView> Findings,

        // The files this submission carries, for task 4's visual comparison. Never null - an empty
        // list is a rows-only change, which is the commonest contribution there is, and a null
        // here would make every caller guard against a case that simply means "no files".
        ReviewSubmissionAssets Assets,

        // The files the PUBLISHED board references, which Assets is compared against. Supplied by
        // the server rather than derived here - see ParsePublishedFiles.
        IReadOnlyList<string> PublishedFiles,

        // Path -> SHA-256 for the published files the submission also carries, so byte-identical
        // pairs are dropped from the comparison. Empty when the server did not send it. See
        // ParsePublishedHashes.
        IReadOnlyDictionary<string, string> PublishedHashes,

        // Schematic name to the image file it is drawn from, so a moved highlight can be put back
        // on its own board. See ParseSchematicImages.
        IReadOnlyDictionary<string, string> SchematicImages);

    public sealed record ReviewChangeSummaryView(
        bool IsNewSystem,
        IReadOnlyList<ReviewSectionView> Sections)
    {
        public int TotalChanges => this.Sections.Sum(section => section.TotalChanges);

        public bool HasChanges => this.TotalChanges > 0;
    }

    public sealed record ReviewSectionView(
        string Section,
        IReadOnlyList<string> Added,
        IReadOnlyList<string> Removed,
        IReadOnlyList<string> Changed,
        IReadOnlyList<ReviewRenameView> Renamed,
        IReadOnlyDictionary<string, IReadOnlyList<ReviewFieldChangeView>> FieldChanges)
    {
        public int TotalChanges =>
            this.Added.Count + this.Removed.Count + this.Changed.Count + this.Renamed.Count;
    }

    public sealed record ReviewRenameView(string From, string To, bool AlsoChanged);

    // ###########################################################################################
    // What a decision produced. Revision is empty for anything but an approval - only publishing
    // moves a system's revision.
    // ###########################################################################################
    public sealed record ReviewDecisionResult(string State, string Revision);

    // ###########################################################################################
    // One field that moved, as the reviewer reads it.
    //
    // An empty Before means the field was BLANK and now has a value; an empty After means it was
    // CLEARED. Both are shown rather than rendered as nothing, because "(blank)" tells a reviewer
    // something and an empty cell tells them the app is broken.
    // ###########################################################################################
    public sealed record ReviewFieldChangeView(string Field, string Before, string After);

    // ###########################################################################################
    // A validation finding, as the reviewer reads it.
    //
    // IsError rather than a severity enum: there are exactly two values and the window's only
    // question is whether to colour it as a problem. An enum mirrored across the wire would be a
    // third place the vocabulary lives.
    // ###########################################################################################
    public sealed record ReviewFindingView(string Code, string Subject, string Message, bool IsError);

    // ###########################################################################################
    // One queue row as the app holds it.
    //
    // CreatedUtc is NULLABLE rather than defaulted, because "waiting since" is shown to the
    // reviewer and a missing timestamp rendered as 1970 would read as a submission that has been
    // waiting fifty years.
    // ###########################################################################################
    public sealed record ReviewQueueRow(
        long Id,
        string SystemId,
        string State,
        string Summary,
        string ContactEmail,
        DateTimeOffset? CreatedUtc);
}
