using CRT.Server.Configuration;
using Handlers.DataHandling;
using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The REVIEWER's side of the API (NewContributeStrategy.md Phase 5, task 2) - the queue and
    // one submission's detail. A rim and nothing else, the same rule AccountEndpoints and
    // SubmissionEndpoints follow.
    //
    // *** SEPARATE FROM SubmissionEndpoints BECAUSE THE AUTHORISATION MODEL IS THE OPPOSITE. ***
    // Those endpoints are for CONTRIBUTORS: no account, ownership proved by a capability token,
    // and "not found" and "not yours" deliberately indistinguishable so the id space cannot be
    // walked. These are for REVIEWERS: an account is required, a role is required, and there is
    // no token. Mixing the two in one file is how a route eventually gets mapped into the wrong
    // group and inherits the wrong rule - the kind of mistake that reads as a one-line diff.
    //
    // *** EVERY ROUTE HERE CHECKS AUTHORITY SERVER-SIDE, ON EVERY REQUEST. *** Phase 6 task 6
    // states it and the reasoning applies already: the desktop app hiding a button is not
    // enforcement, because the app is public source and an attacker calls the API directly.
    // ReviewAuthority answers the question; no route makes its own judgement.
    //
    // THE STATUS CODES:
    //   200 OK       - here is the queue, or the submission.
    //   401          - no usable credentials.
    //   403          - authenticated, but this account may not review. DISTINCT from 401 on
    //                  purpose: a reviewer whose account lacks the role needs to be told that,
    //                  not handed a login prompt that will not help.
    //   404          - no such submission. Unlike the contributor side there is no id-walking
    //                  concern here, because the caller is already a trusted, named account.
    // ###########################################################################################
    public static class ReviewEndpoints
    {
        // A queue longer than this is a sign of a backlog to deal with, not a list to scroll. The
        // store clamps it too - this is the default, not the ceiling.
        public const int DefaultQueueLimit = 100;

        public static void MapReviewEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder review = app.MapGroup("/api/review");

            review.MapGet("/queue", ReviewEndpoints.GetQueueAsync);
            review.MapGet("/submissions/{submissionId:long}", ReviewEndpoints.GetSubmissionAsync);

            // The two asset routes (task 4). They are separated by WHICH SIDE of the comparison
            // they serve, not merged into one route with a "which" parameter, because they take
            // different input and are guarded differently - see ReviewAssetLocator's header. A
            // single route would have to branch internally on untrusted input to pick its own
            // guard, which is how the wrong one eventually gets applied.
            review.MapGet("/submissions/{submissionId:long}/submitted/{hash}", ReviewEndpoints.GetSubmittedAssetAsync);
            review.MapGet("/submissions/{submissionId:long}/published/{**path}", ReviewEndpoints.GetPublishedAssetAsync);

            // ###########################################################################################
            // The review DECISIONS (task 5). All three outcomes.
            //
            // *** APPROVE IS THE ONLY IRREVERSIBLE ONE. *** It publishes: with task 7 struck there
            // is no retained revision, so it overwrites the board in place and the only way back
            // is a correction. ApprovePublishFlow refuses early and often, and every refusal
            // before its final step leaves the tree untouched.
            //
            // Rejecting and returning touch no published data at all, which is why a REVIEWER may
            // reach those two and only an administrator may approve.
            // ###########################################################################################
            review.MapPost("/submissions/{submissionId:long}/approve", ReviewEndpoints.ApproveAsync);
            review.MapPost("/submissions/{submissionId:long}/reject", ReviewEndpoints.RejectAsync);
            review.MapPost("/submissions/{submissionId:long}/request-changes", ReviewEndpoints.RequestChangesAsync);
        }

        // ###########################################################################################
        // POST /api/review/submissions/{id}/approve
        //
        // Publishes the submission into the data tree. **The one irreversible operation in the
        // system** - see ApprovePublishFlow, which owns the whole sequence.
        //
        // This route is a rim and nothing more: it authorises, delegates, and maps the outcome to
        // a status code. It deliberately makes no judgement of its own about whether the publish
        // may proceed, because a second opinion living here is a second place to get it wrong.
        // ###########################################################################################
        private static async Task<IResult> ApproveAsync(
            long submissionId,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            ApprovePublishFlow approvals,
            SubmissionNotifier notifier,
            ServerOptions options,
            ILogger<ApprovePublishFlow> logger,
            CancellationToken cancellationToken)
        {
            (AccountRecord? account, IResult? refusal) =
                await ReviewEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            ApproveOutcome outcome = await approvals.ApproveAsync(
                submissionId,
                account,
                options.DataTreeRoot,
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (outcome.IsPublished)
            {
                // ###########################################################################################
                // *** EVERYTHING FROM HERE ON IS AFTER-THE-FACT, AND NONE OF IT MAY FAIL THE
                // REQUEST (maintainer report, 2026-09-23). ***
                //
                // The board is on disk and the state is Merged. Publishing is the one irreversible
                // operation in the system, so an exception escaping this block is the worst
                // possible answer: the reviewer is told 500, reads it as "the publish failed", and
                // tries again against a tree that has already been overwritten.
                //
                // That is not hypothetical - it is exactly what happened. DataChecksumManifest.Write
                // guarded its WRITE but ran its scan outside the try, a directory walk over ~11,000
                // files threw, and "Approve and publish" answered 500 on a publish that had
                // succeeded.
                //
                // Each step below now contains its own faults, and this outer catch is the
                // belt-and-braces: a step added later cannot reintroduce the same failure just by
                // forgetting. The publish is reported as the success it was, and the log says what
                // went wrong afterwards.
                // ###########################################################################################
                try
                {
                    await ReviewEndpoints.AfterPublishAsync(
                        submissionId, submissions, notifier, options, logger, cancellationToken);
                }
                catch (Exception ex)
                {
                    logger.LogError(
                        ex,
                        "Submission {SubmissionId} WAS published, but the follow-up work after it " +
                        "failed. The data tree is correct; the checksum manifest or the " +
                        "notification e-mail may not be.",
                        submissionId);
                }

                return Results.Ok(new
                {
                    state = SubmissionState.Merged,
                    revision = outcome.Descriptor!.Revision,
                    contentHash = outcome.Descriptor.ContentHash
                });
            }

            if (outcome.IsNotFound)
                return Results.NotFound();

            // 403 for "your account may not", 409 for "somebody else already decided", 400 for
            // everything else. They send a reviewer to completely different places, which is the
            // whole reason ApproveOutcome distinguishes them rather than returning a bare bool.
            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

            return outcome.IsConflict
                ? Results.Conflict(new { error = outcome.Error })
                : Results.BadRequest(new { error = outcome.Error });
        }

        // ###########################################################################################
        // The work that follows a SUCCESSFUL publish: refresh the checksum manifest, then tell the
        // contributor.
        //
        // Split out of ApproveAsync so the "none of this may fail the request" guard has one thing
        // to wrap and cannot accidentally cover the publish itself. Runs synchronously to
        // completion - the reviewer's 200 waits for it - because a manifest that lags the tree is
        // the very bug this exists to fix, and fire-and-forget would reintroduce a window where a
        // client syncs against a stale one. AWAITED, never blocked on: this is an ASP.NET request
        // thread, and GetAwaiter().GetResult() here is how a server deadlocks under load.
        // ###########################################################################################
        private static async Task AfterPublishAsync(
            long submissionId,
            ISubmissionStore submissions,
            SubmissionNotifier notifier,
            ServerOptions options,
            ILogger logger,
            CancellationToken cancellationToken)
        {
            // ###########################################################################################
            // *** REGENERATE dataChecksums.json, OR THE PUBLISH REACHES NOBODY (maintainer
            // report, 2026-09-23). ***
            //
            // CRT decides what to download by comparing that manifest against what it already
            // holds. The board has just been written into the data tree, but until the manifest
            // advertises the new checksum every client concludes there is nothing to fetch -
            // and says so by completing a sync that does nothing. That is exactly what
            // happened on the first real publish: the file on the server was correct, dated
            // minutes earlier, while the manifest beside it was two weeks old.
            //
            // It sits ONE FOLDER UP from the data root, beside `Data/` rather than inside it,
            // which is why ManifestPath is configured separately and cannot be derived from
            // DataTreeRoot.
            //
            // *** IT MUST NOT FAIL THE PUBLISH. *** The tree is already overwritten and this is
            // the one irreversible operation in the system; an error here would be read as
            // "the publish failed" and retried against data that has already changed. A failed
            // regeneration leaves the PREVIOUS manifest in place, so clients stay on what they
            // have until it is rebuilt - stale, but never inconsistent.
            // ###########################################################################################
            int written = DataChecksumManifest.Write(
                options.DataTreeRoot ?? string.Empty,
                options.PublicDataBaseUrl ?? string.Empty,
                options.ManifestPath ?? string.Empty);

            if (written < 0)
            {
                logger.LogWarning(
                    "Submission {SubmissionId} was published but the checksum manifest at " +
                    "[{ManifestPath}] could not be regenerated - clients will not see the " +
                    "change until it is rebuilt.",
                    submissionId,
                    options.ManifestPath);
            }
            else
            {
                logger.LogInformation(
                    "Checksum manifest regenerated after publishing submission " +
                    "{SubmissionId}: {EntryCount} files.",
                    submissionId,
                    written);
            }

            // ###########################################################################################
            // *** THE MAIL COMES LAST, AFTER THE ONE IRREVERSIBLE OPERATION IN THE SYSTEM. ***
            //
            // The board is already written and the state already Merged. Nothing here may throw
            // back to the reviewer - they would read the error as "the publish failed" and try
            // again, against a tree that has already been overwritten. SubmissionNotifier
            // swallows and logs; see its header.
            //
            // The record is re-read rather than ApproveOutcome being widened to carry the
            // address: that type describes the PUBLISH, and the contributor's contact details
            // are nothing to do with whether a publish succeeded. A read here costs one query
            // on the rarest operation the service performs.
            //
            // An approval usually carries no reviewer comment, which is exactly why the
            // published template takes one as optional.
            // ###########################################################################################
            SubmissionRecord? published =
                await submissions.FindAsync(submissionId, cancellationToken);

            if (published is not null)
            {
                await notifier.NotifyDecisionAsync(
                    published.ContactEmail,
                    published.SystemId,
                    SubmissionState.Merged,
                    published.DecisionComment,
                    cancellationToken);
            }
        }

        // ###########################################################################################
        // POST /api/review/submissions/{id}/reject
        //
        // Ends a submission with a reason. Recoverable in the sense that matters: the contributor
        // still holds their draft locally, which is the whole point of the local-first design.
        // ###########################################################################################
        private static Task<IResult> RejectAsync(
            long submissionId,
            ReviewDecisionRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            SubmissionNotifier notifier,
            CancellationToken cancellationToken)
        {
            return ReviewEndpoints.DecideAsync(
                submissionId,
                request,
                context,
                accounts,
                submissions,
                notifier,
                SubmissionState.Rejected,
                ReviewDecisionRules.CanReject,
                cancellationToken);
        }

        // ###########################################################################################
        // POST /api/review/submissions/{id}/request-changes
        //
        // Returns a submission to its contributor with a comment. **The strategy says explicitly
        // not to skip this**: most imperfect contributions are fixable by their author in a
        // minute, and a reject that could have been a conversation costs a contributor.
        // ###########################################################################################
        private static Task<IResult> RequestChangesAsync(
            long submissionId,
            ReviewDecisionRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            SubmissionNotifier notifier,
            CancellationToken cancellationToken)
        {
            return ReviewEndpoints.DecideAsync(
                submissionId,
                request,
                context,
                accounts,
                submissions,
                notifier,
                SubmissionState.ChangesRequested,
                ReviewDecisionRules.CanRequestChanges,
                cancellationToken);
        }

        // ###########################################################################################
        // One decision, whichever it is.
        //
        // The two outcomes differ only in the state they write and which rule answers "may this
        // account do it", so they share everything else rather than carrying two copies of the
        // authorise-check-record sequence. The RULE is passed in rather than an enum branched on
        // here: a caller cannot then reach the wrong rule for the outcome it named.
        //
        // *** THE STATE IS RE-READ AND RE-CHECKED SERVER-SIDE, EVERY TIME. *** The app disables a
        // button on a submission it believes is still pending, and that belief can be seconds out
        // of date - another reviewer may have decided it in the meantime. The check here is the
        // one that counts, and it is why ReviewDecisionRules takes the state rather than trusting
        // the caller to have looked.
        // ###########################################################################################
        private static async Task<IResult> DecideAsync(
            long submissionId,
            ReviewDecisionRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            SubmissionNotifier notifier,
            string newState,
            DecisionRule rule,
            CancellationToken cancellationToken)
        {
            (AccountRecord? account, IResult? refusal) =
                await ReviewEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SubmissionRecord? record = await submissions.FindAsync(submissionId, cancellationToken);

            if (record is null)
                return Results.NotFound();

            if (!rule(account, record.State, out string why))
            {
                // 409, not 403: the account may well be allowed to review, and what is wrong is
                // the SUBMISSION's state - somebody else decided it first. A 403 would send the
                // reviewer looking at their own permissions for a conflict that is about timing.
                return ReviewAuthority.CanReview(account)
                    ? Results.Conflict(new { error = why })
                    : Results.Json(new { error = why }, statusCode: StatusCodes.Status403Forbidden);
            }

            // *** THE REASON IS REQUIRED FOR BOTH OF THESE OUTCOMES. *** Contributors have no
            // account, so the contact email and this message are the entire feedback channel. A
            // rejection with no reason is indistinguishable from being ignored.
            if (!ReviewDecisionRules.IsUsableReason(request?.Comment, out string reasonProblem))
                return Results.BadRequest(new { error = reasonProblem });

            await submissions.SetDecisionAsync(
                submissionId,
                newState,
                account!.Id,
                request!.Comment!.Trim(),
                DateTimeOffset.UtcNow,
                cancellationToken);

            // ###########################################################################################
            // *** AFTER the decision is recorded, and it cannot fail this request. ***
            //
            // The reviewer's verdict is already durable at this point, so a mail problem must not
            // surface as an error - they would try the decision again and find it already made.
            // SubmissionNotifier swallows and logs; see its header.
            //
            // For these two outcomes the comment is REQUIRED (IsUsableReason above refuses the
            // request without one), so the mail always carries a real explanation rather than
            // "your contribution was rejected" and nothing else.
            // ###########################################################################################
            await notifier.NotifyDecisionAsync(
                record.ContactEmail,
                record.SystemId,
                newState,
                request.Comment!.Trim(),
                cancellationToken);

            return Results.Ok(new { state = newState });
        }

        // The shape of "may this account do this to a submission in this state".
        private delegate bool DecisionRule(AccountRecord? account, string? state, out string reason);

        // ###########################################################################################
        // What a reviewer sends with a decision.
        //
        // A record with one field rather than a bare string, because a decision will grow more to
        // say - Phase 6's notification work needs to know whether to email, and an approval will
        // carry the revision. Widening a record is a field; widening a bare string is a breaking
        // change to every caller.
        // ###########################################################################################
        public sealed record ReviewDecisionRequest(string? Comment);

        // ###########################################################################################
        // GET /api/review/submissions/{id}/submitted/{hash}
        //
        // The bytes a contributor UPLOADED, so the review app can draw the "after" side - task 4's
        // images side by side, the moved highlight on its schematic, a scope baseline plotted.
        //
        // *** THE HASH MUST BE ONE THIS SUBMISSION REFERENCES. *** The blob store is shared and
        // content-addressed, so without that scope check a reviewer holding any hash could read
        // any upload in the store. ReviewAssetLocator decides; this route only obeys.
        // ###########################################################################################
        private static async Task<IResult> GetSubmittedAssetAsync(
            long submissionId,
            string hash,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            BlobStore blobs,
            CancellationToken cancellationToken)
        {
            (AccountRecord? _, IResult? refusal) =
                await ReviewEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SubmissionManifest? manifest =
                await submissions.LoadPayloadAsync(submissionId, cancellationToken);

            if (!ReviewAssetLocator.IsSubmittedBlobAllowed(manifest, hash))
                return Results.NotFound();

            if (!blobs.TryGetReadablePath(hash, out string path))
                return Results.NotFound();

            // The NAME is taken from the manifest entry, not from the request, so the content type
            // is decided by what the submission says this file is rather than by anything the
            // caller can vary independently of the bytes.
            string? name = manifest!.Files
                .FirstOrDefault(file => string.Equals(file.Sha256, hash, StringComparison.Ordinal))
                ?.Path;

            return ReviewEndpoints.FileResult(path, name);
        }

        // ###########################################################################################
        // GET /api/review/submissions/{id}/published/{path}
        //
        // The bytes CURRENTLY PUBLISHED for this system, so the review app can draw the "before"
        // side. Without it every comparison is one-sided: a reviewer can see the new image but not
        // what it replaces, which is the whole question for a replaced schematic.
        //
        // *** THE PATH IS UNTRUSTED AND REACHES THE DATA TREE. *** This is the file-disclosure
        // boundary; containment is ReviewAssetLocator's job and is proved by tests that point at
        // real files outside the tree.
        //
        // `{**path}` is a catch-all so a path with slashes arrives whole. It is NOT url-decoded
        // segment by segment into anything that gets recombined - the locator takes the single
        // string and resolves it once.
        // ###########################################################################################
        private static async Task<IResult> GetPublishedAssetAsync(
            long submissionId,
            string path,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (AccountRecord? _, IResult? refusal) =
                await ReviewEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            // Scoped to the submission being reviewed rather than taking a system id directly:
            // the reviewer is looking at a submission, and deriving the system from it means the
            // route cannot be pointed at a system the caller simply named.
            SubmissionManifest? manifest =
                await submissions.LoadPayloadAsync(submissionId, cancellationToken);

            if (!ReviewAssetLocator.TryLocatePublishedFile(
                    options.DataTreeRoot, manifest, path, out string resolved))
            {
                // One 404 for unsafe, absent and out-of-scope alike. Distinguishing them would let
                // the tree be probed for what exists.
                return Results.NotFound();
            }

            return ReviewEndpoints.FileResult(resolved, path);
        }

        // ###########################################################################################
        // Sends a located file.
        //
        // *** THE CONTENT TYPE COMES FROM AN ALLOWLIST, NEVER FROM THE REQUEST. *** These bytes are
        // contributor-supplied and are served from the server's own origin, so a file served as
        // text/html would run script there against a signed-in reviewer. See ReviewAssetLocator.
        //
        // `enableRangeProcessing` because a scope baseline or a full board scan is megabytes and a
        // reviewer may scrub through one; the framework answers the range, this file does not
        // reimplement it.
        // ###########################################################################################
        private static IResult FileResult(string resolvedPath, string? nameForType)
        {
            return Results.File(
                resolvedPath,
                ReviewAssetLocator.ContentTypeFor(nameForType),
                enableRangeProcessing: true);
        }

        // ###########################################################################################
        // GET /api/review/queue
        //
        // Everything waiting for a decision, oldest first - see ISubmissionStore.GetQueueAsync for
        // why that ordering rather than newest-first.
        // ###########################################################################################
        private static async Task<IResult> GetQueueAsync(
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            CancellationToken cancellationToken)
        {
            (AccountRecord? account, IResult? refusal) =
                await ReviewEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            IReadOnlyList<SubmissionRecord> queue =
                await submissions.GetQueueAsync(ReviewEndpoints.DefaultQueueLimit, cancellationToken);

            return Results.Ok(new
            {
                // Told plainly rather than inferred from the presence of an action, so the app can
                // say WHY a publish button is absent instead of silently not drawing one.
                canPublish = ReviewAuthority.CanPublish(account),
                count = queue.Count,
                submissions = queue.Select(ReviewEndpoints.ToQueueRow).ToList()
            });
        }

        // ###########################################################################################
        // GET /api/review/submissions/{id}
        //
        // One submission, with the manifest and the findings a reviewer needs to decide. The
        // manifest is the ROWS - what the board would become - which is what ReviewSummary
        // compares against the published board.
        // ###########################################################################################
        private static async Task<IResult> GetSubmissionAsync(
            long submissionId,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            PublishedBoardReader publishedBoards,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (AccountRecord? account, IResult? refusal) =
                await ReviewEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SubmissionRecord? record = await submissions.FindAsync(submissionId, cancellationToken);

            if (record is null)
                return Results.NotFound();

            SubmissionManifest? manifest =
                await submissions.LoadPayloadAsync(submissionId, cancellationToken);

            ReviewComparison comparison = await ReviewEndpoints.SummariseAsync(
                manifest,
                publishedBoards,
                options,
                cancellationToken);

            return Results.Ok(new
            {
                canPublish = ReviewAuthority.CanPublish(account),
                submission = ReviewEndpoints.ToQueueRow(record),
                manifest,
                findings = await submissions.GetFindingsAsync(submissionId, cancellationToken),
                changes = comparison.Changes,

                // The files the PUBLISHED board references - what the review app compares the
                // submission's own files against. See PublishedFilePaths for why this cannot be
                // derived client-side from `changes`.
                publishedFiles = comparison.PublishedFiles,

                // The SHA-256 of each published file the submission also carries, so the review
                // app can drop the pairs that are byte-identical instead of showing every file on
                // both sides as replaced. See PublishedFileHashes.
                publishedHashes = comparison.PublishedHashes,

                // Which image each schematic is drawn from, so a moved highlight lands on the
                // right board. See SchematicImageFiles.
                schematicImages = comparison.SchematicImages
            });
        }

        // ###########################################################################################
        // The change summary a reviewer opens on (task 3), computed HERE rather than in the app.
        //
        // *** THE SERVER IS THE ONLY PLACE BOTH HALVES EXIST. *** The submitted board travels in
        // the manifest; the published one is a workbook in the data tree that only the server can
        // read. Sending the published board to the client so it could do the comparison would
        // mean shipping a multi-megabyte workbook per submission opened, to compute something that
        // is a few hundred bytes once computed.
        //
        // A MISSING MANIFEST YIELDS NULL rather than an empty summary. A submission whose payload
        // could not be loaded has nothing to compare, and reporting "no changes" for it would
        // invite a reviewer to approve something they have not seen. The client shows the
        // findings instead, which is where the reason will be.
        // ###########################################################################################
        private static async Task<ReviewComparison> SummariseAsync(
            SubmissionManifest? manifest,
            PublishedBoardReader publishedBoards,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string dataTreeRoot = options.DataTreeRoot ?? string.Empty;

            if (manifest is null)
            {
                return new ReviewComparison(
                    null,
                    [],
                    new Dictionary<string, string>(),
                    new Dictionary<string, string>());
            }

            BoardData? published = await publishedBoards
                .TryReadAsync(dataTreeRoot, manifest, cancellationToken)
                .ConfigureAwait(false);

            IReadOnlyList<string> publishedFiles = ReviewEndpoints.PublishedFilePaths(published);

            // Hashed for the files on BOTH sides only - see PublishedFileHashes for why they are
            // hashed here rather than read from dataChecksums.json.
            IReadOnlyDictionary<string, string> publishedHashes = await PublishedFileHashes
                .ComputeAsync(
                    dataTreeRoot,
                    PublishedFileHashes.PathsOnBothSides(publishedFiles, manifest),
                    cancellationToken)
                .ConfigureAwait(false);

            // *** THE REVISION DATE IS NOW COMPARED PROPERLY. *** It used to be a real gap:
            // SubmissionRows carried no revision date, so the published board's was used for BOTH
            // sides and a revision-date change was invisible to the reviewer. The field was added
            // 2026-09-22, and PublishMerge resolves it the same way the publish itself will -
            // submitted if present, otherwise the published one, which is what an older client
            // that omits the field must produce.
            //
            // Using the manifest's BaseRevision here would still be wrong, and worse than the old
            // silence: that is the date the contributor STARTED from, so it would report a change
            // backwards on every submission built against an older revision.
            BoardData submitted = PublishMerge.Build(manifest, published);

            // *** THE CALIBRATIONS ARE COMPARED TOO, and they cannot ride inside either board. ***
            // BoardData has no calibration section; they live in the JSON sidecar. Omitting them
            // here would let a contributor's calibration change reach the published tree with no
            // reviewer having seen it - which is the one thing this screen exists to prevent.
            //
            // The PUBLISHED side comes from the sidecar beside the published workbook; a system
            // with none yields an empty list, which correctly reports every submitted calibration
            // as an addition.
            return new ReviewComparison(
                ReviewSummary.Compare(
                    published,
                    submitted,
                    manifest.Renames,
                    publishedCalibrations: ReviewEndpoints.PublishedCalibrations(dataTreeRoot, manifest, published),
                    submittedCalibrations: PublishMerge.CalibrationsOf(manifest)),
                publishedFiles,
                publishedHashes,
                ReviewEndpoints.SchematicImageFiles(submitted, published));
        }

        // ###########################################################################################
        // Which IMAGE FILE each schematic is drawn from, so the review app can put a moved
        // highlight back on the board it belongs to (task 4).
        //
        // *** A HIGHLIGHT NAMES A SCHEMATIC, NOT A FILE. *** Its natural key is
        // SchematicName|BoardLabel, and the picture lives on a BoardSchematicEntry. Without this
        // mapping the review app knows a rectangle moved on "Sheet 1" and has no way to find the
        // image of Sheet 1 to draw it on.
        //
        // *** THE SUBMITTED BOARD WINS, and that ordering is the point. *** A submission may ADD a
        // schematic, or repoint an existing one at a new image, and in both cases the reviewer must
        // see the board as the submission proposes it. The published board fills in only what the
        // submission does not mention, so a highlight on an untouched schematic still draws.
        // ###########################################################################################
        private static IReadOnlyDictionary<string, string> SchematicImageFiles(
            BoardData submitted,
            BoardData? published)
        {
            var files = new Dictionary<string, string>(StringComparer.Ordinal);

            // Published first, then submitted OVERWRITES - so the submission's own view wins.
            foreach (BoardSchematicEntry entry in
                (published?.Schematics ?? []).Concat(submitted.Schematics))
            {
                if (string.IsNullOrWhiteSpace(entry.SchematicName) ||
                    string.IsNullOrWhiteSpace(entry.SchematicImageFile))
                {
                    continue;
                }

                files[entry.SchematicName] = entry.SchematicImageFile;
            }

            return files;
        }

        // ###########################################################################################
        // Every FILE the published board references, so the review app can tell an added image
        // from a replaced one and spot a deletion (task 4).
        //
        // *** THE CLIENT CANNOT DERIVE THIS FROM THE CHANGE SUMMARY, and the first version of the
        // review window tried to. *** A summary's row keys are NATURAL KEYS - for an image that is
        // BoardLabel|Region|Pin|Name - not file paths. Comparing those against the manifest's file
        // paths matches nothing, so every submitted image would have been reported as an addition
        // and no deletion would ever have appeared. Silent in both directions: the panel draws
        // perfectly and describes the wrong thing.
        //
        // So the server states it outright. It already holds the published board, and this is the
        // only place both halves exist - the same reasoning the change summary itself is computed
        // here rather than in the app.
        //
        // A new system yields an EMPTY list, which correctly makes every image an addition.
        // ###########################################################################################
        // ###########################################################################################
        // The calibrations the PUBLISHED board carries, read from its JSON sidecar.
        //
        // *** NOT FROM BoardData, WHICH HAS NO CALIBRATION SECTION. *** They live only in the
        // sidecar, under their own root, so this reads the file directly - the same file the
        // publish writes.
        //
        // A system with no sidecar, or no calibrations in it, yields an EMPTY list, which
        // correctly reports every submitted calibration as an addition. Never throws: an
        // unreadable sidecar must not make a submission impossible to open, and the reviewer sees
        // the findings instead.
        // ###########################################################################################
        private static IReadOnlyList<KiCadCalibrationEntry> PublishedCalibrations(
            string dataTreeRoot,
            SubmissionManifest manifest,

            // The board that has ALREADY been read, rather than reading it again. A second read
            // would re-parse a multi-megabyte workbook for a handful of numbers, and - more to the
            // point - would need its own BoardDataReader cache key, which is precisely the
            // stale-cache trap PublishedBoardReader's header documents.
            BoardData? published)
        {
            PublishedBoardLocation location = PublishedBoardLocator.Locate(dataTreeRoot, manifest);

            if (!location.Exists)
                return [];

            var entries = new List<KiCadCalibrationEntry>();

            // TryLoadKiCadCalibration answers for ONE schematic at a time, so the names have to
            // come from the board - which is the only thing that lists them.
            foreach (BoardSchematicEntry schematic in published?.Schematics ?? [])
            {
                if (string.IsNullOrWhiteSpace(schematic.SchematicName))
                    continue;

                if (BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                        location.WorkbookPath,
                        schematic.SchematicName,
                        out string cadName,
                        out double offsetX,
                        out double offsetY,
                        out double scaleX,
                        out double scaleY,
                        out bool mirrorX,
                        out bool mirrorY))
                {
                    entries.Add(new KiCadCalibrationEntry
                    {
                        SchematicName = schematic.SchematicName,
                        CadName = cadName,
                        OffsetX = offsetX,
                        OffsetY = offsetY,
                        ScaleX = scaleX,
                        ScaleY = scaleY,
                        MirrorX = mirrorX,
                        MirrorY = mirrorY
                    });
                }
            }

            return entries;
        }

        // ###########################################################################################
        // *** THE SAME COLLECTOR THE SUBMISSION SIDE USES, AND THAT IS THE WHOLE POINT
        // (fixed 2026-09-23). ***
        //
        // This used to read ComponentImages ALONE, while SubmissionManifestBuilder.CollectReferencedFiles
        // collects four sources: schematic images, component images, component local files and
        // board local files. The review app compares one list against the other, so every file
        // from the three missing sources was present on the submitted side and absent on the
        // published side - and was reported to the reviewer as ADDED.
        //
        // Reported by the maintainer: a submission that changed one component's short description
        // listed every schematic image as "Added / Not in the published board". Nothing was
        // actually being added; the published list simply did not know those files existed.
        //
        // Calling the shared collector is what stops the two lists drifting again. A file source
        // added to a submission is now automatically known to the published side too, rather than
        // being a second place somebody has to remember to edit.
        //
        // CollectReferencedFiles already de-duplicates (several component-image rows legitimately
        // cite one capture) and returns a sorted list, so no further work is needed here.
        // ###########################################################################################
        private static IReadOnlyList<string> PublishedFilePaths(BoardData? published)
        {
            if (published is null)
                return [];

            return SubmissionManifestBuilder.CollectReferencedFiles(published);
        }

        // ###########################################################################################
        // Both halves of what comparing a submission produces.
        //
        // Changes is NULL when there was no manifest to compare - distinct from a summary with no
        // changes, which the client words differently.
        // ###########################################################################################
        private sealed record ReviewComparison(
            ReviewChangeSummary? Changes,
            IReadOnlyList<string> PublishedFiles,

            // Path -> SHA-256 for the published files the submission also names, so the review
            // app can drop byte-identical pairs. See PublishedFileHashes.
            IReadOnlyDictionary<string, string> PublishedHashes,

            // Schematic name to the image file it is drawn from, so a moved highlight can be put
            // back on its own board. See SchematicImageFiles.
            IReadOnlyDictionary<string, string> SchematicImages);

        // ###########################################################################################
        // Resolves the caller and refuses when they may not review.
        //
        // Returns the refusal rather than throwing, so each route reads as "authorise, then do the
        // work" with no exception-handling middleware in between deciding what a 403 looks like.
        //
        // 401 AND 403 ARE DISTINGUISHED HERE, unlike on the contributor side where 404 covers both
        // "no such thing" and "not yours". The reasoning differs because the risk differs: there,
        // telling the truth lets an anonymous caller enumerate submissions; here the caller is an
        // authenticated named account, and being told "your account cannot do this" is the only
        // way they learn to ask for the role rather than retrying a login that already worked.
        // ###########################################################################################
        private static async Task<(AccountRecord? Account, IResult? Refusal)> AuthoriseAsync(
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            // *** ServerOptions IS RESOLVED HERE RATHER THAN TAKEN AS A PARAMETER. *** Passing it
            // in would mean adding it to the several routes that do not otherwise need it, and a
            // route whose author forgot would silently stop extending sessions - a bug with no
            // symptom until somebody is asked for a password they were promised they would not
            // need. Resolving from the context means every caller gets it by construction.
            //
            // Resolved as a PLAIN SINGLETON, not IOptions<T> - Program.cs binds one instance and
            // registers that object directly. Optional: a null simply disables sliding expiry
            // rather than failing the request.
            ServerOptions? options =
                context.RequestServices.GetService(typeof(ServerOptions)) as ServerOptions;

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                ReviewEndpoints.BearerToken(context),
                accounts,
                DateTimeOffset.UtcNow,
                cancellationToken,
                options);

            if (account is null)
                return (null, Results.Unauthorized());

            if (!ReviewAuthority.CanReview(account))
                return (null, ReviewEndpoints.NotAReviewer());

            return (account, null);
        }

        // ###########################################################################################
        // The 403 for a signed-in account that may not review.
        //
        // *** NOT Results.Forbid(), which cannot work in this service (code review, 2026-09-25). ***
        // Forbid() is a challenge to the AUTHENTICATION middleware: it asks the registered default
        // scheme to write the refusal. This service authenticates its own opaque bearer tokens and
        // registers no scheme at all, so executing Forbid() threw InvalidOperationException and the
        // reviewer got a 500 - never the 403 this header documents and the review app's "this
        // account is not allowed to review" message depends on. ApproveAsync's own refusal already
        // wrote its 403 directly; this now does the same. The sentence matches the review app's.
        // ###########################################################################################
        internal static IResult NotAReviewer() =>
            Results.Json(
                new { error = "This account is not allowed to review submissions." },
                statusCode: StatusCodes.Status403Forbidden);

        // ###########################################################################################
        // One row as a reviewer sees it.
        //
        // *** THE UPLOAD TOKEN HASH IS NOT HERE, AND MUST NEVER BE. *** SubmissionRecord carries
        // it because the store needs it to check ownership; it is the contributor's capability for
        // that submission, and echoing it to any other caller would hand a reviewer the ability to
        // act as the contributor. The contact email is included because a reviewer has to be able
        // to reply to the person - that is the whole channel, since contributors have no account.
        // ###########################################################################################
        private static object ToQueueRow(SubmissionRecord record) => new
        {
            id = record.Id,
            systemId = record.SystemId,
            state = record.State,
            summary = record.Summary,
            contactEmail = record.ContactEmail,
            baseRevision = record.BaseRevision,
            createdUtc = record.CreatedUtc,
            decidedUtc = record.DecidedUtc
        };

        // ###########################################################################################
        // The bearer token, or null.
        //
        // Duplicated from AccountEndpoints' own private copy rather than shared, and that is a
        // deliberate small duplication: making it public would put a parsing helper on the
        // account API's surface for no caller outside this assembly. If a third copy is ever
        // wanted, that is the moment to move it into a shared internal helper rather than now.
        // ###########################################################################################
        private static string? BearerToken(HttpContext context)
        {
            string header = context.Request.Headers.Authorization.ToString();

            const string prefix = "Bearer ";

            if (!header.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return null;

            string token = header[prefix.Length..].Trim();

            return token.Length == 0 ? null : token;
        }
    }
}
