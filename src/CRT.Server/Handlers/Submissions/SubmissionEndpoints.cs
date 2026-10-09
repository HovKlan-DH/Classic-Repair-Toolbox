using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Maps the submission flows onto HTTP. A RIM AND NOTHING ELSE, the same rule AccountEndpoints
    // follows: read the request into plain values, call one SubmissionFlows method, turn the
    // verdict into a status code.
    //
    // *** CONTRIBUTING REQUIRES NO ACCOUNT. *** Anyone may submit; they give an email address so
    // they can be told whether their work was accepted, and nothing else. See
    // NewContributeStrategy.md, "CONTRIBUTING NEEDS NO ACCOUNT" - in short, the threat model never
    // claimed accounts defended against bad submissions, and a sign-up wall before a hobbyist can
    // fix a typo is how a contribution does not happen.
    //
    // OWNERSHIP IS PROVED BY A CAPABILITY TOKEN, not by identity. Creating a submission returns a
    // random 256-bit token; every later call for that submission presents it in the
    // X-Submission-Token header. Without it, the submission id - a small consecutive integer -
    // would be all anyone needed to upload into a stranger's work.
    //
    // An Authorization header is still HONOURED when present, so a signed-in maintainer's
    // submission is attributed to their account. Its absence is the ordinary case.
    //
    // THE STATUS CODES:
    //   201 Created  - submission accepted, here is the token and what to upload.
    //   400          - the manifest is wrong; the findings say how.
    //   404          - no such submission, OR the token does not match. The two are deliberately
    //                  indistinguishable: distinguishing them would confirm which ids exist and
    //                  let a caller enumerate submissions by walking integers.
    //   409          - the submission is no longer accepting uploads.
    //   416          - the offset does not match; the body says where to resume.
    //   429          - this address has submitted its daily allowance (Retry-After says when).
    //   507          - the server's disk is below its free-space reserve; nothing is wrong
    //                  with the submission. (Both added by the security review, 2026-09-25, and
    //                  both carry { errors: [...] } so CRT shows a sentence, not a number.)
    // ###########################################################################################
    public static class SubmissionEndpoints
    {
        public static void MapSubmissionEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder submissions = app.MapGroup("/api/submissions");

            submissions.MapPost("/", SubmissionEndpoints.CreateAsync)
                .WithBodyLimit(RequestBodyLimits.ManifestBytes);
            submissions.MapPut("/{submissionId:long}/blobs/{hash}", SubmissionEndpoints.UploadAsync)
                .WithBodyLimit(RequestBodyLimits.BlobChunkBytes);
            submissions.MapGet("/{submissionId:long}/blobs/{hash}", SubmissionEndpoints.GetUploadStateAsync);
            submissions.MapPost("/{submissionId:long}/finalise", SubmissionEndpoints.FinaliseAsync);
            submissions.MapGet("/{submissionId:long}", SubmissionEndpoints.GetSubmissionAsync);

            // The contributor discarded their own draft of this board in CRT (owner request,
            // 2026-09-28) - see CRT.Data's DraftDiscardContract. No body.
            submissions.MapPost("/{submissionId:long}/" + DraftDiscardContract.RouteSegment, SubmissionEndpoints.DraftDiscardedAsync);
        }

        // ###########################################################################################
        // POST /api/submissions/{id}/draft-discarded - the contributor discarded their draft.
        //
        // Proved by the capability token, like the status check; 404 for a submission that does not
        // exist and for a token that does not match, indistinguishably. 204 whether this notice was
        // the first or a repeat - CRT only needs to know it arrived.
        // ###########################################################################################
        private static async Task<IResult> DraftDiscardedAsync(
            long submissionId,
            HttpContext context,
            ISubmissionStore store,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            SubmissionRecord? submission = await store.FindAsync(submissionId, cancellationToken);

            if (!SubmissionEndpoints.HoldsToken(submission, context))
                return Results.NotFound();

            await DraftDiscardFlow.RecordAsync(submission!, store, accounts, DateTimeOffset.UtcNow, cancellationToken);

            return Results.NoContent();
        }

        // ###########################################################################################
        // POST /api/submissions
        //
        // Steps 1 and 2 of the transport in one round trip: the manifest goes up, the list of
        // hashes the server lacks comes back.
        // ###########################################################################################
        private static async Task<IResult> CreateAsync(
            SubmissionManifest manifest,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore store,
            BlobStore blobs,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            // NO AUTHENTICATION REQUIRED - see the class header. An Authorization header is
            // honoured when present, so a signed-in maintainer's submission is attributed to their
            // account, but its absence is the ordinary case and not an error.
            AccountRecord? account = await SubmissionEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            // A maintainer or administrator is exempt from the per-address submission limit - trusted
            // by the database rows, not by anything the request says. See SubmissionRateLimitPolicy.
            // "Maintainer" means in at least one board's pool (Phase 6 roles), read here per request
            // like everywhere else.
            Submitter submitter = account is null
                ? Submitter.Anonymous(manifest.ContactEmail, SubmissionEndpoints.ClientAddress(context))
                : Submitter.SignedIn(
                    account.Id,
                    SubmissionEndpoints.ClientAddress(context),
                    isTrusted: ReviewAuthority.CanReviewAnything(new ReviewAccess(
                        account, await accounts.GetReviewedBoardIdsAsync(account.Id, cancellationToken))));

            // Where this board's files would land. Used ONLY to resolve and containment-check the
            // submitted paths - nothing is written here at submission time, because nothing is
            // published until a maintainer promotes it by hand.
            string containmentRoot = SubmissionEndpoints.ResolveContainmentRoot(options);

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                manifest, submitter, containmentRoot, store, blobs, DateTimeOffset.UtcNow, cancellationToken,
                PublishedTreeProbe.For(options.DataTreeRoot));

            // The refusals all carry { errors: [...] } - the one shape CRT reads findings from - so
            // a rate limit or a full disk reaches the contributor as a sentence, not a status code.
            if (outcome.IsRateLimited)
            {
                context.Response.Headers.RetryAfter =
                    ((int)Math.Ceiling(outcome.RetryAfter.TotalSeconds))
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);

                return Results.Json(new { errors = outcome.Findings }, statusCode: StatusCodes.Status429TooManyRequests);
            }

            if (outcome.IsNoRoom)
                return Results.Json(new { errors = outcome.Findings }, statusCode: StatusCodes.Status507InsufficientStorage);

            if (!outcome.IsAccepted)
                return Results.BadRequest(new { errors = outcome.Findings });

            return Results.Created(
                $"/api/submissions/{outcome.Negotiation!.SubmissionId}",
                outcome.Negotiation);
        }

        // ###########################################################################################
        // PUT /api/submissions/{id}/blobs/{hash}?offset=N
        //
        // The body is the raw bytes, not a form or a JSON wrapper: a 200 MB file base64-encoded
        // inside JSON is a third larger and has to be buffered to be parsed, which defeats
        // streaming entirely.
        // ###########################################################################################
        private static async Task<IResult> UploadAsync(
            long submissionId,
            string hash,
            long? offset,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore store,
            BlobStore blobs,
            CancellationToken cancellationToken)
        {
            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                submissionId,
                hash,
                offset ?? 0,
                context.Request.Body,
                SubmissionEndpoints.UploadToken(context),
                store,
                blobs,
                DateTimeOffset.UtcNow,
                cancellationToken);

            return outcome.Status switch
            {
                BlobUploadStatus.Completed =>
                    Results.Ok(new BlobUploadAnswer(outcome.ResumeFrom, true)),

                BlobUploadStatus.Partial =>
                    Results.Ok(new BlobUploadAnswer(outcome.ResumeFrom, false)),

                // 404 for both "no such submission" and "not yours" - see the class header.
                BlobUploadStatus.NotFound => Results.NotFound(),

                BlobUploadStatus.Expired or BlobUploadStatus.WrongState =>
                    Results.Json(new { message = outcome.Error }, statusCode: StatusCodes.Status409Conflict),

                BlobUploadStatus.UnexpectedHash =>
                    Results.BadRequest(new { message = outcome.Error }),

                // 416 Range Not Satisfiable, carrying where to actually resume from.
                BlobUploadStatus.ChunkRejected =>
                    Results.Json(
                        new BlobChunkRejectedAnswer(outcome.Error, outcome.ResumeFrom),
                        statusCode: StatusCodes.Status416RangeNotSatisfiable),

                BlobUploadStatus.HashMismatch =>
                    Results.BadRequest(new { message = outcome.Error }),

                // The disk is below its reserve. Not the contributor's fault and not retried by
                // CRT's upload loop (507 is not a transient status there), so they are told plainly.
                BlobUploadStatus.NoRoom =>
                    Results.Json(new { message = outcome.Error }, statusCode: StatusCodes.Status507InsufficientStorage),

                _ => Results.BadRequest(new { message = "The upload could not be accepted." })
            };
        }

        // ###########################################################################################
        // GET /api/submissions/{id}/blobs/{hash}
        //
        // How much of this blob has arrived. A client that lost its connection asks this before
        // resuming, rather than guessing or restarting - which is what makes resumption work
        // across a client restart, not just a retry.
        // ###########################################################################################
        private static async Task<IResult> GetUploadStateAsync(
            long submissionId,
            string hash,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore store,
            BlobStore blobs,
            CancellationToken cancellationToken)
        {
            UploadState? state = await SubmissionFlows.GetUploadStateAsync(
                submissionId, hash, SubmissionEndpoints.UploadToken(context), store, blobs, cancellationToken);

            if (state is null)
                return Results.NotFound();

            return Results.Ok(new BlobUploadAnswer(state.Uploaded, state.Complete));
        }

        // ###########################################################################################
        // POST /api/submissions/{id}/finalise
        //
        // 200 with IsAccepted true means QUEUED FOR REVIEW - never published.
        // ###########################################################################################
        private static async Task<IResult> FinaliseAsync(
            long submissionId,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore store,
            BlobStore blobs,
            SubmissionNotifier notifier,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                submissionId, SubmissionEndpoints.UploadToken(context), store, blobs,
                DateTimeOffset.UtcNow, cancellationToken);

            if (result.Findings.Any(finding => finding.Code == "submission.not_found"))
                return Results.NotFound();

            // ###########################################################################################
            // *** THE MAINTAINERS ARE TOLD, AFTER THE SUBMISSION IS DURABLY QUEUED (Phase 6 task 11,
            // 2026-09-25). *** Nothing here may fail the request: the contributor's upload is
            // complete and recorded, and a mail problem is the server's to log, not theirs to
            // retry. Who is told is SubmissionRouting's decision - the board's maintainers, or the
            // administrators when there are none or the submission changes shared files.
            // ###########################################################################################
            if (result.IsAccepted)
            {
                try
                {
                    SubmissionRecord? record = await store.FindAsync(submissionId, cancellationToken);

                    if (record is not null)
                    {
                        IReadOnlyList<MailRecipient> recipients = await SubmissionRouting.RecipientsForAsync(
                            record, accounts, cancellationToken);

                        // A completely new board or an update (owner request, 2026-10-03) - new when
                        // the BETA tree holds nothing of it, the review queue's own test
                        // (ReviewQueueFlow), so the mail and the queue's "New board" heading agree.
                        bool isNewBoard = !PublishedBoardLocator.LocateBoard(options.DataTreeRoot, record.BoardId).Exists;

                        await notifier.NotifyMaintainersAsync(
                            recipients, record.BoardId, record.Id, record.Summary, isNewBoard, cancellationToken);
                    }
                }
                catch (Exception ex)
                {
                    context.RequestServices
                        .GetRequiredService<ILogger<SubmissionNotifier>>()
                        .LogWarning(ex, "Submission {SubmissionId} was queued but its maintainers could not be told.", submissionId);
                }
            }

            return Results.Ok(result);
        }

        // ###########################################################################################
        // GET /api/submissions/{id}
        //
        // The state of ONE submission, for the "my submissions" view (Phase 4 task 6).
        //
        // ONE AT A TIME, BY TOKEN - there is deliberately no "list everything I have submitted".
        // With no account there is no identity to list against, and using the contact address as
        // one would turn it into exactly the identity the whole design avoids: anyone who knew
        // somebody's address could enumerate their contributions.
        //
        // So the CLIENT keeps the record. CRT stores the id and token of each submission it made
        // and asks about them individually, which needs no identity at all and means the server
        // holds no queryable link between a person and their work.
        // ###########################################################################################
        private static async Task<IResult> GetSubmissionAsync(
            long submissionId,
            HttpContext context,
            ISubmissionStore store,
            CancellationToken cancellationToken)
        {
            SubmissionRecord? submission = await store.FindAsync(submissionId, cancellationToken);

            if (!SubmissionEndpoints.HoldsToken(submission, context))
                return Results.NotFound();

            IReadOnlyList<ValidationFinding> findings =
                await store.GetFindingsAsync(submissionId, cancellationToken);

            // "merged" is in BETA; once the board has been published to production since, the
            // contributor is told "published" - see ProductionPromotionRules.ContributorFacingState.
            BoardRecord? board = submission!.State == SubmissionState.Merged
                ? await store.FindBoardAsync(submission.BoardId, cancellationToken)
                : null;

            // Whether a BETA rollback returned it (migration 0015) - only a pending row can read so.
            DateTimeOffset? returnedUtc = submission.State == SubmissionState.Pending &&
                (await store.GetBetaReturnsAsync([submission.Id], cancellationToken)).TryGetValue(submission.Id, out DateTimeOffset returned)
                    ? returned
                    : null;

            return Results.Ok(SubmissionEndpoints.BuildStatus(
                submission,
                ProductionPromotionRules.ContributorFacingState(
                    submission.State, submission.DecidedUtc, board?.ProductionPublishedUtc, returnedUtc),
                amendedByMaintainer: await store.GetLatestAmendmentAsync(submissionId, cancellationToken) is not null,
                findings));
        }

        // ###########################################################################################
        // The status answer - CRT.Data's SubmissionStatus, which CRT reads back as the SAME type
        // (SubmissionClient.GetStatusAsync), so no field can be renamed at one end only (code review,
        // 2026-09-25: it was an anonymous object with hand-typed names, renamed at both ends with
        // nothing to notice if only one had moved). SubmissionStatusWireTests puts this through the
        // server's JSON settings and CRT's.
        //
        // *** MAINTAINERCOMMENT IS THE CONTRIBUTOR'S ONLY FEEDBACK. *** Contributing needs no account,
        // so there is no inbox and no thread - the contact email and this sentence are the whole
        // channel back to the person who did the work. Reserved from the start as "always present,
        // always empty for now", so filling it in later was a server change alone with nothing to
        // update on contributors' machines - which is what happened when the review decisions landed
        // on 2026-09-22.
        // ###########################################################################################
        internal static SubmissionStatus BuildStatus(
            SubmissionRecord submission,
            string contributorFacingState,
            bool amendedByMaintainer,
            IReadOnlyList<ValidationFinding> findings) => new()
        {
            Id = submission.Id,
            BoardId = submission.BoardId,
            State = contributorFacingState,
            Summary = submission.Summary ?? string.Empty,
            CreatedUtc = submission.CreatedUtc,
            DecidedUtc = submission.DecidedUtc,
            MaintainerComment = submission.DecisionComment ?? string.Empty,
            AmendedByMaintainer = amendedByMaintainer,
            Findings = [.. findings]
        };

        // -------------------------------------------------------------------------------------
        // Helpers.
        // -------------------------------------------------------------------------------------

        // ###########################################################################################
        // The capability token proving ownership of a submission in flight.
        //
        // A HEADER RATHER THAN A QUERY PARAMETER, deliberately: query strings are written to access
        // logs by default on most web servers, and this value is what authorises writing to a
        // submission. A header is not logged unless somebody asks for it to be.
        // ###########################################################################################
        private const string UploadTokenHeader = SubmissionFormat.UploadTokenHeader;

        // ###########################################################################################
        // The client's address, recorded against a submission for rate limiting.
        //
        // WITH NO ACCOUNT THIS IS THE ONLY THING TO LIMIT AGAINST, and it is weak: one address can
        // be a household, an office or a whole country behind CGNAT, and an attacker can rotate
        // addresses. That is an acknowledged cost of anonymous submission - the blob size cap and
        // the 24-hour abandoned-upload sweep are what actually bound the damage.
        //
        // UseForwardedHeaders has already run and trusts only loopback proxies (see Program.cs),
        // so this is the real client address rather than Apache's.
        // ###########################################################################################
        internal static string? ClientAddress(HttpContext context)
        {
            return context.Connection.RemoteIpAddress?.ToString();
        }

        private static string? UploadToken(HttpContext context)
        {
            string value = context.Request.Headers[SubmissionEndpoints.UploadTokenHeader].ToString();

            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        // Used by the read-only status endpoint, which does not go through SubmissionFlows. The
        // comparison is fixed-time for the same reason SubmissionFlows.OwnsSubmission's is.
        private static bool HoldsToken(SubmissionRecord? submission, HttpContext context)
        {
            string? presented = SubmissionEndpoints.UploadToken(context);

            if (submission is null || presented is null)
                return false;

            return SecureToken.HashesEqual(submission.UploadTokenHash, SecureToken.Hash(presented));
        }

        private static async Task<AccountRecord?> AuthenticateAsync(
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            string header = context.Request.Headers.Authorization.ToString();

            const string prefix = "Bearer ";

            if (!header.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return null;

            string token = header[prefix.Length..].Trim();

            return token.Length == 0
                ? null
                : await AccountFlows.AuthenticateAsync(token, accounts, DateTimeOffset.UtcNow, cancellationToken);
        }

        // ###########################################################################################
        // Where this board's files would live in the BETA tree.
        //
        // BUILT FROM THE CONFIGURED ROOT PLUS THE MANIFEST'S OWN IDENTITY, and every path in the
        // submission is then containment-checked against it. The identity values are untrusted, so
        // they are passed through SubmissionPathRules like everything else - a manufacturer of
        // "../.." would otherwise relocate the whole board folder.
        //
        // The folder is NOT created here and nothing is written to it: a submission is queued, and
        // publication remains a manual act by the project owner.
        // ###########################################################################################
        // ###########################################################################################
        // The root a submission's file paths are contained to.
        //
        // *** THE DATA ROOT, NOT THE BOARD'S OWN FOLDER (fixed 2026-09-23). *** This used to
        // resolve down to "<root>/Commodore/C64/250407" and hand that over as the containment base,
        // which was wrong twice over:
        //
        //   - a submitted path is ALREADY data-root-relative ("Commodore/C64/250407/Sheet1.png"),
        //     so validating it against the board folder measured it from one level too deep;
        //   - a SHARED file ("Commodore/Shared files/Component images/6526.png") sits outside the
        //     board folder by design, so a submission citing one would have been refused.
        //
        // The publish side made the identical mistake and it is what actually broke: PublishPlan
        // wrote 1,215 files into "250407/Commodore/C64/250407/...", duplicating the whole board
        // inside itself. The two must use the same base or a submission validates against one
        // location and publishes to another, which is precisely how that went unnoticed.
        //
        // Containment is not weakened: every path is still resolved and refused if it escapes.
        // What decides a path BELONGS to this submission is SubmissionValidator's identity check,
        // which is where that question belongs.
        // ###########################################################################################
        private static string ResolveContainmentRoot(ServerOptions options)
        {
            return options.DataTreeRoot ?? string.Empty;
        }
    }
}
