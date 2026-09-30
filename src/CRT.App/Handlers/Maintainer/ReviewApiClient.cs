using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using CRT;
using Handlers.DataHandling;
using Handlers.OnlineHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // Talks to the review API over HTTP (NewContributeStrategy.md Phase 5, task 2).
    //
    // *** AN UNTESTED I/O BOUNDARY, ON PURPOSE. *** CLAUDE.md test rule 6 forbids tests that make
    // network calls, and the abstraction below a boundary is the thing to test rather than the
    // boundary itself. That is why this class is as THIN as it is: the routes are in
    // ReviewApiRoutes and the parsing is in ReviewApiParser, both pure and both covered. What is
    // left here is sending a request and reading a status code, which is HttpClient's job.
    //
    // Everything it adds is about not lying to the caller:
    //
    //   - a FAILURE IS A VALUE, never an exception. This is called from UI event handlers, where
    //     an unhandled task exception is a crash; and a maintainer whose office wifi dropped needs
    //     a sentence, not a stack trace.
    //   - an UNREACHABLE SERVER AND AN EMPTY QUEUE ARE DIFFERENT ANSWERS, carried all the way
    //     through. If they collapse into one, a broken connection reads as "nothing to review"
    //     and a backlog goes unnoticed.
    //   - a 401 and a 403 are DISTINGUISHED, because the server distinguishes them and they mean
    //     different things to the person reading: sign in again, versus your account lacks the
    //     role. Telling somebody to sign in again when the credentials were fine is a loop they
    //     cannot escape.
    // ###########################################################################################
    public sealed class ReviewApiClient : IDisposable
    {
        private readonly HttpClient thisHttp;
        private readonly string thisBaseAddress;

        public ReviewApiClient(string baseAddress, HttpClient? http = null)
        {
            this.thisBaseAddress = baseAddress;

            // The caller may supply its own client (a shared one, or one with a proxy configured).
            //
            // *** A LITTLE LONGER THAN THE WINDOW WAITS (2026-09-28). *** It was 20 seconds, so a
            // dead server did not leave a window looking frozen - but an approval publishing a large
            // board can legitimately take longer than that, and was then reported as "did not
            // answer" while the server carried on. Every request the maintainer waits for now runs
            // under the BusyOverlay, which shows the wait and gives up after WaitLimit.Maximum; this
            // is only the backstop for the checks nobody is watching (the minute queue check).
            this.thisHttp = http ?? new HttpClient { Timeout = WaitLimit.Maximum + TimeSpan.FromSeconds(10) };

            // ###########################################################################################
            // *** EVERY REQUEST NAMES THE CRT THAT SENT IT (2026-09-29). *** "CRT <version>", the
            // User-Agent CRT's own check-in, sync and board views send (OnlineServices.VersionForServer),
            // so the server's log - and the session row, which keeps it - says which build a
            // maintainer was on. It sent none while it was a separate application. Nothing on the
            // server reads it; it is for whoever reads the log. Set on a supplied client too, unless
            // that client already names itself.
            // ###########################################################################################
            if (!this.thisHttp.DefaultRequestHeaders.Contains("User-Agent"))
                this.thisHttp.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", ReviewApiClient.UserAgent);
        }

        // What every request sends as its User-Agent - see the constructor. CRT's ONE definition of
        // it (OnlineServices.VersionForServer), not a copy (code review, 2026-09-29): the check-in,
        // the board views and the review API's sessions must name a build the same way, or "which
        // build was the maintainer on" stops matching the other tables when one format changes.
        internal static string UserAgent => OnlineServices.VersionForServer;

        public void Dispose() => this.thisHttp.Dispose();

        // ###########################################################################################
        // Signs in. A wrong password and an unreachable server are different outcomes.
        // ###########################################################################################
        public async Task<ReviewApiResult<ReviewSession>> LoginAsync(
            string email,
            string password,
            CancellationToken cancellationToken = default)
        {
            string body = JsonSerializer.Serialize(new { email, password });

            using var content = new StringContent(body, Encoding.UTF8, "application/json");

            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Post, ReviewApiRoutes.Login(this.thisBaseAddress))
                    {
                        Content = content
                    },
                    ReviewApiParser.ParseLogin,
                    session: null,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        // ###########################################################################################
        // Asks for a password reset link.
        //
        // *** THE SERVER ALWAYS ANSWERS 202, whether or not the address is known, and this must
        // NOT try to be more helpful than that. *** Reporting "no such account" would turn the
        // sign-in screen into a tool for discovering which addresses hold maintainer accounts -
        // and those are the accounts that can publish to every user of a system.
        //
        // So there is nothing to parse: a 202 is the whole answer, and the caller shows the same
        // sentence either way.
        // ###########################################################################################
        public async Task<ReviewApiResult<string>> ForgotPasswordAsync(
            string email,
            CancellationToken cancellationToken = default)
        {
            string body = JsonSerializer.Serialize(new { email });

            try
            {
                using var content = new StringContent(body, Encoding.UTF8, "application/json");

                using var request = new HttpRequestMessage(
                    HttpMethod.Post, ReviewApiRoutes.ForgotPassword(this.thisBaseAddress))
                {
                    Content = content
                };

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                ReviewApiResult<string>? failure = ReviewApiClient.StatusFailure<string>(response.StatusCode);

                if (failure is not null)
                    return failure;

                // The server's own sentence if it sent one, so the wording lives in one place.
                string payload = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                return ReviewApiResult<string>.Ok(
                    ReviewApiParser.ParseMessage(payload)
                        ?? "If that address has an account, a reset link is on its way to it.");
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                return ReviewApiResult<string>.Failed(
                    ReviewApiFailure.Unreachable, "The server did not answer in time.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<string>.Failed(
                    ReviewApiFailure.Unreachable, $"Could not reach the server: {exception.Message}");
            }
        }

        // ###########################################################################################
        // COMPLETE A PASSWORD RESET - the code from the mail, plus the new password.
        //
        // *** THIS IS THE OTHER HALF OF A FLOW THAT HAD NO ENDING. *** ForgotPasswordAsync has
        // existed since Phase 3 and sends the mail; nothing could ever redeem what it sent,
        // because the mail carried a link to an unmapped path and the maintainer had no way to submit a
        // code. So a maintainer locked out of the one tool that can publish stayed locked out.
        //
        // *** THE SERVER'S OWN MESSAGE IS READ BACK, SUCCESS OR FAILURE. *** Expired, already
        // used, not valid and "password rejected" each have a distinct sentence there, and every
        // one of them tells the person something they can act on. Substituting a generic failure
        // here would throw away the only useful part of the answer.
        // ###########################################################################################
        public async Task<ReviewApiResult<string>> ResetPasswordAsync(
            string code,
            string newPassword,
            CancellationToken cancellationToken = default)
        {
            string body = JsonSerializer.Serialize(new { token = code, newPassword });

            try
            {
                using var content = new StringContent(body, Encoding.UTF8, "application/json");

                using var request = new HttpRequestMessage(
                    HttpMethod.Post, ReviewApiRoutes.ResetPassword(this.thisBaseAddress))
                {
                    Content = content
                };

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                string payload =
                    await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                if (response.IsSuccessStatusCode)
                {
                    return ReviewApiResult<string>.Ok(
                        ReviewApiParser.ParseMessage(payload)
                            ?? "Your password is set. You can sign in with it now.");
                }

                // *** THE 400 BODY IS THE USEFUL PART, so StatusFailure is not used for it. ***
                // That helper maps a 400 to a generic sentence, which would hide "that code has
                // expired" behind "something went wrong" - and the difference decides whether the
                // person asks for a new code or retypes their password.
                return ReviewApiResult<string>.Failed(
                    ReviewApiFailure.Refused,
                    ReviewApiParser.ParseMessage(payload)
                        ?? ReviewApiParser.ParseFirstError(payload)
                        ?? "That code was not accepted. Ask for a new one.");
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                return ReviewApiResult<string>.Failed(
                    ReviewApiFailure.Unreachable, "The server did not answer in time.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<string>.Failed(
                    ReviewApiFailure.Unreachable, $"Could not reach the server: {exception.Message}");
            }
        }

        // ###########################################################################################
        // ACCEPTING AN INVITATION (2026-09-27): the code from the mail, a name and a password. Like
        // the password reset it runs before anybody is signed in, and the 400's own sentence is the
        // useful part ("that invitation has expired - ask the administrator for a new one"), so it
        // is read back rather than replaced by a status code's wording.
        // ###########################################################################################
        public async Task<ReviewApiResult<AcceptInvitationAnswer>> AcceptInvitationAsync(
            string code,
            string displayName,
            string password,
            CancellationToken cancellationToken = default)
        {
            string body = ReviewApiClient.Body(new AcceptInvitationRequest(code, displayName, password));

            try
            {
                using var content = new StringContent(body, Encoding.UTF8, "application/json");

                using var request = new HttpRequestMessage(
                    HttpMethod.Post, ReviewApiRoutes.AcceptInvitation(this.thisBaseAddress))
                {
                    Content = content
                };

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                string payload =
                    await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                if (response.IsSuccessStatusCode)
                {
                    return ReviewApiParser.ParseAcceptInvitation(payload) is AcceptInvitationAnswer answer
                        ? ReviewApiResult<AcceptInvitationAnswer>.Ok(answer)
                        : ReviewApiResult<AcceptInvitationAnswer>.Failed(
                            ReviewApiFailure.UnreadableAnswer, "The server's answer could not be read. Try signing in.");
                }

                if (response.StatusCode == HttpStatusCode.BadRequest)
                {
                    return ReviewApiResult<AcceptInvitationAnswer>.Failed(
                        ReviewApiFailure.Refused,
                        ReviewApiParser.ParseMessage(payload)
                            ?? ReviewApiParser.ParseFirstError(payload)
                            ?? "That invitation was not accepted.");
                }

                return ReviewApiClient.StatusFailure<AcceptInvitationAnswer>(response.StatusCode)
                    ?? ReviewApiResult<AcceptInvitationAnswer>.Failed(ReviewApiFailure.ServerError, "The server could not accept the invitation.");
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                return ReviewApiResult<AcceptInvitationAnswer>.Failed(
                    ReviewApiFailure.Unreachable, "The server did not answer in time.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<AcceptInvitationAnswer>.Failed(
                    ReviewApiFailure.Unreachable, $"Could not reach the server: {exception.Message}");
            }
        }

        // ###########################################################################################
        // SIGN OUT - revoke this session on the SERVER, not just locally.
        //
        // *** CLEARING THE LOCAL COPY ALONE WOULD NOT BE SIGNING OUT. *** The token would stay
        // live server-side for the rest of its (sliding) lifetime, so anyone who recovered the
        // file - or a copy of it made earlier - would still hold a working credential. The whole
        // point of the button is that the project owner can END a session, and only the server can
        // do that.
        //
        // *** THE FIELD IS `refreshToken`, AND THAT NAME IS THE TRAP ReviewSession's HEADER
        // DESCRIBES. *** It is not exchanged for anything; it IS the session token. The endpoint
        // looks it up by hash and revokes the row.
        //
        // *** RETURNS NOTHING AND FAILS AT NOTHING. *** The endpoint answers 204 for an unknown or
        // already-revoked token on purpose, so there is no failure worth surfacing. More
        // importantly, the caller must forget the session LOCALLY whatever happens here - a
        // maintainer who cannot reach the server must still be able to sign out of their own
        // machine, which is exactly when they would most want to.
        // ###########################################################################################
        public async Task LogoutAsync(
            ReviewSession session,
            CancellationToken cancellationToken = default)
        {
            if (session is null || string.IsNullOrWhiteSpace(session.BearerToken))
                return;

            string body = JsonSerializer.Serialize(new { refreshToken = session.BearerToken });

            try
            {
                using var content = new StringContent(body, Encoding.UTF8, "application/json");

                using var request = new HttpRequestMessage(
                    HttpMethod.Post, ReviewApiRoutes.Logout(this.thisBaseAddress))
                {
                    Content = content
                };

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception)
            {
                // Deliberately swallowed - see the header. Signing out locally must always succeed.
            }
        }

        // ###########################################################################################
        // The queue.
        // ###########################################################################################
        public async Task<ReviewApiResult<ReviewQueueResponse>> GetQueueAsync(
            ReviewSession session,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.Queue(this.thisBaseAddress)),
                    ReviewApiParser.ParseQueue,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        // ###########################################################################################
        // One submission, with its change summary and findings.
        // ###########################################################################################
        public async Task<ReviewApiResult<ReviewSubmissionDetail>> GetSubmissionAsync(
            ReviewSession session,
            long submissionId,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(
                        HttpMethod.Get,
                        ReviewApiRoutes.Submission(this.thisBaseAddress, submissionId)),
                    ReviewApiParser.ParseSubmission,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        // ###########################################################################################
        // The three review DECISIONS (task 5).
        //
        // *** APPROVE IS THE IRREVERSIBLE ONE. *** It publishes, and with no retained revision the
        // only way back is a correction. The client cannot make that any safer - the server
        // refuses what must be refused - but it is worth knowing which of these three is which.
        //
        // A CONFLICT is reported distinctly, because "somebody else already decided this" is a
        // completely different message from "that failed": the maintainer should refresh and look at
        // what was decided, not retry.
        // ###########################################################################################
        //
        // `shownRemovals` is the list of files the maintainer was shown the publish would remove
        // (the detail's FileRemovalPreview). The server refuses the approval when that is no longer
        // the list it would remove, so nothing goes that was not on screen.
        public Task<ReviewApiResult<ReviewDecisionResult>> ApproveAsync(
            ReviewSession session,
            long submissionId,
            IReadOnlyList<string>? shownRemovals = null,
            CancellationToken cancellationToken = default) =>
            this.DecideAsync(
                ReviewApiRoutes.Approve(this.thisBaseAddress, submissionId),
                comment: null,
                session,
                cancellationToken,
                shownRemovals ?? []);

        // ###########################################################################################
        // The maintainer's table (2026-09-25). A change is sent with the amendment version the table
        // was opened at, and refused as a conflict if somebody else changed the submission since.
        // ###########################################################################################
        public async Task<ReviewApiResult<ReviewTableData>> GetTableAsync(
            ReviewSession session,
            long submissionId,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.SubmissionTable(this.thisBaseAddress, submissionId)),
                    ReviewApiParser.ParseTable,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        public Task<ReviewApiResult<ReviewAmendResult>> AmendAsync(
            ReviewSession session,
            long submissionId,
            int expectedVersion,
            SubmissionRows rows,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.SubmissionAmend(this.thisBaseAddress, submissionId),
                new AmendRequest(expectedVersion, rows),
                ReviewApiParser.ParseAmend,
                session,
                cancellationToken);

        public Task<ReviewApiResult<ReviewDecisionResult>> RejectAsync(
            ReviewSession session,
            long submissionId,
            string comment,
            CancellationToken cancellationToken = default) =>
            this.DecideAsync(
                ReviewApiRoutes.Reject(this.thisBaseAddress, submissionId),
                comment,
                session,
                cancellationToken);

        public Task<ReviewApiResult<ReviewDecisionResult>> RequestChangesAsync(
            ReviewSession session,
            long submissionId,
            string comment,
            CancellationToken cancellationToken = default) =>
            this.DecideAsync(
                ReviewApiRoutes.RequestChanges(this.thisBaseAddress, submissionId),
                comment,
                session,
                cancellationToken);

        // ###########################################################################################
        // One decision. The three differ only in their route and whether they carry a comment.
        // ###########################################################################################
        private async Task<ReviewApiResult<ReviewDecisionResult>> DecideAsync(
            string route,
            string? comment,
            ReviewSession? session,
            CancellationToken cancellationToken,
            IReadOnlyList<string>? expectedRemovals = null)
        {
            string body = ReviewApiClient.Body(new ReviewDecisionRequest(comment, expectedRemovals));

            try
            {
                using var content = new StringContent(body, Encoding.UTF8, "application/json");
                using var request = new HttpRequestMessage(HttpMethod.Post, route) { Content = content };

                if (session is not null)
                {
                    request.Headers.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", session.BearerToken);
                }

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                if (response.StatusCode == HttpStatusCode.Conflict)
                {
                    // Somebody else decided it first. A distinct failure because the answer is
                    // "refresh and see what they did", not "try again".
                    return ReviewApiResult<ReviewDecisionResult>.Failed(
                        ReviewApiFailure.Conflict,
                        await ReviewApiClient.ReadErrorAsync(response, cancellationToken)
                            ?? "This submission has already been decided.");
                }

                if (response.StatusCode == HttpStatusCode.BadRequest)
                {
                    // The server refused the request itself - almost always a reason that is too
                    // short. Its own message is the useful one.
                    return ReviewApiResult<ReviewDecisionResult>.Failed(
                        ReviewApiFailure.Refused,
                        await ReviewApiClient.ReadErrorAsync(response, cancellationToken)
                            ?? "The server refused this decision.");
                }

                ReviewApiResult<ReviewDecisionResult>? failure =
                    ReviewApiClient.StatusFailure<ReviewDecisionResult>(response.StatusCode);

                if (failure is not null)
                    return failure;

                string payload = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                ReviewDecisionResult? parsed = ReviewApiParser.ParseDecision(payload);

                return parsed is null
                    ? ReviewApiResult<ReviewDecisionResult>.Failed(
                        ReviewApiFailure.UnreadableAnswer,
                        "The server sent an answer this version does not understand.")
                    : ReviewApiResult<ReviewDecisionResult>.Ok(parsed);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                // *** A TIMEOUT ON A DECISION IS NOT A FAILURE TO DECIDE. *** The request may well
                // have reached the server and published. Saying so matters more here than
                // anywhere else in this client, because the obvious response to "that failed" is
                // to press the button again.
                return ReviewApiResult<ReviewDecisionResult>.Failed(
                    ReviewApiFailure.Unreachable,
                    "The server did not answer in time. Refresh the queue before trying again - " +
                    "the decision may already have been recorded.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<ReviewDecisionResult>.Failed(
                    ReviewApiFailure.Unreachable,
                    $"Could not reach the server: {exception.Message}");
            }
        }

        // The server's own error sentence, when it sent one. Its message is written for the person
        // reading; a generic one here would throw that away.
        private static async Task<string?> ReadErrorAsync(
            HttpResponseMessage response, CancellationToken cancellationToken)
        {
            try
            {
                string body = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                return ReviewApiParser.ParseError(body);
            }
            catch (Exception exception) when (exception is HttpRequestException or JsonException)
            {
                return null;
            }
        }

        // ###########################################################################################
        // The administrator's lists and changes (Phase 6 roles). GETs through the shared send path;
        // the two changes POST a body and read the server's own sentence back on a refusal, the
        // way a decision does.
        // ###########################################################################################
        public async Task<ReviewApiResult<ReviewSystemsResponse>> GetSystemsAsync(
            ReviewSession session,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.AdminSystems(this.thisBaseAddress)),
                    ReviewApiParser.ParseSystems,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        public async Task<ReviewApiResult<ReviewAccountsResponse>> GetAccountsAsync(
            ReviewSession session,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.AdminAccounts(this.thisBaseAddress)),
                    ReviewApiParser.ParseAccounts,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        public Task<ReviewApiResult<string>> AddMaintainerAsync(
            ReviewSession session,
            string systemId,
            long accountId,
            CancellationToken cancellationToken = default) =>
            this.PostMaintainerChangeAsync(
                ReviewApiRoutes.AdminAddMaintainer(this.thisBaseAddress), new MaintainerChangeRequest(systemId, accountId), session, cancellationToken);

        public Task<ReviewApiResult<string>> RemoveMaintainerAsync(
            ReviewSession session,
            string systemId,
            long accountId,
            CancellationToken cancellationToken = default) =>
            this.PostMaintainerChangeAsync(
                ReviewApiRoutes.AdminRemoveMaintainer(this.thisBaseAddress), new MaintainerChangeRequest(systemId, accountId), session, cancellationToken);

        // Inviting an address with no account (2026-09-27) - the answer is the server's sentence.
        public Task<ReviewApiResult<string>> InviteMaintainerAsync(
            ReviewSession session,
            string systemId,
            string email,
            CancellationToken cancellationToken = default) =>
            this.PostMaintainerChangeAsync(
                ReviewApiRoutes.AdminInviteMaintainer(this.thisBaseAddress), new MaintainerInviteRequest(systemId, email), session, cancellationToken);

        public Task<ReviewApiResult<string>> WithdrawInvitationAsync(
            ReviewSession session,
            long invitationId,
            CancellationToken cancellationToken = default) =>
            this.PostMaintainerChangeAsync(
                ReviewApiRoutes.AdminWithdrawInvitation(this.thisBaseAddress), new InvitationWithdrawRequest(invitationId), session, cancellationToken);

        // A pool change, an invitation or a withdrawal: the server's own sentence comes back on
        // success too when it sends one ("An invitation is on its way to ..."), else "Done.".
        private async Task<ReviewApiResult<string>> PostMaintainerChangeAsync(
            string route,
            object change,
            ReviewSession session,
            CancellationToken cancellationToken)
        {
            string body = ReviewApiClient.Body(change);

            try
            {
                using var content = new StringContent(body, Encoding.UTF8, "application/json");
                using var request = new HttpRequestMessage(HttpMethod.Post, route) { Content = content };

                request.Headers.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", session.BearerToken);

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                // The server's refusals here are SENTENCES written for the administrator - "has
                // not verified their address", "no such system" - so they are read back rather
                // than replaced by a status code's generic wording.
                if (response.StatusCode is HttpStatusCode.BadRequest or HttpStatusCode.NotFound or HttpStatusCode.Forbidden)
                {
                    return ReviewApiResult<string>.Failed(
                        response.StatusCode == HttpStatusCode.Forbidden ? ReviewApiFailure.NotPermitted : ReviewApiFailure.Refused,
                        await ReviewApiClient.ReadErrorAsync(response, cancellationToken)
                            ?? "The server refused this change.");
                }

                ReviewApiResult<string>? failure = ReviewApiClient.StatusFailure<string>(response.StatusCode);

                if (failure is not null)
                    return failure;

                string payload = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                return ReviewApiResult<string>.Ok(ReviewApiParser.ParseMessage(payload) ?? "Done.");
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                return ReviewApiResult<string>.Failed(
                    ReviewApiFailure.Unreachable, "The server did not answer in time. Refresh to see what it holds.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<string>.Failed(
                    ReviewApiFailure.Unreachable, $"Could not reach the server: {exception.Message}");
            }
        }

        // ###########################################################################################
        // BETA to production (2026-09-25). The list through the shared GET path; the plan and the
        // publish POST a body and read the server's own sentence back on a refusal - a 409 here
        // says BETA changed since it was checked, which the maintainer must read rather than retry.
        // ###########################################################################################
        public async Task<ReviewApiResult<ProductionListResponse>> GetProductionListAsync(
            ReviewSession session,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.ProductionList(this.thisBaseAddress)),
                    ReviewApiParser.ParseProductionList,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        // A submission's file tree: the BETA data after approving it (2026-09-28).
        public async Task<ReviewApiResult<SubmissionFilesAnswer>> GetSubmissionFilesAsync(
            ReviewSession session,
            long submissionId,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.SubmissionFiles(this.thisBaseAddress, submissionId)),
                    ReviewApiParser.ParseSubmissionFiles,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        // ###########################################################################################
        // A file as a PUBLISHED data tree holds it - read from its public address, exactly as CRT
        // downloads it (2026-09-28). No session: the address is public, and the bearer token is
        // never sent anywhere but the API.
        // ###########################################################################################
        public async Task<ReviewApiResult<byte[]>> GetPublishedDataFileAsync(
            string? dataBaseUrl,
            string relativePath,
            CancellationToken cancellationToken = default)
        {
            string? address = ReviewApiRoutes.PublishedDataFile(dataBaseUrl, relativePath);

            if (address is null)
            {
                return ReviewApiResult<byte[]>.Failed(
                    ReviewApiFailure.UnreadableAnswer,
                    "The server did not say where its data is published, so this file cannot be opened from here.");
            }

            return await this
                .SendBytesAsync(() => new HttpRequestMessage(HttpMethod.Get, address), session: null, cancellationToken)
                .ConfigureAwait(false);
        }

        public Task<ReviewApiResult<ProductionPlanView>> GetProductionPlanAsync(
            ReviewSession session,
            string systemId,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.ProductionPlan(this.thisBaseAddress),
                new ProductionPlanRequest(systemId),
                ReviewApiParser.ParseProductionPlan,
                session,
                cancellationToken);

        // ###########################################################################################
        // Rolling a BETA board back and returning its submissions to the queue (2026-09-27). The
        // plan is shown first; the rollback carries the REASON, which the server requires because
        // it is the contributor's only feedback.
        // ###########################################################################################
        public Task<ReviewApiResult<BetaRollbackPlanView>> GetBetaRollbackPlanAsync(
            ReviewSession session,
            string systemId,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.BetaRollbackPlan(this.thisBaseAddress),
                new BetaRollbackPlanRequest(systemId),
                ReviewApiParser.ParseBetaRollbackPlan,
                session,
                cancellationToken);

        // `reject`: Beta > Prod's "Reject" (2026-09-28) - the same rollback, its submissions
        // rejected rather than returned to the queue.
        public Task<ReviewApiResult<BetaRollbackResult>> RollBackBetaAsync(
            ReviewSession session,
            string systemId,
            string comment,
            bool reject = false,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.BetaRollback(this.thisBaseAddress),
                new BetaRollbackRequest(systemId, comment, reject),
                ReviewApiParser.ParseBetaRollback,
                session,
                cancellationToken);

        // `shownRemovals`: the files the plan said this would remove from production - see
        // ApproveAsync.
        public Task<ReviewApiResult<ProductionPublishResult>> PublishToProductionAsync(
            ReviewSession session,
            string systemId,
            string expectedBetaContentHash,
            IReadOnlyList<string>? shownRemovals = null,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.ProductionPublish(this.thisBaseAddress),
                new ProductionPublishRequest(systemId, expectedBetaContentHash, shownRemovals ?? []),
                ReviewApiParser.ParseProductionPublish,
                session,
                cancellationToken);

        // ###########################################################################################
        // The "Systems" screen (2026-09-27): every system through the shared GET path, and one
        // system's facts as a POST with its id, reading the server's own sentence back on a refusal.
        // ###########################################################################################
        public async Task<ReviewApiResult<SystemOverviewAnswer>> GetSystemOverviewAsync(
            ReviewSession session,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.Systems(this.thisBaseAddress)),
                    ReviewApiParser.ParseSystemOverview,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        public Task<ReviewApiResult<SystemDetailAnswer>> GetSystemDetailAsync(
            ReviewSession session,
            string systemId,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.SystemDetail(this.thisBaseAddress),
                new SystemDetailRequest(systemId),
                ReviewApiParser.ParseSystemDetail,
                session,
                cancellationToken);

        // ###########################################################################################
        // The drop-down lists as CRT shows them from BETA, and the systems not in them yet; and
        // saving where one goes (owner request, 2026-09-27). A refusal reads the server's sentence.
        // ###########################################################################################
        public async Task<ReviewApiResult<SystemListingAnswer>> GetSystemListingAsync(
            ReviewSession session,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.SystemListing(this.thisBaseAddress)),
                    ReviewApiParser.ParseSystemListing,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        public Task<ReviewApiResult<SetPlacementAnswer>> SetPlacementAsync(
            ReviewSession session,
            SetPlacementRequest request,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.SystemListing(this.thisBaseAddress),
                request,
                ReviewApiParser.ParseSetPlacement,
                session,
                cancellationToken);

        // ###########################################################################################
        // The administrator's "Unused files" screen (2026-09-25). The list reads every workbook in
        // the tree on the server, so it takes a few seconds.
        // ###########################################################################################
        public async Task<ReviewApiResult<UnusedFileListing>> GetUnusedFilesAsync(
            ReviewSession session,
            string tree,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendAsync(
                    () => new HttpRequestMessage(HttpMethod.Get, ReviewApiRoutes.AdminUnusedFiles(this.thisBaseAddress, tree)),
                    ReviewApiParser.ParseUnusedFiles,
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        public Task<ReviewApiResult<UnusedFileRemovalResult>> RemoveUnusedFilesAsync(
            ReviewSession session,
            string tree,
            IReadOnlyList<string> files,
            CancellationToken cancellationToken = default) =>
            this.PostProductionAsync(
                ReviewApiRoutes.AdminRemoveUnusedFiles(this.thisBaseAddress),
                new UnusedFilesRemoveRequest(tree, files),
                ReviewApiParser.ParseUnusedFileRemoval,
                session,
                cancellationToken);

        // ###########################################################################################
        // A request body: one of CRT.Data's ReviewApiContract records, in the JSON the server reads
        // (ReviewApiContract.WireSettings, which the server applies too). Serialised by its RUNTIME
        // type, since callers pass it as `object`. Internal so ReviewWireContractTests can send
        // exactly what this client sends.
        // ###########################################################################################
        internal static string Body(object body) =>
            JsonSerializer.Serialize(body, body.GetType(), ReviewApiContract.WireSettings);

        private async Task<ReviewApiResult<T>> PostProductionAsync<T>(
            string route,
            object body,
            Func<string, T?> parse,
            ReviewSession session,
            CancellationToken cancellationToken)
            where T : class
        {
            try
            {
                using var content = new StringContent(ReviewApiClient.Body(body), Encoding.UTF8, "application/json");
                using var request = new HttpRequestMessage(HttpMethod.Post, route) { Content = content };

                request.Headers.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", session.BearerToken);

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                // Every refusal here carries a sentence written for the maintainer.
                if (response.StatusCode is HttpStatusCode.BadRequest or HttpStatusCode.NotFound
                    or HttpStatusCode.Forbidden or HttpStatusCode.Conflict or HttpStatusCode.ServiceUnavailable)
                {
                    ReviewApiFailure failure = response.StatusCode switch
                    {
                        HttpStatusCode.Forbidden => ReviewApiFailure.NotPermitted,
                        HttpStatusCode.Conflict => ReviewApiFailure.Conflict,
                        HttpStatusCode.NotFound => ReviewApiFailure.NotFound,
                        _ => ReviewApiFailure.Refused
                    };

                    return ReviewApiResult<T>.Failed(
                        failure,
                        await ReviewApiClient.ReadErrorAsync(response, cancellationToken) ?? "The server refused this.");
                }

                ReviewApiResult<T>? statusFailure = ReviewApiClient.StatusFailure<T>(response.StatusCode);

                if (statusFailure is not null)
                    return statusFailure;

                T? parsed = parse(await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false));

                return parsed is null
                    ? ReviewApiResult<T>.Failed(ReviewApiFailure.UnreadableAnswer, "The server sent an answer this version does not understand.")
                    : ReviewApiResult<T>.Ok(parsed);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                // As for a decision: a publish that timed out may well have happened.
                return ReviewApiResult<T>.Failed(
                    ReviewApiFailure.Unreachable,
                    "The server did not answer in time. Refresh before trying again - it may already have been done.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<T>.Failed(ReviewApiFailure.Unreachable, $"Could not reach the server: {exception.Message}");
            }
        }

        // ###########################################################################################
        // The bytes a contributor UPLOADED for this submission, addressed by hash (task 4).
        //
        // Returns BYTES rather than a decoded image: this project is the one that must not decode
        // an image outside the UI thread's control, and more to the point the Maintainer tab's own tests
        // can assert on bytes with no display. The window turns them into a Bitmap.
        // ###########################################################################################
        public async Task<ReviewApiResult<byte[]>> GetSubmittedAssetAsync(
            ReviewSession session,
            long submissionId,
            string hash,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendBytesAsync(
                    () => new HttpRequestMessage(
                        HttpMethod.Get,
                        ReviewApiRoutes.SubmittedAsset(this.thisBaseAddress, submissionId, hash)),
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        // ###########################################################################################
        // The bytes CURRENTLY PUBLISHED at this path, for the "before" side of a comparison.
        //
        // *** A 404 HERE IS ORDINARY, not a fault. *** A newly added image has no published
        // counterpart, which is the commonest case on the very submissions a maintainer most wants
        // to look at. The caller must render "nothing there before" rather than an error.
        // ###########################################################################################
        public async Task<ReviewApiResult<byte[]>> GetPublishedAssetAsync(
            ReviewSession session,
            long submissionId,
            string relativePath,
            CancellationToken cancellationToken = default)
        {
            return await this
                .SendBytesAsync(
                    () => new HttpRequestMessage(
                        HttpMethod.Get,
                        ReviewApiRoutes.PublishedAsset(this.thisBaseAddress, submissionId, relativePath)),
                    session,
                    cancellationToken)
                .ConfigureAwait(false);
        }

        // ###########################################################################################
        // One request, turned into a result.
        //
        // The request is built by a FACTORY rather than passed in, because an HttpRequestMessage
        // cannot be sent twice - taking one as a parameter would make a retry here impossible to
        // add later without a subtle bug.
        // ###########################################################################################
        private async Task<ReviewApiResult<T>> SendAsync<T>(
            Func<HttpRequestMessage> buildRequest,
            Func<string, T?> parse,
            ReviewSession? session,
            CancellationToken cancellationToken)
            where T : class
        {
            try
            {
                using HttpRequestMessage request = buildRequest();

                if (session is not null)
                {
                    request.Headers.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", session.BearerToken);
                }

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                // Shared with the bytes path so the two cannot come to disagree about what a
                // status code means to the person reading it.
                ReviewApiResult<T>? failure = ReviewApiClient.StatusFailure<T>(response.StatusCode);

                if (failure is not null)
                    return failure;

                string payload = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                T? parsed = parse(payload);

                // A success status carrying something unreadable is NOT a success. Saying so
                // beats handing the window a null it would have to guess about.
                return parsed is null
                    ? ReviewApiResult<T>.Failed(
                        ReviewApiFailure.UnreadableAnswer,
                        "The server sent an answer this version does not understand.")
                    : ReviewApiResult<T>.Ok(parsed);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                // The caller asked to stop; not a failure to report.
                throw;
            }
            catch (TaskCanceledException)
            {
                // HttpClient reports its own TIMEOUT as a cancellation, which is why this is
                // caught separately from the guarded clause above. Without that distinction a
                // timeout reads as "the user cancelled" and is silently swallowed.
                return ReviewApiResult<T>.Failed(
                    ReviewApiFailure.Unreachable,
                    "The server did not answer in time.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<T>.Failed(
                    ReviewApiFailure.Unreachable,
                    $"Could not reach the server: {exception.Message}");
            }
        }

        // ###########################################################################################
        // The same request handling, for a response whose body is BYTES rather than JSON.
        //
        // Shares nothing with SendAsync above except through StatusFailure, deliberately: merging
        // them would mean a generic method branching on whether T is byte[], which is the kind of
        // cleverness that makes the next change to either path risky. The one thing that MUST NOT
        // drift between them - what each status code means to the person reading - is the one
        // thing that is shared.
        //
        // *** A 404 IS RETURNED AS NotFound, NOT AS AN ERROR. *** An image with no published
        // counterpart is the commonest case on an interesting submission, and reporting it as a
        // server error would put a red message where "this is new" belongs.
        // ###########################################################################################
        private async Task<ReviewApiResult<byte[]>> SendBytesAsync(
            Func<HttpRequestMessage> buildRequest,
            ReviewSession? session,
            CancellationToken cancellationToken)
        {
            try
            {
                using HttpRequestMessage request = buildRequest();

                if (session is not null)
                {
                    request.Headers.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", session.BearerToken);
                }

                using HttpResponseMessage response =
                    await this.thisHttp.SendAsync(request, cancellationToken).ConfigureAwait(false);

                if (response.StatusCode == HttpStatusCode.NotFound)
                {
                    return ReviewApiResult<byte[]>.Failed(
                        ReviewApiFailure.NotFound,
                        "There is no published file at that path.");
                }

                ReviewApiResult<byte[]>? failure = ReviewApiClient.StatusFailure<byte[]>(response.StatusCode);

                if (failure is not null)
                    return failure;

                byte[] bytes = await response.Content
                    .ReadAsByteArrayAsync(cancellationToken)
                    .ConfigureAwait(false);

                // A success carrying NO bytes is not a success. An empty image decodes to nothing
                // and would draw as a blank panel, which reads as "unchanged" - the one wrong
                // answer a comparison must never give.
                return bytes.Length == 0
                    ? ReviewApiResult<byte[]>.Failed(
                        ReviewApiFailure.UnreadableAnswer,
                        "The server sent an empty file.")
                    : ReviewApiResult<byte[]>.Ok(bytes);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (TaskCanceledException)
            {
                return ReviewApiResult<byte[]>.Failed(
                    ReviewApiFailure.Unreachable,
                    "The server did not answer in time.");
            }
            catch (HttpRequestException exception)
            {
                return ReviewApiResult<byte[]>.Failed(
                    ReviewApiFailure.Unreachable,
                    $"Could not reach the server: {exception.Message}");
            }
        }

        // ###########################################################################################
        // What a non-success status means to the person reading, or null when it is a success.
        //
        // Shared by both send paths so the two cannot come to disagree about what a 403 means -
        // the message is what a maintainer acts on, and "sign in again" for an account that is
        // signed in perfectly well is a loop they cannot escape.
        // ###########################################################################################
        private static ReviewApiResult<T>? StatusFailure<T>(HttpStatusCode status)
            where T : class
        {
            if (status == HttpStatusCode.Unauthorized)
                return ReviewApiResult<T>.Failed(ReviewApiFailure.NotSignedIn, "Sign in to continue.");

            if (status == HttpStatusCode.Forbidden)
            {
                return ReviewApiResult<T>.Failed(
                    ReviewApiFailure.NotPermitted,
                    "This account is not allowed to review submissions.");
            }

            if (status == HttpStatusCode.TooManyRequests)
            {
                return ReviewApiResult<T>.Failed(
                    ReviewApiFailure.RateLimited,
                    "Too many attempts. Wait a moment and try again.");
            }

            // Anything that is not 2xx, NOT merely anything 4xx and up. HttpClient follows
            // redirects itself, so a 3xx arriving here means it was told not to or ran out of
            // hops - which is a failed request, not a success with an odd code. Matching
            // IsSuccessStatusCode exactly keeps this faithful to the version it replaced.
            if ((int)status is < 200 or > 299)
            {
                return ReviewApiResult<T>.Failed(
                    ReviewApiFailure.ServerError,
                    $"The server answered {(int)status}.");
            }

            return null;
        }
    }

    public enum ReviewApiFailure
    {
        None,
        Unreachable,
        NotSignedIn,
        NotPermitted,
        RateLimited,
        UnreadableAnswer,
        ServerError,

        // Somebody else decided this submission first. Distinct from ServerError because the
        // answer is "refresh and look at what they did", not "try again".
        Conflict,

        // The server refused the request itself - most often a rejection reason too short to be
        // worth sending to the contributor. Its own message says which.
        Refused,

        // *** NOT AN ERROR IN EVERY CONTEXT. *** For a PUBLISHED asset it usually means the file
        // is new in this submission and has no counterpart to compare against, which is ordinary
        // and must be drawn as "nothing there before" rather than as a failure. Distinguished
        // from ServerError precisely so the caller can tell the two apart.
        NotFound,

        // ###########################################################################################
        // *** NOT A FAILURE OF THE REQUEST - THE WINDOW STOPPED WAITING (2026-09-28). *** Everything
        // the maintainer waits for runs under Main's BusyOverlay, which gives up after WaitLimit's
        // two minutes. The server may well have finished regardless, so a caller that changed
        // something must look again before it says what happened (ServerWait).
        // ###########################################################################################
        TimedOut
    }

    // ###########################################################################################
    // Either a value or a reason there is not one.
    //
    // Deliberately not "a value plus a nullable error": a caller holding both has to remember
    // which to check, and the one that gets forgotten is the error.
    // ###########################################################################################
    public sealed class ReviewApiResult<T> where T : class
    {
        private ReviewApiResult(T? value, ReviewApiFailure failure, string message)
        {
            this.Value = value;
            this.Failure = failure;
            this.Message = message;
        }

        public T? Value { get; }

        public ReviewApiFailure Failure { get; }

        // Written for the maintainer to read on screen, not for a log.
        public string Message { get; }

        public bool IsOk => this.Failure == ReviewApiFailure.None && this.Value is not null;

        public static ReviewApiResult<T> Ok(T value) => new(value, ReviewApiFailure.None, string.Empty);

        public static ReviewApiResult<T> Failed(ReviewApiFailure failure, string message) =>
            new(null, failure, message);
    }
}
