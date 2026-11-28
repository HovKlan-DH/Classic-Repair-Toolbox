using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using CRT;
using Handlers.DataHandling;
using Handlers.Online;

namespace Handlers.Collections
{
    // ###########################################################################################
    // SHARING A COLLECTION WITH THE SERVER (2026-11-26; CollectionContract says how). The untested-network rim is HttpClient;
    // everything it decides - send the collection, upload what the server asks for, send it again - is here and tested on a
    // fake handler (CollectionClientTests).
    //
    // Signed in with the account CRT is signed in to (the "Maintainer" tab's session, shared with the whole of CRT). Answers
    // a result, never throws for a refusal or a server out of reach: the collection window says it in words.
    // ###########################################################################################
    public sealed class CollectionClient
    {
        private readonly string thisRootUrl;

        private readonly HttpClient thisHttp;

        // `rootUrl`: the contribution server without "/api" (AppConfig.ContributionServerRootUrl - the test instance's for a
        // pre-release). `http`: a test's, on a fake handler.
        public CollectionClient(string rootUrl, HttpClient? http = null)
        {
            this.thisRootUrl = rootUrl.TrimEnd('/');
            this.thisHttp = http ?? SubmissionClient.CreateHttpClient(AppConfig.SubmissionUploadTimeout);
        }

        public enum Outcome
        {
            Done,
            Shared,
            Deleted,
            SignInNeeded,
            GoneFromServer,
            Refused,
            Failed
        }

        // A share's result. `SentDeleteAtUtc` and `SharedContentHash`: the deletion date sent, and the checksum of what was sent
        // - what the collection keeps, never the server's echo of the date (its database keeps milliseconds, so the echo would
        // read as a change at once).
        public sealed record Result(
            Outcome Kind,
            IReadOnlyList<string> Errors,
            CollectionAnswer? Answer = null,
            string? SharedContentHash = null,
            DateTimeOffset? SentDeleteAtUtc = null)
        {
            public string Message => string.Join(" ", this.Errors);
        }

        public const string NotReachable = "Could not reach the server - check the internet connection, and try again in a moment.";

        public const string SignInFirst = "Sign in first - an account shares collections (the account icon in the top right corner).";

        // ###########################################################################################
        // SHARES THE COLLECTION AS IT IS NOW: sent whole; the captures the server does not hold yet sent then, each from
        // `readCapture`; and the collection again. `progress` hears each capture as it goes.
        // ###########################################################################################
        public async Task<Result> ShareAsync(
            LocalCollection collection,
            string bearerToken,
            Func<string, byte[]?> readCapture,
            DateTimeOffset now,
            IProgress<string>? progress = null,
            CancellationToken cancellationToken = default)
        {
            DateTimeOffset deleteAt = CollectionSharing.DeleteAtFor(collection, now);

            // A date that passed while it was not shared again counts from today with the choice made.
            if (deleteAt <= now)
                deleteAt = CollectionRules.DeleteAt(now, collection.Lifetime);

            CollectionContentRequest request = CollectionSharing.RequestOf(collection, deleteAt);
            IReadOnlyList<string> problems = CollectionRules.ProblemsOf(request, now);

            if (problems.Count > 0)
                return new Result(Outcome.Refused, problems);

            try
            {
                for (int attempt = 0; attempt < 2; attempt++)
                {
                    progress?.Report(attempt == 0 ? "Sending the collection" : "Sending the collection again");

                    using HttpResponseMessage response = await this.SendContentAsync(collection.ServerId, collection.LocalId, request, bearerToken, cancellationToken);
                    string body = await response.Content.ReadAsStringAsync(cancellationToken);

                    if (response.IsSuccessStatusCode)
                    {
                        CollectionAnswer? answer = JsonSerializer.Deserialize<CollectionAnswer>(body, ReviewApiContract.WireSettings);

                        return answer is null
                            ? new Result(Outcome.Failed, ["The server answered nothing CRT could read."])
                            : new Result(Outcome.Shared, [], answer, CollectionSharing.ContentHash(request), deleteAt);
                    }

                    if (response.StatusCode == HttpStatusCode.Conflict && attempt == 0 &&
                        CollectionClient.RefusalOf(body).MissingImages is { Count: > 0 } missing)
                    {
                        if (await this.UploadAsync(missing, readCapture, bearerToken, progress, cancellationToken) is { } failed)
                            return failed;

                        continue;
                    }

                    return CollectionClient.ResultOfRefusal(response.StatusCode, body, collection.ServerId is not null);
                }

                return new Result(Outcome.Failed, ["The server still asked for captures after they were sent."]);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
            {
                return new Result(Outcome.Failed, [CollectionClient.NotReachable]);
            }
        }

        // ###########################################################################################
        // Takes the collection's page off the server. One the server deleted already is deleted.
        // ###########################################################################################
        public async Task<Result> DeleteAsync(string serverId, string bearerToken, CancellationToken cancellationToken = default)
        {
            try
            {
                using var message = new HttpRequestMessage(HttpMethod.Delete, $"{this.thisRootUrl}{CollectionRoutes.Collections}/{Uri.EscapeDataString(serverId)}");
                message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

                using HttpResponseMessage response = await this.thisHttp.SendAsync(message, cancellationToken);

                if (response.IsSuccessStatusCode || response.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.Gone)
                    return new Result(Outcome.Deleted, []);

                return CollectionClient.ResultOfRefusal(response.StatusCode, await response.Content.ReadAsStringAsync(cancellationToken), wasShared: true);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
            {
                return new Result(Outcome.Failed, [CollectionClient.NotReachable]);
            }
        }

        // ###########################################################################################
        // "STOP SHARING" (owner decision, 2026-11-27: "Offline, same link back"), and sharing it again with nothing changed - its
        // page offline or back, under the same link, its captures and comments kept on the server either way. Done, or what
        // came of it; a page the server no longer has (410) is GoneFromServer.
        // ###########################################################################################
        public Task<Result> SetSharingAsync(string serverId, bool shared, string bearerToken, CancellationToken cancellationToken = default) =>
            this.CommentRequestAsync(HttpMethod.Put, CollectionRoutes.SharingOf(Uri.EscapeDataString(serverId)), bearerToken, new CollectionSharingRequest(shared), cancellationToken);

        // ###########################################################################################
        // "DELETE IMAGE" on a shared collection (owner request, 2026-11-28: "should also remove it right away from the shared
        // page, and do not touch the expiry date"; server 16.2.0): that one capture off its page now, nothing else sent - its
        // deletion date, its sharing and any change not shared yet stay as they are. Done; a page the server no longer has
        // (410) is GoneFromServer.
        // ###########################################################################################
        public Task<Result> RemoveImageAsync(string serverId, string sha256, string bearerToken, CancellationToken cancellationToken = default) =>
            this.CommentRequestAsync(
                HttpMethod.Delete,
                CollectionRoutes.ImageOf(Uri.EscapeDataString(serverId), Uri.EscapeDataString(sha256)),
                bearerToken,
                null,
                cancellationToken);

        // ###########################################################################################
        // COMMENTS (step 3, 2026-11-26) - the owner's: every comment on the collection, hidden ones too; answering one capture;
        // deleting the owner's own; hiding and showing others'. Editing is the page's.
        // ###########################################################################################
        public sealed record CommentsResult(Outcome Kind, IReadOnlyList<string> Errors, IReadOnlyList<CollectionCommentAnswer> Comments)
        {
            public string Message => string.Join(" ", this.Errors);
        }

        public async Task<CommentsResult> ListCommentsAsync(string serverId, string bearerToken, CancellationToken cancellationToken = default)
        {
            try
            {
                using HttpResponseMessage response = await this.SendAsync(HttpMethod.Get, CollectionRoutes.CommentsOf(Uri.EscapeDataString(serverId)), bearerToken, null, cancellationToken);
                string body = await response.Content.ReadAsStringAsync(cancellationToken);

                if (!response.IsSuccessStatusCode)
                {
                    Result refused = CollectionClient.ResultOfRefusal(response.StatusCode, body, wasShared: true, notFoundIsGone: false);
                    return new CommentsResult(refused.Kind, refused.Errors, []);
                }

                CollectionCommentsAnswer? answer = JsonSerializer.Deserialize<CollectionCommentsAnswer>(body, ReviewApiContract.WireSettings);

                return new CommentsResult(Outcome.Done, [], answer?.Comments ?? []);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
            {
                return new CommentsResult(Outcome.Failed, [CollectionClient.NotReachable], []);
            }
        }

        public Task<Result> AddCommentAsync(string serverId, string imageSha256, string body, string bearerToken, CancellationToken cancellationToken = default) =>
            this.CommentRequestAsync(HttpMethod.Post, CollectionRoutes.CommentsOf(Uri.EscapeDataString(serverId)), bearerToken, new CollectionCommentRequest(imageSha256, body), cancellationToken);

        public Task<Result> DeleteCommentAsync(string serverId, long commentId, string bearerToken, CancellationToken cancellationToken = default) =>
            this.CommentRequestAsync(HttpMethod.Delete, CollectionRoutes.CommentOf(Uri.EscapeDataString(serverId), commentId), bearerToken, null, cancellationToken);

        public Task<Result> SetCommentHiddenAsync(string serverId, long commentId, bool hidden, string bearerToken, CancellationToken cancellationToken = default) =>
            this.CommentRequestAsync(HttpMethod.Put, CollectionRoutes.CommentOf(Uri.EscapeDataString(serverId), commentId) + "/hidden", bearerToken, new CollectionCommentHideRequest(hidden), cancellationToken);

        // ###########################################################################################
        // The account's collections on the server, each with how many comments its page shows and the newest's id - what marks
        // a collection with comments not read yet. Null when it could not be asked (signed out, out of reach).
        // ###########################################################################################
        public async Task<IReadOnlyList<CollectionSummaryAnswer>?> ListMineAsync(string bearerToken, CancellationToken cancellationToken = default)
        {
            try
            {
                using HttpResponseMessage response = await this.SendAsync(HttpMethod.Get, CollectionRoutes.Mine, bearerToken, null, cancellationToken);

                if (!response.IsSuccessStatusCode)
                    return null;

                return JsonSerializer.Deserialize<MyCollectionsAnswer>(await response.Content.ReadAsStringAsync(cancellationToken), ReviewApiContract.WireSettings)?.Collections;
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
            {
                return null;
            }
        }

        private async Task<Result> CommentRequestAsync(HttpMethod method, string path, string bearerToken, object? content, CancellationToken cancellationToken)
        {
            try
            {
                using HttpResponseMessage response = await this.SendAsync(method, path, bearerToken, content, cancellationToken);

                return response.IsSuccessStatusCode
                    ? new Result(Outcome.Done, [])
                    : CollectionClient.ResultOfRefusal(response.StatusCode, await response.Content.ReadAsStringAsync(cancellationToken), wasShared: true, notFoundIsGone: false);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
            {
                return new Result(Outcome.Failed, [CollectionClient.NotReachable]);
            }
        }

        private Task<HttpResponseMessage> SendAsync(HttpMethod method, string path, string bearerToken, object? content, CancellationToken cancellationToken)
        {
            var message = new HttpRequestMessage(method, $"{this.thisRootUrl}{path}");

            if (content is not null)
                message.Content = JsonContent.Create(content, content.GetType(), options: ReviewApiContract.WireSettings);

            message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

            return this.thisHttp.SendAsync(message, cancellationToken);
        }

        // A first share (POST) carries the collection's local id as its key: sent again with its answer lost, the server changes
        // the page it made instead of making a second (CollectionRoutes.ClientKeyHeader; code review, 2026-11-27).
        private Task<HttpResponseMessage> SendContentAsync(string? serverId, string localId, CollectionContentRequest request, string bearerToken, CancellationToken cancellationToken)
        {
            var message = new HttpRequestMessage(
                serverId is null ? HttpMethod.Post : HttpMethod.Put,
                serverId is null ? $"{this.thisRootUrl}{CollectionRoutes.Collections}" : $"{this.thisRootUrl}{CollectionRoutes.Collections}/{Uri.EscapeDataString(serverId)}")
            {
                Content = JsonContent.Create(request, options: ReviewApiContract.WireSettings)
            };

            message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

            if (serverId is null && CollectionRules.IsClientKey(localId))
                message.Headers.Add(CollectionRoutes.ClientKeyHeader, localId);

            return this.thisHttp.SendAsync(message, cancellationToken);
        }

        // Each capture the server asked for - null when all went, else why one did not.
        private async Task<Result?> UploadAsync(
            IReadOnlyList<string> missing,
            Func<string, byte[]?> readCapture,
            string bearerToken,
            IProgress<string>? progress,
            CancellationToken cancellationToken)
        {
            for (int at = 0; at < missing.Count; at++)
            {
                string sha256 = missing[at];

                if (!DataTreeLayout.IsSha256(sha256) || readCapture(sha256) is not { } bytes)
                    return new Result(Outcome.Refused, ["A capture's file is no longer on this computer - take it out of the collection, or capture it again."]);

                progress?.Report($"Sending capture {at + 1} of {missing.Count}");

                using var message = new HttpRequestMessage(HttpMethod.Put, $"{this.thisRootUrl}{CollectionRoutes.Images}/{sha256}")
                {
                    Content = new ByteArrayContent(bytes)
                };

                message.Content.Headers.ContentType = new MediaTypeHeaderValue(CollectionPng.ContentType);
                message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);

                using HttpResponseMessage response = await this.thisHttp.SendAsync(message, cancellationToken);

                if (!response.IsSuccessStatusCode)
                    return CollectionClient.ResultOfRefusal(response.StatusCode, await response.Content.ReadAsStringAsync(cancellationToken), wasShared: false);
            }

            return null;
        }

        // ###########################################################################################
        // A refusal in words: the server's own sentences when it wrote any, an "update CRT" answer told to the rest of CRT.
        //
        // *** ON THE COMMENT ROUTES ONLY 410 SAYS THE PAGE IS GONE (`notFoundIsGone: false`; code review, 2026-11-27). *** A
        // comment's 404 is that COMMENT - deleted by its writer or the administrator meanwhile - and read as the page gone,
        // "Hide" on it made CRT forget a live page, whose next "Share" made a second one. The server answers 410 for its
        // owner's collection taken down (CollectionCommentFlows.ListForOwnerAsync).
        // ###########################################################################################
        private static Result ResultOfRefusal(HttpStatusCode status, string body, bool wasShared, bool notFoundIsGone = true)
        {
            if (status == HttpStatusCode.Unauthorized)
                return new Result(Outcome.SignInNeeded, [CollectionClient.SignInFirst]);

            if (SubmissionClient.SignalIfOutdated((int)status, body) is { } outdated)
                return new Result(Outcome.Failed, [outdated]);

            IReadOnlyList<string> errors = CollectionClient.RefusalOf(body).Errors is { Count: > 0 } said
                ? said
                : [SubmissionClient.RefusalMessage(body, $"The server refused it ({(int)status}).")];

            // A collection shared before that the server no longer has: deleted there - its time up, or taken down.
            if (wasShared && (status == HttpStatusCode.Gone || (notFoundIsGone && status == HttpStatusCode.NotFound)))
                return new Result(Outcome.GoneFromServer, errors);

            return new Result((int)status >= 500 ? Outcome.Failed : Outcome.Refused, errors);
        }

        internal static CollectionRefusalAnswer RefusalOf(string body)
        {
            try
            {
                return JsonSerializer.Deserialize<CollectionRefusalAnswer>(body, ReviewApiContract.WireSettings) ?? new CollectionRefusalAnswer([]);
            }
            catch (JsonException)
            {
                return new CollectionRefusalAnswer([]);
            }
        }
    }
}
