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
    // *** NO ACCOUNT (owner decision, 2026-11-28: "Anybody, no account"). *** Every request carries the collection's own KEY
    // (LocalCollection.OwnerKey, CollectionRoutes.OwnerKeyHeader) - the server knows its page by it, and nothing else may change
    // it. Answers a result, never throws for a refusal or a server out of reach: the collection window says it in words.
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

        public const string NoKey = "This collection has no key to share it with - close its window and open it again.";

        // ###########################################################################################
        // SHARES THE COLLECTION AS IT IS NOW, under its key: sent whole - a new page while it has none (POST), else its page
        // changed (PUT); the captures the server does not hold yet sent then, each from `readCapture`; and the collection again.
        // `progress` hears each capture as it goes. A first share whose answer was lost, sent again under the same key, changes
        // the page it made (CollectionFlows.CreateAsync).
        // ###########################################################################################
        public async Task<Result> ShareAsync(
            LocalCollection collection,
            Func<string, byte[]?> readCapture,
            DateTimeOffset now,
            IProgress<string>? progress = null,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(collection);

            if (collection.OwnerKey is not { } key || !CollectionRules.IsOwnerKey(key))
                return new Result(Outcome.Refused, [CollectionClient.NoKey]);

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

                    using HttpResponseMessage response = await this.SendContentAsync(collection.ServerId, request, key, cancellationToken);
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
                        if (await this.UploadAsync(missing, readCapture, key, progress, cancellationToken) is { } failed)
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
        // Takes the collection's page off the server. One the server deleted already - or does not know under this key - is
        // deleted.
        // ###########################################################################################
        public async Task<Result> DeleteAsync(string serverId, string ownerKey, CancellationToken cancellationToken = default)
        {
            try
            {
                using HttpResponseMessage response = await this.SendAsync(
                    HttpMethod.Delete, $"{CollectionRoutes.Collections}/{Uri.EscapeDataString(serverId)}", ownerKey, null, cancellationToken);

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
        // page offline or back, under the same link, its captures kept on the server either way. Done, or what came of it; a page
        // the server no longer has is GoneFromServer.
        // ###########################################################################################
        public Task<Result> SetSharingAsync(string serverId, bool shared, string ownerKey, CancellationToken cancellationToken = default) =>
            this.RequestAsync(HttpMethod.Put, CollectionRoutes.SharingOf(Uri.EscapeDataString(serverId)), ownerKey, new CollectionSharingRequest(shared), cancellationToken);

        // ###########################################################################################
        // "DELETE IMAGE" on a shared collection (owner request, 2026-11-28: "should also remove it right away from the shared
        // page, and do not touch the expiry date"; server 16.1.0): that one capture off its page now, nothing else sent - its
        // deletion date, its sharing and any change not shared yet stay as they are. Done; a page the server no longer has is
        // GoneFromServer.
        // ###########################################################################################
        public Task<Result> RemoveImageAsync(string serverId, string sha256, string ownerKey, CancellationToken cancellationToken = default) =>
            this.RequestAsync(
                HttpMethod.Delete,
                CollectionRoutes.ImageOf(Uri.EscapeDataString(serverId), Uri.EscapeDataString(sha256)),
                ownerKey,
                null,
                cancellationToken);

        private async Task<Result> RequestAsync(HttpMethod method, string path, string ownerKey, object? content, CancellationToken cancellationToken)
        {
            try
            {
                using HttpResponseMessage response = await this.SendAsync(method, path, ownerKey, content, cancellationToken);

                return response.IsSuccessStatusCode
                    ? new Result(Outcome.Done, [])
                    : CollectionClient.ResultOfRefusal(response.StatusCode, await response.Content.ReadAsStringAsync(cancellationToken), wasShared: true);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
            {
                return new Result(Outcome.Failed, [CollectionClient.NotReachable]);
            }
        }

        private Task<HttpResponseMessage> SendAsync(HttpMethod method, string path, string ownerKey, object? content, CancellationToken cancellationToken)
        {
            var message = new HttpRequestMessage(method, $"{this.thisRootUrl}{path}");

            if (content is not null)
                message.Content = JsonContent.Create(content, content.GetType(), options: ReviewApiContract.WireSettings);

            message.Headers.Add(CollectionRoutes.OwnerKeyHeader, ownerKey);

            return this.thisHttp.SendAsync(message, cancellationToken);
        }

        // The whole collection: POST while it has no page, PUT to its page after.
        private Task<HttpResponseMessage> SendContentAsync(string? serverId, CollectionContentRequest request, string ownerKey, CancellationToken cancellationToken)
        {
            var message = new HttpRequestMessage(
                serverId is null ? HttpMethod.Post : HttpMethod.Put,
                serverId is null ? $"{this.thisRootUrl}{CollectionRoutes.Collections}" : $"{this.thisRootUrl}{CollectionRoutes.Collections}/{Uri.EscapeDataString(serverId)}")
            {
                Content = JsonContent.Create(request, options: ReviewApiContract.WireSettings)
            };

            message.Headers.Add(CollectionRoutes.OwnerKeyHeader, ownerKey);

            return this.thisHttp.SendAsync(message, cancellationToken);
        }

        // Each capture the server asked for - null when all went, else why one did not.
        private async Task<Result?> UploadAsync(
            IReadOnlyList<string> missing,
            Func<string, byte[]?> readCapture,
            string ownerKey,
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
                message.Headers.Add(CollectionRoutes.OwnerKeyHeader, ownerKey);

                using HttpResponseMessage response = await this.thisHttp.SendAsync(message, cancellationToken);

                if (!response.IsSuccessStatusCode)
                    return CollectionClient.ResultOfRefusal(response.StatusCode, await response.Content.ReadAsStringAsync(cancellationToken), wasShared: false);
            }

            return null;
        }

        // ###########################################################################################
        // A refusal in words: the server's own sentences when it wrote any, an "update CRT" answer told to the rest of CRT. A
        // page shared before that the server no longer has - its time up, taken down, or (404) under this key no more - is gone.
        // ###########################################################################################
        private static Result ResultOfRefusal(HttpStatusCode status, string body, bool wasShared)
        {
            if (SubmissionClient.SignalIfOutdated((int)status, body) is { } outdated)
                return new Result(Outcome.Failed, [outdated]);

            IReadOnlyList<string> errors = CollectionClient.RefusalOf(body).Errors is { Count: > 0 } said
                ? said
                : [SubmissionClient.RefusalMessage(body, $"The server refused it ({(int)status}).")];

            if (wasShared && status is HttpStatusCode.Gone or HttpStatusCode.NotFound)
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
