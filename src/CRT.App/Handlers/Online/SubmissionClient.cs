using CRT;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.Online
{
    // ###########################################################################################
    // Talks to CRT.Server's submission endpoints: hash the files, send the manifest, upload what
    // the server lacks, finalise.
    //
    // AN I/O BOUNDARY, deliberately untested, the same as OnlineServices' network half and
    // ScopeScpiClient. Every DECISION it makes lives elsewhere and is unit tested:
    // SubmissionManifestBuilder decides what goes in the manifest, SubmissionRetryPolicy decides
    // whether a failure is worth another attempt, SubmissionProgress decides what the user is
    // told. What is left here is HttpClient and SHA256.
    //
    // *** NOTHING HERE RUNS ON THE UI THREAD. *** Hashing a 76 MB system and uploading it are both
    // long operations, and "do not let submission block the UI" is a named trap in
    // NewContributeStrategy.md. Every method is async throughout and takes a CancellationToken
    // that is honoured, so Cancel means cancelled rather than "stop reporting progress".
    //
    // THE CAPABILITY TOKEN IS THE WHOLE AUTHORISATION MODEL. Contributing needs no account, so the
    // token returned by Create is what proves this client owns the submission. It is held in
    // memory for the duration and handed back to the caller so a submission can be resumed after a
    // restart - losing it means the submission can never be finished.
    // ###########################################################################################
    public sealed class SubmissionClient
    {
        // The header carrying the capability token. A header rather than a query parameter because
        // query strings are written to access logs by default on most web servers, and this value
        // authorises writing to a submission.
        private const string TokenHeader = "X-Submission-Token";

        // Chunked so a dropped connection loses at most this much, and so progress moves visibly
        // rather than jumping per file. 4 MB is large enough that per-request overhead is
        // irrelevant and small enough that a failure costs seconds, not minutes.
        private const int ChunkBytes = 4 * 1024 * 1024;

        private readonly string thisBaseUrl;

        public SubmissionClient(string? baseUrl = null)
        {
            this.thisBaseUrl = (baseUrl ?? AppConfig.CrtServerBaseUrl).TrimEnd('/');
        }

        // ###########################################################################################
        // Hashes every file the manifest will reference.
        //
        // STREAMED, never loaded whole: a board scan is tens of megabytes and the shipped C64
        // system is 76 MB across 1100 files. "Never load a whole system into memory" is stated
        // outright in the strategy document's performance note, and SHA256.HashDataAsync over a
        // FileStream is the cheap way to honour it.
        //
        // A file that cannot be read is REPORTED rather than throwing, because one unreadable file
        // out of 1100 should produce a message naming it, not a failed submission naming nothing.
        // ###########################################################################################
        public static async Task<FileHashResult> HashFilesAsync(
            string dataRoot,
            string draftSystemFolder,
            IReadOnlyList<string> relativePaths,
            IProgress<SubmissionProgress>? progress = null,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(relativePaths);

            var hashes = new Dictionary<string, SubmissionFileInfo>(StringComparer.Ordinal);
            var problems = new List<string>();

            int done = 0;

            foreach (string relative in relativePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                progress?.Report(new SubmissionProgress(
                    SubmissionPhase.Hashing, done, relativePaths.Count, 0, 0, relative));

                // TWO ROOTS, not one - see SubmissionFileLocator. A draft over a published system
                // references officially published files (under Data/) and drafted ones (under the
                // draft's own Files/ folder) in the same submission, so resolving against a single
                // root reports most of a real system as missing.
                //
                // The path came from the app's own board data rather than over the network, but it
                // is still checked: a hand-edited spreadsheet is untrusted input too, and this is
                // the one place the app turns a stored string into a file read.
                if (!SubmissionFileLocator.TryLocate(
                        dataRoot, draftSystemFolder, relative, out string absolute, out string reason))
                {
                    problems.Add($"[{relative}] cannot be used: {reason}");
                    done++;
                    continue;
                }

                try
                {
                    var info = new FileInfo(absolute);

                    await using FileStream stream = new(
                        absolute, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true);

                    byte[] hash = await SHA256.HashDataAsync(stream, cancellationToken);

                    hashes[relative] = new SubmissionFileInfo(Convert.ToHexStringLower(hash), info.Length);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    problems.Add($"[{relative}] could not be read: {ex.Message}");
                }

                done++;
            }

            return new FileHashResult(hashes, problems);
        }

        // ###########################################################################################
        // Step 1 and 2: send the manifest, get back the token and what to upload.
        // ###########################################################################################
        public async Task<HashNegotiationResponse> CreateAsync(
            SubmissionManifest manifest,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            using HttpClient http = SubmissionClient.CreateHttpClient(AppConfig.ApiTimeout);

            using HttpResponseMessage response = await http.PostAsJsonAsync(
                $"{this.thisBaseUrl}/submissions", manifest, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                string body = await response.Content.ReadAsStringAsync(cancellationToken);

                throw new SubmissionRejectedException(
                    (int)response.StatusCode,
                    SubmissionClient.ExtractFindings(body),
                    $"The server refused the submission ({(int)response.StatusCode}).");
            }

            HashNegotiationResponse? negotiation =
                await response.Content.ReadFromJsonAsync<HashNegotiationResponse>(cancellationToken);

            return negotiation
                ?? throw new SubmissionRejectedException(0, [], "The server's reply could not be read.");
        }

        // ###########################################################################################
        // Step 3: upload one blob, resuming from wherever the server says it already has.
        //
        // ASKS THE SERVER WHERE TO RESUME rather than assuming zero. That is what makes resumption
        // survive a client restart and not merely a retry within one session - the offset lives on
        // the server, where it cannot disagree with the bytes on disk.
        // ###########################################################################################
        public async Task UploadBlobAsync(
            long submissionId,
            string uploadToken,
            string hash,
            string absolutePath,
            IProgress<SubmissionProgress>? progress = null,
            SubmissionProgress? overall = null,
            CancellationToken cancellationToken = default)
        {
            using HttpClient http = SubmissionClient.CreateHttpClient(AppConfig.SubmissionUploadTimeout);

            long offset = await this.GetResumeOffsetAsync(http, submissionId, uploadToken, hash, cancellationToken);

            await using FileStream source = new(
                absolutePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true);

            if (offset > 0 && offset <= source.Length)
                source.Seek(offset, SeekOrigin.Begin);

            byte[] buffer = new byte[SubmissionClient.ChunkBytes];

            while (offset < source.Length)
            {
                cancellationToken.ThrowIfCancellationRequested();

                int read = await source.ReadAsync(buffer, cancellationToken);

                if (read <= 0)
                    break;

                int attempts = 0;

                while (true)
                {
                    try
                    {
                        using var content = new ByteArrayContent(buffer, 0, read);

                        using var request = new HttpRequestMessage(
                            HttpMethod.Put,
                            $"{this.thisBaseUrl}/submissions/{submissionId}/blobs/{hash}?offset={offset}")
                        {
                            Content = content
                        };

                        request.Headers.Add(SubmissionClient.TokenHeader, uploadToken);

                        using HttpResponseMessage response = await http.SendAsync(request, cancellationToken);

                        if (response.IsSuccessStatusCode)
                            break;

                        // 416 means the two sides disagree about how much arrived. Re-ask and
                        // continue from there rather than retrying the same wrong offset - this is
                        // a RESUME, not a failure, which is why SubmissionRetryPolicy excludes it.
                        if ((int)response.StatusCode == 416)
                        {
                            offset = await this.GetResumeOffsetAsync(
                                http, submissionId, uploadToken, hash, cancellationToken);

                            source.Seek(offset, SeekOrigin.Begin);
                            read = 0;
                            break;
                        }

                        if (!SubmissionRetryPolicy.ShouldRetry((int)response.StatusCode, attempts))
                        {
                            string body = await response.Content.ReadAsStringAsync(cancellationToken);

                            throw new SubmissionRejectedException(
                                (int)response.StatusCode, [],
                                $"Uploading [{Path.GetFileName(absolutePath)}] failed: {body}");
                        }

                        attempts++;
                        await Task.Delay(SubmissionRetryPolicy.DelayBefore(attempts), cancellationToken);
                    }
                    catch (HttpRequestException) when (SubmissionRetryPolicy.ShouldRetry(null, attempts))
                    {
                        attempts++;
                        await Task.Delay(SubmissionRetryPolicy.DelayBefore(attempts), cancellationToken);
                    }
                }

                offset += read;

                if (progress is not null && overall is not null)
                {
                    progress.Report(overall with
                    {
                        BytesDone = overall.BytesDone + offset,
                        CurrentFile = Path.GetFileName(absolutePath)
                    });
                }
            }
        }

        // ###########################################################################################
        // Finalise. The answer says whether it was queued for review or rejected, and why.
        // ###########################################################################################
        public async Task<SubmissionResult> FinaliseAsync(
            long submissionId,
            string uploadToken,
            CancellationToken cancellationToken = default)
        {
            using HttpClient http = SubmissionClient.CreateHttpClient(AppConfig.ApiTimeout);

            using var request = new HttpRequestMessage(
                HttpMethod.Post, $"{this.thisBaseUrl}/submissions/{submissionId}/finalise");

            request.Headers.Add(SubmissionClient.TokenHeader, uploadToken);

            using HttpResponseMessage response = await http.SendAsync(request, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                string body = await response.Content.ReadAsStringAsync(cancellationToken);

                throw new SubmissionRejectedException(
                    (int)response.StatusCode,
                    SubmissionClient.ExtractFindings(body),
                    $"The server could not finish the submission ({(int)response.StatusCode}).");
            }

            SubmissionResult? result =
                await response.Content.ReadFromJsonAsync<SubmissionResult>(cancellationToken);

            return result
                ?? throw new SubmissionRejectedException(0, [], "The server's reply could not be read.");
        }

        // ###########################################################################################
        // What the server currently says about one already-sent submission (Phase 4, task 6).
        //
        // Returns null rather than throwing when the answer cannot be had - offline, a timeout, a
        // server that has forgotten the submission, or a token that no longer works. The "my
        // submissions" view calls this once per row, and one unreachable row must not fail the
        // whole refresh; a null simply leaves that row showing its last known state, which is
        // exactly what the cached value in the receipt is for.
        //
        // A 404 is NOT an error worth reporting either. It is what the server answers both for "no
        // such submission" and for "not yours" (deliberately indistinguishable, so the id space
        // cannot be walked), and an expired or pruned submission is an ordinary end state rather
        // than a fault.
        // ###########################################################################################
        public async Task<SubmissionStatus?> GetStatusAsync(
            long submissionId,
            string uploadToken,
            CancellationToken cancellationToken = default)
        {
            try
            {
                using HttpClient http = SubmissionClient.CreateHttpClient(AppConfig.ApiTimeout);

                using var request = new HttpRequestMessage(
                    HttpMethod.Get, $"{this.thisBaseUrl}/submissions/{submissionId}");

                request.Headers.Add(SubmissionClient.TokenHeader, uploadToken);

                using HttpResponseMessage response = await http.SendAsync(request, cancellationToken);

                if (!response.IsSuccessStatusCode)
                    return null;

                return await response.Content.ReadFromJsonAsync<SubmissionStatus>(cancellationToken);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                // The CALLER cancelled (the window closed, or a newer refresh superseded this one).
                // That is not a failed lookup and must not be reported as one, so it propagates -
                // swallowing it here would make a closing window look like an offline server.
                // Checked against the token rather than caught blind, because HttpClient also
                // raises this for its own TIMEOUT, which IS a failed lookup and falls through to
                // the handler below.
                throw;
            }
            catch (Exception ex) when (ex is HttpRequestException or IOException or JsonException or TaskCanceledException)
            {
                // Deliberately quiet: this runs per row on a view the user opened to read, and a
                // log line per unreachable submission every time they open it is noise, not
                // diagnosis. The row says "last checked" and the user can see it did not move.
                return null;
            }
        }

        // ###########################################################################################
        // How much of this blob the server already holds.
        //
        // A failure here is answered with 0 rather than thrown: the worst case is re-uploading
        // bytes the server already had, which is wasteful but correct. Failing the whole submission
        // because a progress query did not answer would be the wrong trade.
        // ###########################################################################################
        private async Task<long> GetResumeOffsetAsync(
            HttpClient http,
            long submissionId,
            string uploadToken,
            string hash,
            CancellationToken cancellationToken)
        {
            try
            {
                using var request = new HttpRequestMessage(
                    HttpMethod.Get, $"{this.thisBaseUrl}/submissions/{submissionId}/blobs/{hash}");

                request.Headers.Add(SubmissionClient.TokenHeader, uploadToken);

                using HttpResponseMessage response = await http.SendAsync(request, cancellationToken);

                if (!response.IsSuccessStatusCode)
                    return 0;

                using JsonDocument document = JsonDocument.Parse(
                    await response.Content.ReadAsStringAsync(cancellationToken));

                return document.RootElement.TryGetProperty("uploaded", out JsonElement uploaded)
                    ? uploaded.GetInt64()
                    : 0;
            }
            catch (Exception ex) when (ex is HttpRequestException or JsonException or TaskCanceledException)
            {
                return 0;
            }
        }

        private static HttpClient CreateHttpClient(TimeSpan timeout)
        {
            return new HttpClient { Timeout = timeout };
        }

        // The server answers a refusal with { "errors": [ ... ] }. Pulling the findings out lets
        // the UI show what is actually wrong rather than a status code.
        private static IReadOnlyList<ValidationFinding> ExtractFindings(string body)
        {
            try
            {
                using JsonDocument document = JsonDocument.Parse(body);

                if (!document.RootElement.TryGetProperty("errors", out JsonElement errors))
                    return [];

                return JsonSerializer.Deserialize<List<ValidationFinding>>(
                    errors.GetRawText(),
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];
            }
            catch (JsonException)
            {
                return [];
            }
        }
    }

    // ###########################################################################################
    // The hashes, plus anything that could not be hashed. Problems name the file, because one
    // unreadable file out of 1100 must produce a message the contributor can act on.
    // ###########################################################################################
    public sealed record FileHashResult(
        IReadOnlyDictionary<string, SubmissionFileInfo> Hashes,
        IReadOnlyList<string> Problems);

    // ###########################################################################################
    // The server refused something. Findings are present when the refusal was about the data
    // rather than the request.
    // ###########################################################################################
    public sealed class SubmissionRejectedException : Exception
    {
        public SubmissionRejectedException(
            int statusCode,
            IReadOnlyList<ValidationFinding> findings,
            string message)
            : base(message)
        {
            this.StatusCode = statusCode;
            this.Findings = findings;
        }

        public int StatusCode { get; }

        public IReadOnlyList<ValidationFinding> Findings { get; }
    }
}
