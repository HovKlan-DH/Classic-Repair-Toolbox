using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHERE THE MAINTAINER'S TABLE READS A FILE CELL'S FILE for its hover card (owner request,
    // 2026-09-26) - see Controls/BoardTable/BoardTableEditor.FilePreview.cs. The same two routes the change
    // summary's pictures use: the published file by path, the submitted one by its hash.
    //
    // An I/O boundary, so it is thin: which hash a path means and what a file is saved as before
    // it is opened are ReviewTableFiles', tested there.
    //
    // *** EACH FILE IS FETCHED ONCE PER TABLE. *** Hovering back and forth over a board scan would
    // otherwise download it every time. A failed fetch is not remembered, so it is tried again.
    // The current side is keyed by the submitted HASH, so a file changed by a save is fetched anew.
    // ###########################################################################################
    internal sealed class ReviewTableFileSource : CRT.IBoardTableFileSource
    {
        private readonly ReviewApiClient thisClient;
        private readonly ReviewSession thisSession;
        private readonly long thisSubmissionId;
        private readonly Func<IReadOnlyList<SubmittedFileFact>> thisSubmittedFiles;
        private readonly Func<bool> thisNothingPublished;
        private readonly Func<string, Task<bool>> thisLaunch;
        private readonly ConcurrentDictionary<string, Task<byte[]?>> thisFetches = new(StringComparer.Ordinal);

        // `submittedFiles` is read at each request, so it follows the submission as a save changes
        // it. `nothingPublished` says whether the table was opened on a NEW system, whose
        // "published" side is the submission itself (ReviewTableFiles.HashToRead). `launch` hands a
        // local file to the operating system, and says whether it could.
        public ReviewTableFileSource(
            ReviewApiClient client,
            ReviewSession session,
            long submissionId,
            Func<IReadOnlyList<SubmittedFileFact>> submittedFiles,
            Func<bool> nothingPublished,
            Func<string, Task<bool>> launch)
        {
            this.thisClient = client ?? throw new ArgumentNullException(nameof(client));
            this.thisSession = session ?? throw new ArgumentNullException(nameof(session));
            this.thisSubmissionId = submissionId;
            this.thisSubmittedFiles = submittedFiles ?? throw new ArgumentNullException(nameof(submittedFiles));
            this.thisNothingPublished = nothingPublished ?? throw new ArgumentNullException(nameof(nothingPublished));
            this.thisLaunch = launch ?? throw new ArgumentNullException(nameof(launch));
        }

        public string PublishedLabel => ReviewTableFiles.SideLabels(this.thisNothingPublished()).Published;

        public string CurrentLabel => ReviewTableFiles.SideLabels(this.thisNothingPublished()).Current;

        // A new system's file is the submission's on both sides, so "Unchanged" would say nothing.
        public bool SaysUnchanged => !this.thisNothingPublished();

        public async Task<byte[]?> ReadAsync(string path, CRT.BoardTableFileSide side)
        {
            string? hash = ReviewTableFiles.HashToRead(this.thisSubmittedFiles(), path, side, this.thisNothingPublished());

            string key = hash is not null ? "submitted|" + hash : "published|" + path;

            Task<byte[]?> fetch = this.thisFetches.GetOrAdd(key, _ => this.FetchAsync(path, hash));
            byte[]? bytes = await fetch;

            if (bytes is null)
                this.thisFetches.TryRemove(new KeyValuePair<string, Task<byte[]?>>(key, fetch));

            return bytes;
        }

        // ###########################################################################################
        // Saves the file into a folder of its own under the temp folder and hands it to the operating
        // system - a PDF opens in the PDF viewer. The name is ReviewTableFiles.TryGetOpenName's: only
        // a type a submission may carry, and a web page as text.
        // ###########################################################################################
        public async Task<string?> OpenAsync(string path, CRT.BoardTableFileSide side)
        {
            if (!ReviewTableFiles.TryGetOpenName(path, out string fileName))
                return "This type of file is not opened from here.";

            byte[]? bytes = await this.ReadAsync(path, side);

            if (bytes is null)
                return "There is no file at this path.";

            // Saved and handed to the operating system - shared with the file tree (OpenedFiles).
            return await OpenedFiles.SaveAndLaunchAsync(bytes, fileName, this.thisLaunch);
        }

        // The submitted file by its hash when the submission carries one at this path, otherwise the
        // published file at the path. Null for anything the server does not answer with bytes.
        private async Task<byte[]?> FetchAsync(string path, string? hash)
        {
            try
            {
                ReviewApiResult<byte[]> result = hash is not null
                    ? await this.thisClient.GetSubmittedAssetAsync(this.thisSession, this.thisSubmissionId, hash)
                    : await this.thisClient.GetPublishedAssetAsync(this.thisSession, this.thisSubmissionId, path);

                return result.IsOk ? result.Value : null;
            }
            catch (Exception)
            {
                // The source promises the card never throws; a network fault is a missing picture.
                return null;
            }
        }
    }
}
