using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // WHERE A FILE TREE'S FILES COME FROM (owner request, 2026-09-28: "I need to check the Excel
    // file - does its format look correct, and JSON etc."). For every maintainer, not only whoever
    // can reach the data from their own computer: the bytes come from where the tree entry says -
    //
    //   - BETA or production: the file's public address, exactly as CRT downloads it;
    //   - the submission: its upload, by hash, through the review API.
    //
    // Read for the hover card as it is, and fetched under the window's "please wait" to be opened
    // (a board scan is megabytes), then saved and handed to the operating system (OpenedFiles).
    //
    // *** EACH FILE IS FETCHED ONCE PER TREE SHOWN. *** Hovering back and forth over a board scan
    // would otherwise download it every time, and opening one just looked at would fetch it again.
    // A failed fetch is not remembered, so it is tried again. One of these per tree shown, so a
    // file changed since is fetched anew the next time the tree is.
    // ###########################################################################################
    internal sealed class FileTreeFiles : IFileTreeFiles
    {
        private readonly Control thisAnchor;
        private readonly ReviewApiClient thisClient;
        private readonly ReviewSession? thisSession;
        private readonly long? thisSubmissionId;
        private readonly string? thisBetaDataUrl;
        private readonly string? thisProductionDataUrl;
        private readonly ConcurrentDictionary<string, byte[]> thisFetched = new(StringComparer.Ordinal);

        // `anchor` is where the wait is shown and what launches the file; `submissionId` and the
        // session reach a submission's own files, the two addresses the published ones.
        public FileTreeFiles(
            Control anchor,
            ReviewApiClient client,
            ReviewSession? session,
            long? submissionId,
            string? betaDataUrl,
            string? productionDataUrl)
        {
            this.thisAnchor = anchor ?? throw new ArgumentNullException(nameof(anchor));
            this.thisClient = client ?? throw new ArgumentNullException(nameof(client));
            this.thisSession = session;
            this.thisSubmissionId = submissionId;
            this.thisBetaDataUrl = betaDataUrl;
            this.thisProductionDataUrl = productionDataUrl;
        }

        public async Task<byte[]?> ReadAsync(BoardFileEntry file)
        {
            try
            {
                ReviewApiResult<byte[]> result = await this.FetchAsync(file, CancellationToken.None);
                return result.IsOk ? result.Value : null;
            }
            catch (Exception)
            {
                // The card promises never to throw; a network fault is a missing picture.
                return null;
            }
        }

        public async Task<string?> OpenAsync(BoardFileEntry file)
        {
            ArgumentNullException.ThrowIfNull(file);

            if (file.OpenFrom == BoardFileSource.NotWrittenYet)
                return FileTreeWording.Note(file);

            if (!OpenedFiles.TryGetOpenName(file.Path, out string fileName))
                return "This type of file is not opened from here.";

            ReviewApiResult<byte[]> bytes = await ServerWait.CallAsync(
                this.thisAnchor, MaintainerWaitWording.OpeningFile(fileName), token => this.FetchAsync(file, token));

            if (!bytes.IsOk)
                return bytes.Failure == ReviewApiFailure.NotFound ? "The file is not there any more." : bytes.Message;

            return await OpenedFiles.SaveAndLaunchAsync(bytes.Value!, fileName, this.LaunchAsync);
        }

        private async Task<ReviewApiResult<byte[]>> FetchAsync(BoardFileEntry file, CancellationToken token)
        {
            string key = $"{file.OpenFrom}|{file.Sha256}|{file.Path}";

            if (this.thisFetched.TryGetValue(key, out byte[]? held))
                return ReviewApiResult<byte[]>.Ok(held);

            ReviewApiResult<byte[]> result = await (file.OpenFrom switch
            {
                BoardFileSource.Submission when this.thisSession is not null && this.thisSubmissionId is long id && file.Sha256 is string hash =>
                    this.thisClient.GetSubmittedAssetAsync(this.thisSession, id, hash, token),
                BoardFileSource.Production => this.thisClient.GetPublishedDataFileAsync(this.thisProductionDataUrl, file.Path, token),
                BoardFileSource.Beta => this.thisClient.GetPublishedDataFileAsync(this.thisBetaDataUrl, file.Path, token),
                _ => Task.FromResult(ReviewApiResult<byte[]>.Failed(ReviewApiFailure.NotFound, "There is nothing to open yet."))
            });

            if (result.IsOk && result.Value is byte[] bytes)
                this.thisFetched[key] = bytes;

            return result;
        }

        // A platform with no program for the type answers false; the tree says it could not open.
        private Task<bool> LaunchAsync(string fullPath) => MaintainerFileLauncher.LaunchAsync(this.thisAnchor, fullPath);
    }
}
