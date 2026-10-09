using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHERE A BOARD'S TABLE READS A FILE CELL'S FILE (2026-10-03) - the Boards screen's Board data
    // view. Every file a row may cite there is already in BETA (an edit cannot bring in a new one -
    // the server's BoardEditFlow refuses it), so BOTH sides are read from BETA's public address,
    // exactly as CRT downloads them and as the file tree opens them (FileTreeFiles): the side the
    // table was opened on, and the file a changed cell names now.
    //
    // An I/O boundary, so it is thin, like ReviewTableFileSource. Each file is fetched once per
    // table; a failed fetch is not remembered, so it is tried again.
    // ###########################################################################################
    internal sealed class PublishedTableFileSource : CRT.IBoardTableFileSource
    {
        private readonly ReviewApiClient thisClient;
        private readonly string? thisBetaDataUrl;
        private readonly string thisTreeName;
        private readonly Func<string, Task<bool>> thisLaunch;
        private readonly ConcurrentDictionary<string, Task<byte[]?>> thisFetches = new(StringComparer.Ordinal);

        // `dataUrl` is where the table's tree is published - BETA's, or the stable source's for its
        // read-only table (2026-10-04), named `treeName` when a file is not there.
        public PublishedTableFileSource(ReviewApiClient client, string? dataUrl, Func<string, Task<bool>> launch, string treeName = "BETA")
        {
            this.thisClient = client ?? throw new ArgumentNullException(nameof(client));
            this.thisBetaDataUrl = dataUrl;
            this.thisTreeName = string.IsNullOrWhiteSpace(treeName) ? "BETA" : treeName;
            this.thisLaunch = launch ?? throw new ArgumentNullException(nameof(launch));
        }

        // The two sides of a file cell the change replaced: BETA's, and the maintainer's.
        public string PublishedLabel => BoardSections.BaselineLabel;

        public string CurrentLabel => BoardSections.ChangeLabel;

        public async Task<byte[]?> ReadAsync(string path, CRT.BoardTableFileSide side)
        {
            Task<byte[]?> fetch = this.thisFetches.GetOrAdd(path ?? string.Empty, this.FetchAsync);
            byte[]? bytes = await fetch;

            if (bytes is null)
                this.thisFetches.TryRemove(new KeyValuePair<string, Task<byte[]?>>(path ?? string.Empty, fetch));

            return bytes;
        }

        // Saved into a folder of its own under the temp folder and handed to the operating system,
        // as the submission's table does (OpenedFiles).
        public async Task<string?> OpenAsync(string path, CRT.BoardTableFileSide side)
        {
            if (!ReviewTableFiles.TryGetOpenName(path, out string fileName))
                return "This type of file is not opened from here.";

            byte[]? bytes = await this.ReadAsync(path, side);

            if (bytes is null)
                return $"There is no file at this path in {this.thisTreeName}.";

            return await OpenedFiles.SaveAndLaunchAsync(bytes, fileName, this.thisLaunch);
        }

        private async Task<byte[]?> FetchAsync(string path)
        {
            try
            {
                ReviewApiResult<byte[]> result = await this.thisClient.GetPublishedDataFileAsync(this.thisBetaDataUrl, path);
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
