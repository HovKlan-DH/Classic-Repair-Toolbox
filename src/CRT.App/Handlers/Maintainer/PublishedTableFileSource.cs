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
    // *** COMPARED WITH THE OTHER SOURCE (owner request, 2026-10-09 - "Compare sources"), the side
    // the table is coloured against is the OTHER source's, *** so it is read from that source's
    // address (`Baseline`), and each picture is headed by the source it is from.
    //
    // An I/O boundary, so it is thin, like ReviewTableFileSource. Each file is fetched once per
    // table and side; a failed fetch is not remembered, so it is tried again.
    // ###########################################################################################
    internal sealed class PublishedTableFileSource : CRT.IBoardTableFileSource
    {
        private readonly ReviewApiClient thisClient;
        private readonly string? thisDataUrl;
        private readonly string thisTreeName;
        private readonly Func<string, Task<bool>> thisLaunch;
        private readonly ConcurrentDictionary<(string? Url, string Path), Task<byte[]?>> thisFetches = new();

        // `dataUrl` is where the table's tree is published - BETA's, or the stable source's for its
        // read-only table (2026-10-04), named `treeName` when a file is not there.
        public PublishedTableFileSource(ReviewApiClient client, string? dataUrl, Func<string, Task<bool>> launch, string treeName = BoardSections.BetaTreeName)
        {
            this.thisClient = client ?? throw new ArgumentNullException(nameof(client));
            this.thisDataUrl = dataUrl;
            this.thisTreeName = string.IsNullOrWhiteSpace(treeName) ? BoardSections.BetaTreeName : treeName;
            this.thisLaunch = launch ?? throw new ArgumentNullException(nameof(launch));
        }

        // ###########################################################################################
        // The source the table is coloured against, when it is the OTHER one: where its files are
        // published, what it is called when a file is not there, and the two pictures' headings.
        // Null - the ordinary table - reads both sides from `dataUrl`.
        // ###########################################################################################
        public PublishedTableBaseline? Baseline { get; init; }

        // The two sides of a file cell the change replaced: BETA's, and the maintainer's - or, compared,
        // the two sources.
        public string PublishedLabel => this.Baseline?.PublishedLabel ?? BoardSections.BaselineLabel;

        public string CurrentLabel => this.Baseline?.CurrentLabel ?? BoardSections.ChangeLabel;

        public async Task<byte[]?> ReadAsync(string path, CRT.BoardTableFileSide side)
        {
            var key = (this.UrlFor(side), path ?? string.Empty);
            Task<byte[]?> fetch = this.thisFetches.GetOrAdd(key, this.FetchAsync);
            byte[]? bytes = await fetch;

            if (bytes is null)
                this.thisFetches.TryRemove(new KeyValuePair<(string?, string), Task<byte[]?>>(key, fetch));

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
            {
                string tree = side == CRT.BoardTableFileSide.Published && this.Baseline is { } baseline
                    ? baseline.TreeName
                    : this.thisTreeName;

                return $"There is no file at this path in {tree}.";
            }

            return await OpenedFiles.SaveAndLaunchAsync(bytes, fileName, this.thisLaunch);
        }

        // Where a side's file is published - the other source's for the side compared against.
        internal string? UrlFor(CRT.BoardTableFileSide side) =>
            side == CRT.BoardTableFileSide.Published && this.Baseline is { } baseline ? baseline.DataUrl : this.thisDataUrl;

        private async Task<byte[]?> FetchAsync((string? Url, string Path) file)
        {
            try
            {
                ReviewApiResult<byte[]> result = await this.thisClient.GetPublishedDataFileAsync(file.Url, file.Path);
                return result.IsOk ? result.Value : null;
            }
            catch (Exception)
            {
                // The source promises the card never throws; a network fault is a missing picture.
                return null;
            }
        }
    }

    // The other source a compared table is coloured against - see PublishedTableFileSource.Baseline.
    internal sealed record PublishedTableBaseline(string? DataUrl, string TreeName, string PublishedLabel, string CurrentLabel);
}
