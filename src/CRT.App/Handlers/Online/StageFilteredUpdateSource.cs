using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Velopack;
using Velopack.Logging;
using Velopack.Sources;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // Wraps an update source and drops every release whose stage the user has not opted into,
    // BEFORE Velopack ranks what is left and picks the newest.
    //
    // The filtering has to happen here rather than on the single candidate Velopack hands back.
    // GithubSource's "prerelease" flag is all-or-nothing, so with only "notify for BETA" ticked
    // Velopack still ranks the alphas alongside the betas and returns whichever is newest overall.
    // Checking the stage afterwards and discarding the result meant a newer ALPHA permanently HID
    // an available BETA: every subsequent check returned that same alpha, was discarded again, and
    // the user was told they were up to date indefinitely. Filtering the feed instead leaves the
    // newest BETA as the top candidate, which is exactly what that user asked for.
    // ###########################################################################################
    internal sealed class StageFilteredUpdateSource : IUpdateSource
    {
        private readonly IUpdateSource thisInner;
        private readonly bool thisAllowAlpha;
        private readonly bool thisAllowBeta;

        public StageFilteredUpdateSource(IUpdateSource inner, bool allowAlpha, bool allowBeta)
        {
            this.thisInner = inner;
            this.thisAllowAlpha = allowAlpha;
            this.thisAllowBeta = allowBeta;
        }

        public async Task<VelopackAssetFeed> GetReleaseFeed(
            IVelopackLogger logger,
            string? appId,
            string channel,
            Guid? stagingId = null,
            VelopackAsset? latestLocalRelease = null)
        {
            VelopackAssetFeed feed = await this.thisInner
                .GetReleaseFeed(logger, appId, channel, stagingId, latestLocalRelease)
                .ConfigureAwait(false);

            if (feed?.Assets == null)
            {
                return feed!;
            }

            VelopackAsset[] allowed = feed.Assets
                .Where(asset => UpdateChannelFilter.IsVersionAllowed(
                    asset?.Version?.ToString(),
                    this.thisAllowAlpha,
                    this.thisAllowBeta))
                .ToArray();

            return new VelopackAssetFeed { Assets = allowed };
        }

        public Task DownloadReleaseEntry(
            IVelopackLogger logger,
            VelopackAsset releaseEntry,
            string localFile,
            Action<int> progress,
            CancellationToken cancelToken = default)
        {
            return this.thisInner.DownloadReleaseEntry(logger, releaseEntry, localFile, progress, cancelToken);
        }
    }
}
