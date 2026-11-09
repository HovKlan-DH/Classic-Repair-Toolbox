using System.Text.Json;
using CRT.Server.Handlers.Updates;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Which release's server build is next (ServerReleaseFeed, 2026-11-09): the newest carrying one, a
    // release for the live service and a release or pre-release for the test instance - never a draft,
    // never files from anywhere but this repository's own downloads.
    // ###########################################################################################
    public sealed class ServerReleaseFeedTests
    {
        private static object Release(string tag, bool preRelease, DateTimeOffset published, bool draft = false, long size = 20_000_000, string? archiveUrl = null, bool withBundle = true)
        {
            var assets = new List<object>
            {
                new { id = 11, name = ServerReleaseFeed.ArchiveName, size, browser_download_url = archiveUrl ?? ServerReleaseFeed.DownloadPrefix + tag + "/" + ServerReleaseFeed.ArchiveName },
                new { id = 12, name = "Classic-Repair-Toolbox-win-Setup.exe", size = 90_000_000, browser_download_url = ServerReleaseFeed.DownloadPrefix + tag + "/Classic-Repair-Toolbox-win-Setup.exe" },
            };

            if (withBundle)
                assets.Add(new { id = 13, name = ServerReleaseFeed.BundleName, size = 9000, browser_download_url = ServerReleaseFeed.DownloadPrefix + tag + "/" + ServerReleaseFeed.BundleName });

            return new { tag_name = tag, prerelease = preRelease, draft, published_at = published.ToString("O"), assets };
        }

        private static readonly DateTimeOffset Day = new(2026, 11, 9, 12, 0, 0, TimeSpan.Zero);

        private static ServerReleaseFeed.Candidate? Newest(string channel, params object[] releases) =>
            ServerReleaseFeed.Newest(JsonSerializer.Serialize(releases), channel, out _, out _);

        [Fact]
        public void The_live_service_takes_the_newest_release_and_the_test_instance_the_newest_of_both()
        {
            object[] releases =
            [
                ServerReleaseFeedTests.Release("4.2.0-beta.1", preRelease: true, ServerReleaseFeedTests.Day.AddDays(2)),
                ServerReleaseFeedTests.Release("4.1.0", preRelease: false, ServerReleaseFeedTests.Day.AddDays(1)),
                ServerReleaseFeedTests.Release("4.0.0", preRelease: false, ServerReleaseFeedTests.Day),
            ];

            Assert.Equal("4.1.0", ServerReleaseFeedTests.Newest(ServerUpdateOutcomes.ReleasesChannel, releases)!.Tag);
            Assert.Equal("4.2.0-beta.1", ServerReleaseFeedTests.Newest(ServerUpdateOutcomes.PreReleasesChannel, releases)!.Tag);
        }

        // GitHub's own order is not trusted to be the newest first: the publishing time decides.
        [Fact]
        public void The_newest_is_the_one_published_last_whatever_the_order()
        {
            Assert.Equal("4.1.0", ServerReleaseFeedTests.Newest(
                ServerUpdateOutcomes.ReleasesChannel,
                ServerReleaseFeedTests.Release("4.0.0", false, ServerReleaseFeedTests.Day),
                ServerReleaseFeedTests.Release("4.1.0", false, ServerReleaseFeedTests.Day.AddDays(1)))!.Tag);
        }

        [Fact]
        public void A_draft_a_release_without_both_files_or_a_build_too_large_is_passed_over()
        {
            Assert.Null(ServerReleaseFeedTests.Newest(ServerUpdateOutcomes.PreReleasesChannel, ServerReleaseFeedTests.Release("4.1.0", false, ServerReleaseFeedTests.Day, draft: true)));
            Assert.Null(ServerReleaseFeedTests.Newest(ServerUpdateOutcomes.PreReleasesChannel, ServerReleaseFeedTests.Release("4.1.0", false, ServerReleaseFeedTests.Day, withBundle: false)));
            Assert.Null(ServerReleaseFeedTests.Newest(ServerUpdateOutcomes.PreReleasesChannel, ServerReleaseFeedTests.Release("4.1.0", false, ServerReleaseFeedTests.Day, size: ServerReleaseFeed.ArchiveMaxBytes + 1)));

            // Older CRT releases carry no server build at all.
            Assert.Equal("4.0.0", ServerReleaseFeedTests.Newest(
                ServerUpdateOutcomes.ReleasesChannel,
                new { tag_name = "4.0.1", prerelease = false, draft = false, published_at = ServerReleaseFeedTests.Day.AddDays(1).ToString("O"), assets = Array.Empty<object>() },
                ServerReleaseFeedTests.Release("4.0.0", false, ServerReleaseFeedTests.Day))!.Tag);
        }

        // ###########################################################################################
        // *** NOTHING IS DOWNLOADED FROM ANYWHERE ELSE. *** A list somebody tampered with names other places:
        // another host, another repository, another release's file, a path climbing out, plain HTTP.
        // ###########################################################################################
        [Theory]
        [InlineData("https://evil.example/crt-server-linux-x64.tar.gz")]
        [InlineData("https://github.com/someone/Classic-Repair-Toolbox/releases/download/4.1.0/crt-server-linux-x64.tar.gz")]
        [InlineData("https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/download/4.0.0/crt-server-linux-x64.tar.gz")]
        [InlineData("https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/download/4.1.0/../4.0.0/crt-server-linux-x64.tar.gz")]
        [InlineData("http://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/download/4.1.0/crt-server-linux-x64.tar.gz")]
        public void Files_anywhere_but_this_releases_own_downloads_are_passed_over(string archiveUrl)
        {
            Assert.Null(ServerReleaseFeedTests.Newest(
                ServerUpdateOutcomes.ReleasesChannel,
                ServerReleaseFeedTests.Release("4.1.0", false, ServerReleaseFeedTests.Day, archiveUrl: archiveUrl)));
        }

        [Fact]
        public void The_trusted_source_is_this_repositorys_release_workflow_on_main()
        {
            Assert.Equal("HovKlan-DH/Classic-Repair-Toolbox", ServerReleaseFeed.Repository);
            Assert.Equal("HovKlan-DH/Classic-Repair-Toolbox/.github/workflows/build-and-release.yml", ServerReleaseFeed.SignerWorkflow);
            Assert.Equal("refs/heads/main", ServerReleaseFeed.SourceRef);
            Assert.StartsWith("https://api.github.com/repos/HovKlan-DH/Classic-Repair-Toolbox/releases", ServerReleaseFeed.ReleasesUrl, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The releases feed the minute's look reads (2026-11-09; GitHub's releases.atom, its shape as GitHub
        // writes it): one string that moves whenever a release is published, edited or deleted, and nothing for
        // what is no Atom feed - an error page, or a DTD, which is never expanded.
        // ###########################################################################################
        private static string Feed(params (string Tag, string Updated)[] entries) =>
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
            "<feed xmlns=\"http://www.w3.org/2005/Atom\" xmlns:media=\"http://search.yahoo.com/mrss/\" xml:lang=\"en-US\">\n" +
            "  <id>tag:github.com,2008:https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases</id>\n" +
            "  <title>Release notes from Classic-Repair-Toolbox</title>\n" +
            string.Concat(entries.Select(entry =>
                "  <entry>\n" +
                $"    <id>tag:github.com,2008:Repository/123/{entry.Tag}</id>\n" +
                $"    <updated>{entry.Updated}</updated>\n" +
                $"    <link rel=\"alternate\" type=\"text/html\" href=\"https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases/tag/{entry.Tag}\"/>\n" +
                $"    <title>{entry.Tag}</title>\n" +
                "    <content type=\"html\">&lt;p&gt;Notes&lt;/p&gt;</content>\n" +
                "  </entry>\n")) +
            "</feed>\n";

        [Fact]
        public void The_feeds_fingerprint_moves_with_its_releases_and_only_with_them()
        {
            string one = ServerReleaseFeed.FingerprintOf(ServerReleaseFeedTests.Feed(("4.0.0", "2026-11-01T10:00:00Z")))!;

            Assert.Equal(64, one.Length);
            Assert.Equal(one, ServerReleaseFeed.FingerprintOf(ServerReleaseFeedTests.Feed(("4.0.0", "2026-11-01T10:00:00Z"))));

            // Published, edited, deleted.
            Assert.NotEqual(one, ServerReleaseFeed.FingerprintOf(ServerReleaseFeedTests.Feed(("4.1.0-beta.1", "2026-11-09T10:00:00Z"), ("4.0.0", "2026-11-01T10:00:00Z"))));
            Assert.NotEqual(one, ServerReleaseFeed.FingerprintOf(ServerReleaseFeedTests.Feed(("4.0.0", "2026-11-02T10:00:00Z"))));
            Assert.NotEqual(one, ServerReleaseFeed.FingerprintOf(ServerReleaseFeedTests.Feed()));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("<html><body>Rate limited</body></html>")]
        [InlineData("not xml")]
        [InlineData("<?xml version=\"1.0\"?><!DOCTYPE feed [<!ENTITY x \"y\">]><feed xmlns=\"http://www.w3.org/2005/Atom\"><entry><id>&x;</id></entry></feed>")]
        public void What_is_no_Atom_feed_has_no_fingerprint(string? body) =>
            Assert.Null(ServerReleaseFeed.FingerprintOf(body));

        [Fact]
        public void An_answer_that_is_no_list_is_unreadable_and_no_build_yet_is_said()
        {
            Assert.Null(ServerReleaseFeed.Newest("{\"message\":\"API rate limit exceeded\"}", ServerUpdateOutcomes.ReleasesChannel, out _, out bool readable));
            Assert.False(readable);

            Assert.Null(ServerReleaseFeed.Newest("not json", ServerUpdateOutcomes.ReleasesChannel, out _, out readable));
            Assert.False(readable);

            Assert.Null(ServerReleaseFeed.Newest("[]", ServerUpdateOutcomes.ReleasesChannel, out string why, out readable));
            Assert.True(readable);
            Assert.StartsWith("No release", why, StringComparison.Ordinal);
        }
    }
}
