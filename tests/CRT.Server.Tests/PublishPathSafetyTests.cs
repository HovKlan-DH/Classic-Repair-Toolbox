using CRT.Server.Handlers.Submissions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers PublishPathSafety.FindLinkOnPath - the check that stops a publish writing THROUGH a
    // symbolic link inside the data tree (security review, 2026-09-25).
    //
    // The "is this a link" question is a delegate, so the walk is tested without creating real
    // links - which Windows allows only with elevated rights, and which would make this test pass
    // on one machine and not another.
    // ###########################################################################################
    public sealed class PublishPathSafetyTests
    {
        private static readonly string Root = Path.Combine(Path.GetTempPath(), "crt-link-root");

        private static string Under(params string[] parts) => Path.Combine([PublishPathSafetyTests.Root, .. parts]);

        [Fact]
        public void A_path_with_no_links_passes()
        {
            Assert.Null(PublishPathSafety.FindLinkOnPath(
                PublishPathSafetyTests.Root,
                PublishPathSafetyTests.Under("Commodore", "C64", "250407", "a.png"),
                _ => false));
        }

        // A link part way down redirects everything beneath it - including a perfectly contained
        // submitted path.
        [Fact]
        public void A_linked_FOLDER_on_the_way_is_found()
        {
            string linked = PublishPathSafetyTests.Under("Commodore", "C64");

            Assert.Equal(linked, PublishPathSafety.FindLinkOnPath(
                PublishPathSafetyTests.Root,
                PublishPathSafetyTests.Under("Commodore", "C64", "250407", "a.png"),
                path => path == linked));
        }

        // Writing over a link would write wherever it points.
        [Fact]
        public void A_linked_FILE_at_the_destination_is_found()
        {
            string target = PublishPathSafetyTests.Under("Commodore", "C64", "250407", "a.png");

            Assert.Equal(target, PublishPathSafety.FindLinkOnPath(PublishPathSafetyTests.Root, target, path => path == target));
        }

        // The root itself is the operator's choice - a tree mounted from elsewhere is legitimate -
        // so it is never asked about.
        [Fact]
        public void The_data_root_itself_is_not_questioned()
        {
            var asked = new List<string>();

            PublishPathSafety.FindLinkOnPath(
                PublishPathSafetyTests.Root,
                PublishPathSafetyTests.Under("a.png"),
                path =>
                {
                    asked.Add(path);
                    return false;
                });

            Assert.DoesNotContain(Path.GetFullPath(PublishPathSafetyTests.Root), asked);
            Assert.Single(asked);
        }

        // Nothing should ever ask about a path outside the root; if something does, refuse.
        [Fact]
        public void A_target_outside_the_root_is_refused()
        {
            Assert.NotNull(PublishPathSafety.FindLinkOnPath(
                PublishPathSafetyTests.Root,
                Path.Combine(Path.GetTempPath(), "elsewhere", "a.png"),
                _ => false));
        }

        // A path that does not exist yet - the ordinary case for a new file - is not a link.
        [Fact]
        public void A_path_that_does_not_exist_is_not_a_link()
        {
            Assert.False(PublishPathSafety.IsLink(PublishPathSafetyTests.Under(Guid.NewGuid().ToString("N"), "a.png")));
        }
    }
}
