using CRT.Server.Handlers.Health;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The server's own changelog (owner request, 2026-11-24: 'have a "Changelog" like there is in "About" ... so I
    // can check what has changed in server also'): RELEASE-NOTES.md built into the program, sent to the Maintainer
    // tab's Account > "Server version" without the comment that opens the file - CRT's Markdown reader would show
    // that comment as typed.
    // ###########################################################################################
    public class ServerReleaseNotesTests
    {
        [Fact]
        public void The_comment_opening_the_notes_is_left_out_and_the_versions_kept()
        {
            string notes = "<!--\nThe rules for whoever writes it.\n- A heading is \"# <version>, ...\"\n-->\n\n# 12.2.0, 2026-November-24\n------\n- One\n\n# 12.1.1, 2026-November-24\n------\n- Two\n";

            Assert.Equal(
                "# 12.2.0, 2026-November-24\n------\n- One\n\n# 12.1.1, 2026-November-24\n------\n- Two",
                ServerReleaseNotes.ShownPart(notes));
        }

        // Windows line ends and a byte-order mark an editor put in front read the same.
        [Fact]
        public void Windows_line_ends_and_a_byte_order_mark_change_nothing() =>
            Assert.Equal(
                "# 1.0.0, 2026-October-9\n---\n- First",
                ServerReleaseNotes.ShownPart("﻿<!--\r\nRules\r\n-->\r\n# 1.0.0, 2026-October-9\r\n---\r\n- First\r\n"));

        [Fact]
        public void Notes_with_no_comment_are_sent_as_they_stand() =>
            Assert.Equal("# 1.0.0, 2026-October-9\n- First", ServerReleaseNotes.ShownPart("  # 1.0.0, 2026-October-9\n- First\n\n"));

        // Nothing to show is null, so CRT says nothing rather than draw an empty changelog - and a comment that
        // never closes is all comment.
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   \n ")]
        [InlineData("<!-- only the rules -->")]
        [InlineData("<!--\nrules that never close\n# 1.0.0, 2026-October-9")]
        public void Nothing_left_to_show_is_null(string? notes) =>
            Assert.Null(ServerReleaseNotes.ShownPart(notes));

        // The build carries them (CRT.Server.csproj's EmbeddedResource) - left out, the "Server version" page would
        // say the server sent no changelog. Not WHAT they say: the Stop hook server-version-bump.sh holds each
        // version to a section of its own.
        [Fact]
        public void The_build_carries_its_notes_from_a_version_heading_on()
        {
            string? text = ServerReleaseNotes.Text;

            Assert.NotNull(text);
            Assert.StartsWith("# ", text);
            Assert.DoesNotContain("<!--", text);
        }
    }
}
