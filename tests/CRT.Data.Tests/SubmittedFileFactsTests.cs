using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmittedFileFacts - the per-file account the server sends the Maintainer tab so
    // a maintainer is shown EVERY file that would change, not only the images (security review,
    // 2026-09-25). In the CRT 4.0 tree (2026-10-25) a file is the board's name for it, compared with
    // what the published board lists under that name.
    // ###########################################################################################
    public sealed class SubmittedFileFactsTests
    {
        private static SubmissionManifest Manifest()
        {
            var manifest = new SubmissionManifest
            {
                BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407"
            };

            manifest.Files.Add(new SubmissionFile { Path = "manual.pdf", Sha256 = new string('1', 64), SizeBytes = 5 });
            manifest.Files.Add(new SubmissionFile { Path = "Images/x.png", Sha256 = new string('2', 64), SizeBytes = 6 });
            manifest.Files.Add(new SubmissionFile { Path = "stray.txt", Sha256 = new string('3', 64), SizeBytes = 7 });
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = "manual.pdf" });
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = "Images/x.png" });

            return manifest;
        }

        [Fact]
        public void Each_file_says_whether_a_row_uses_it_and_what_the_published_board_lists_under_its_name()
        {
            IReadOnlyList<SubmittedFileFact> facts = SubmittedFileFacts.Build(
                SubmittedFileFactsTests.Manifest(),
                new Dictionary<string, string>
                {
                    ["Images/x.png"] = new string('9', 64),
                    ["manual.pdf"] = new string('1', 64)
                });

            Assert.Equal(3, facts.Count);

            SubmittedFileFact replaced = facts[0];
            Assert.Equal("Images/x.png", replaced.Path);
            Assert.True(replaced.IsReferenced);
            Assert.True(replaced.ExistsOnServer);
            Assert.False(replaced.IsUnchanged);

            SubmittedFileFact manual = facts[1];
            Assert.Equal("manual.pdf", manual.Path);
            Assert.True(manual.IsReferenced);
            Assert.True(manual.IsUnchanged);

            SubmittedFileFact stray = facts[2];
            Assert.False(stray.IsReferenced);
            Assert.False(stray.ExistsOnServer);
        }

        // Every file is the board's own in the 4.0 tree - there are no shared folders or other
        // boards' files to carry any more.
        [Fact]
        public void Every_file_is_the_boards_own()
        {
            Assert.All(
                SubmittedFileFacts.Build(SubmittedFileFactsTests.Manifest(), null),
                fact => Assert.Equal(SubmissionFileScope.Own, fact.Scope));
        }

        // The scope must travel as its NAME: a number would shift every file's label silently the
        // day a member is added. The enum carries its own converter, so this holds whatever options
        // either end serialises with.
        [Fact]
        public void The_scope_is_written_as_its_name_on_the_wire()
        {
            string json = JsonSerializer.Serialize(
                SubmittedFileFacts.Build(SubmittedFileFactsTests.Manifest(), null),
                new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });

            Assert.Contains("\"scope\":\"Own\"", json, System.StringComparison.Ordinal);
        }
    }
}
