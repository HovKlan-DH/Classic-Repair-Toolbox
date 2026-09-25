using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmittedFileFacts - the per-file account the server sends the review application so
    // a reviewer is shown EVERY file that would change, not only the images (security review,
    // 2026-09-25).
    // ###########################################################################################
    public sealed class SubmittedFileFactsTests
    {
        private static SubmissionManifest Manifest()
        {
            var manifest = new SubmissionManifest
            {
                SystemId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407"
            };

            manifest.Files.Add(new SubmissionFile { Path = "Commodore/C64/250407/manual.pdf", Sha256 = new string('1', 64), SizeBytes = 5 });
            manifest.Files.Add(new SubmissionFile { Path = "Commodore/Shared files/x.png", Sha256 = new string('2', 64), SizeBytes = 6 });
            manifest.Files.Add(new SubmissionFile { Path = "Commodore/C64/250425/stray.txt", Sha256 = new string('3', 64), SizeBytes = 7 });
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = "Commodore/C64/250407/manual.pdf" });
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = "Commodore/Shared files/x.png" });

            return manifest;
        }

        [Fact]
        public void Each_file_carries_its_scope_whether_a_row_uses_it_and_what_is_published_there()
        {
            IReadOnlyList<SubmittedFileFact> facts = SubmittedFileFacts.Build(
                SubmittedFileFactsTests.Manifest(),
                new Dictionary<string, string>
                {
                    ["Commodore/Shared files/x.png"] = new string('9', 64),
                    ["Commodore/C64/250407/manual.pdf"] = new string('1', 64)
                });

            Assert.Equal(3, facts.Count);

            SubmittedFileFact manual = facts[0];
            Assert.Equal("Commodore/C64/250407/manual.pdf", manual.Path);
            Assert.Equal(SubmissionFileScope.Own, manual.Scope);
            Assert.True(manual.IsReferenced);
            Assert.True(manual.IsUnchanged);

            SubmittedFileFact stray = facts[1];
            Assert.Equal(SubmissionFileScope.Foreign, stray.Scope);
            Assert.False(stray.IsReferenced);
            Assert.False(stray.ExistsOnServer);

            SubmittedFileFact shared = facts[2];
            Assert.Equal(SubmissionFileScope.ManufacturerShared, shared.Scope);
            Assert.True(shared.ExistsOnServer);
            Assert.False(shared.IsUnchanged);
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
            Assert.Contains("\"scope\":\"Foreign\"", json, System.StringComparison.Ordinal);
        }
    }
}
