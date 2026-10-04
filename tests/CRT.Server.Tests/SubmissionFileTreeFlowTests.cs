using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SubmissionFileTreeFlow - a submission's files as the BETA data will hold them after
    // approving it (owner request, 2026-09-28), against a real temp tree: the folder listing, the
    // publish plan and the highlight-file check are all real.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class SubmissionFileTreeFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 28, 12, 0, 0, TimeSpan.Zero);

        private const string Board = "Commodore/C64/250407";

        private readonly string thisRoot;
        private readonly string thisDataTree;

        public SubmissionFileTreeFlowTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-file-tree", Guid.NewGuid().ToString("N"));
            this.thisDataTree = Path.Combine(this.thisRoot, "app-data-BETA");
            Directory.CreateDirectory(this.thisDataTree);

            // The masters are the generations a new system's workbook is named after - see
            // ApprovePublishFlowTests.
            File.WriteAllText(Path.Combine(this.thisDataTree, "Classic-Repair-Toolbox.xlsx"), "master");
            DataTreeBuilder.ListingMaster(this.thisDataTree, ApprovePublishFlowTests.OtherListedBoard);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private string InTree(string relative) =>
            Path.Combine(this.thisDataTree, relative.Replace('/', Path.DirectorySeparatorChar));

        private static SubmissionManifest Manifest(string summary, params SubmissionFile[] files)
        {
            var manifest = new SubmissionManifest
            {
                SystemId = SubmissionFileTreeFlowTests.Board,
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Files = [.. files]
            };

            foreach (SubmissionFile file in files)
                manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = file.Path });

            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", Description = summary });
            manifest.Rows.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1",
                BoardLabel = "U8",
                X = "100",
                Y = "200",
                Width = "40",
                Height = "20"
            });

            return manifest;
        }

        private Task<IReadOnlyList<SystemFileEntry>> BuildAsync(SubmissionManifest manifest) =>
            SubmissionFileTreeFlow.BuildAsync(
                this.thisDataTree, manifest, new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance), SubmissionFileTreeFlowTests.Now);

        private static SystemFileEntry Of(IReadOnlyList<SystemFileEntry> entries, string path) =>
            Assert.Single(entries, entry => entry.Path == path);

        // ###########################################################################################
        // A NEW SYSTEM "should just show how the new file-system will look like": its file new, and
        // the workbook and highlight file the approval writes new too, with nothing to open yet.
        // ###########################################################################################
        [Fact]
        public async Task A_new_systems_tree_is_all_new()
        {
            IReadOnlyList<SystemFileEntry> entries = await this.BuildAsync(SubmissionFileTreeFlowTests.Manifest(
                "b",
                new SubmissionFile { Path = $"{SubmissionFileTreeFlowTests.Board}/manual.pdf", Sha256 = new string('a', 64), SizeBytes = 10 }));

            Assert.All(entries, entry => Assert.Equal(SystemFileChange.Added, entry.Change));
            Assert.Equal(SystemFileSource.Submission, Of(entries, $"{SubmissionFileTreeFlowTests.Board}/manual.pdf").OpenFrom);

            SystemFileEntry workbook = Of(entries, $"{SubmissionFileTreeFlowTests.Board}/Data C64 250407 v2.0.0.xlsx");
            Assert.True(workbook.WrittenOnApproval);
            Assert.Equal(SystemFileSource.NotWrittenYet, workbook.OpenFrom);

            // Sizes (owner request, 2026-10-04): the upload's, from the manifest; none for a file
            // with nothing to open yet.
            Assert.Equal(10, Of(entries, $"{SubmissionFileTreeFlowTests.Board}/manual.pdf").SizeBytes);
            Assert.Null(workbook.SizeBytes);
            Assert.Equal(SystemFileSource.NotWrittenYet, Of(entries, $"{SubmissionFileTreeFlowTests.Board}/Data C64 250407 v2.0.0.json").OpenFrom);
        }

        // ###########################################################################################
        // *** THE REPORTED CASE (2026-09-28): only a one-liner changed. *** The workbook is written
        // again, the highlight file is NOT - its content is the same - and the rest of the board's
        // folder is there around them, unchanged: an older generation's workbook, a file nothing
        // cites, hidden and retired files left out.
        // ###########################################################################################
        [Fact]
        public async Task A_rows_only_change_shows_only_the_workbook_changing_among_the_whole_folder()
        {
            SubmissionManifest before = SubmissionFileTreeFlowTests.Manifest("a");
            string workbook = this.InTree($"{SubmissionFileTreeFlowTests.Board}/Data C64 250407 v2.0.0.xlsx");

            Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
            BoardWorkbookWriter.Write(workbook, PublishMerge.Build(before, null));
            BoardSidecarWriter.Write(workbook, before.Rows.ComponentHighlights, []);

            File.WriteAllText(this.InTree($"{SubmissionFileTreeFlowTests.Board}/Data C64 250407.xlsx"), "older generation");
            Directory.CreateDirectory(this.InTree($"{SubmissionFileTreeFlowTests.Board}/Images"));
            File.WriteAllText(this.InTree($"{SubmissionFileTreeFlowTests.Board}/Images/loose.png"), "x");
            File.WriteAllText(this.InTree($"{SubmissionFileTreeFlowTests.Board}/.Data.json.1234.writing.tmp"), "temporary");
            File.WriteAllText(this.InTree($"{SubmissionFileTreeFlowTests.Board}/{SystemDescriptorStore.FileName}"), "{}");

            IReadOnlyList<SystemFileEntry> entries = await this.BuildAsync(SubmissionFileTreeFlowTests.Manifest("b"));

            Assert.Equal(
                [$"{SubmissionFileTreeFlowTests.Board}/Data C64 250407 v2.0.0.xlsx"],
                entries.Where(entry => entry.Change != SystemFileChange.Unchanged).Select(entry => entry.Path));

            Assert.Equal(SystemFileChange.Unchanged, Of(entries, $"{SubmissionFileTreeFlowTests.Board}/Data C64 250407 v2.0.0.json").Change);
            Assert.Equal(SystemFileChange.Unchanged, Of(entries, $"{SubmissionFileTreeFlowTests.Board}/Data C64 250407.xlsx").Change);
            Assert.Equal(SystemFileChange.Unchanged, Of(entries, $"{SubmissionFileTreeFlowTests.Board}/Images/loose.png").Change);

            Assert.DoesNotContain(entries, entry => entry.Path.Contains("/.", StringComparison.Ordinal));
            Assert.DoesNotContain(entries, entry => entry.Path.EndsWith(SystemDescriptorStore.FileName, StringComparison.Ordinal));
        }

        // A moved highlight is a changed highlight file.
        [Fact]
        public async Task A_moved_highlight_changes_the_highlight_file()
        {
            SubmissionManifest before = SubmissionFileTreeFlowTests.Manifest("a");
            string workbook = this.InTree($"{SubmissionFileTreeFlowTests.Board}/Data C64 250407 v2.0.0.xlsx");

            Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
            BoardWorkbookWriter.Write(workbook, PublishMerge.Build(before, null));
            BoardSidecarWriter.Write(workbook, before.Rows.ComponentHighlights, []);

            SubmissionManifest moved = SubmissionFileTreeFlowTests.Manifest("a");
            moved.Rows.ComponentHighlights[0] = new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1",
                BoardLabel = "U8",
                X = "101",
                Y = "200",
                Width = "40",
                Height = "20"
            };

            IReadOnlyList<SystemFileEntry> entries = await this.BuildAsync(moved);

            Assert.Equal(SystemFileChange.Changed, Of(entries, $"{SubmissionFileTreeFlowTests.Board}/Data C64 250407 v2.0.0.json").Change);
        }
    }
}
