using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // THE NEW SUBMISSION RULES AGAINST EVERY BOARD THAT IS ALREADY PUBLISHED (security review,
    // 2026-09-25).
    //
    // *** WHY THIS TEST EXISTS. *** This pipeline has twice shipped a validation rule that
    // rejected correct, already-published data - region-variant labels, then note-only image rows
    // - each found only when the maintainer's first real submission of a board failed. The file
    // rules added in the security review are stricter still (own folder, allowed types, content
    // signatures), so every one of them is run here over every board in Assets/Data, exactly as a
    // client would submit it.
    //
    // It is what found the case the rules had to allow: C128DCR 250477 cites two scope-baseline
    // texts that live in the C128 310378 folder. A plain "own folder only" rule refused a board
    // that is live today; the rule now allows another board's file when it is left unchanged.
    //
    // Reads the real shipped tree - no network, no display - so it runs on CI like any other test.
    // In the "BoardData" collection because BoardDataReader's cache is shared static state.
    // ###########################################################################################
    [Collection("BoardData")]
    public sealed class SubmissionRulesShippedDataTests
    {
        private static string DataRoot()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);

            return Path.Combine(folder!, "Assets", "Data");
        }

        // Every shipped board's newest-generation workbook, as (relative folder, workbook path).
        private static IEnumerable<(string Folder, string Workbook)> ShippedBoards(string dataRoot)
        {
            foreach (string folder in Directory.EnumerateDirectories(dataRoot, "*", SearchOption.AllDirectories))
            {
                List<string> names = Directory.EnumerateFiles(folder, "Data *" + DataGenerationRules.WorkbookExtension)
                    .Select(Path.GetFileName)
                    .Select(name => name!)
                    .ToList();

                if (names.Count == 0)
                    continue;

                Version? generation = DataGenerationRules.ResolveNewestGeneration(names);
                string workbook = names.First(name => DataGenerationRules.TryReadGeneration(name) == generation);

                yield return (Path.GetRelativePath(dataRoot, folder).Replace(Path.DirectorySeparatorChar, '/'), Path.Combine(folder, workbook));
            }
        }

        // The published tree as the server would see it: real listings, real hashes.
        private static PublishedTreeView TreeOver(string dataRoot)
        {
            return new PublishedTreeView(
                relative =>
                {
                    string full = Path.Combine(dataRoot, relative.Replace('/', Path.DirectorySeparatorChar));
                    return File.Exists(full) ? SubmissionRulesShippedDataTests.HashOf(full) : null;
                },
                relative =>
                {
                    string full = relative.Length == 0 ? dataRoot : Path.Combine(dataRoot, relative.Replace('/', Path.DirectorySeparatorChar));

                    return Directory.Exists(full)
                        ? Directory.EnumerateFileSystemEntries(full).Select(Path.GetFileName).Select(name => name!).ToList()
                        : null;
                });
        }

        private static string HashOf(string path)
        {
            using FileStream stream = File.OpenRead(path);
            return Convert.ToHexStringLower(SHA256.HashData(stream));
        }

        [Fact]
        public async Task No_published_board_is_refused_by_the_file_rules()
        {
            string dataRoot = SubmissionRulesShippedDataTests.DataRoot();
            PublishedTreeView tree = SubmissionRulesShippedDataTests.TreeOver(dataRoot);
            var refusals = new List<string>();
            int boards = 0;

            foreach ((string folder, string workbook) in SubmissionRulesShippedDataTests.ShippedBoards(dataRoot))
            {
                string cacheKey = "shipped-rules:" + folder;
                BoardData? board;

                try
                {
                    board = await BoardDataReader.LoadAsync(workbook, cacheKey);
                }
                finally
                {
                    BoardDataReader.ClearCache(cacheKey);
                }

                Assert.NotNull(board);
                boards++;

                string[] parts = folder.Split('/');

                // Built as the client builds it: the files are exactly what the rows cite. Only a
                // file outside the board's own and shared folders needs its real hash - the rule
                // compares it with what is published; every other hash is not looked at here.
                var manifest = new SubmissionManifest
                {
                    SystemId = folder,
                    Manufacturer = parts[0],
                    Hardware = parts[1],
                    Board = parts[2],
                    Rows = new SubmissionRows
                    {
                        Schematics = [.. board!.Schematics],
                        ComponentImages = [.. board.ComponentImages],
                        ComponentLocalFiles = [.. board.ComponentLocalFiles],
                        BoardLocalFiles = [.. board.BoardLocalFiles]
                    }
                };

                foreach (string path in SubmissionManifestBuilder.CollectReferencedFiles(board))
                {
                    bool foreign = SubmissionFileScopes.Classify(parts[0], parts[1], parts[2], path) == SubmissionFileScope.Foreign;

                    manifest.Files.Add(new SubmissionFile
                    {
                        Path = path,
                        Sha256 = foreign
                            ? SubmissionRulesShippedDataTests.HashOf(Path.Combine(dataRoot, path.Replace('/', Path.DirectorySeparatorChar)))
                            : new string('0', 64)
                    });
                }

                refusals.AddRange(SubmissionFileRules.ValidateManifestFiles(manifest, tree)
                    .Select(finding => $"{folder}: {finding.Code} {finding.Subject}"));
            }

            // Anti-vacuity: an empty walk would pass without checking anything.
            Assert.True(boards >= 10, $"Only {boards} shipped boards were found under {dataRoot}.");
            Assert.Empty(refusals);
        }

        // ###########################################################################################
        // *** EVERY PUBLISHED BOARD'S NAME PASSES THE IDENTITY RULES (code review, 2026-09-25). ***
        // A submission names its board by the folder names, and SubmissionValidator refuses parts
        // that are not canonical (identity.parts_not_canonical - no doubled, leading or trailing
        // spaces) or that do not build the id they came with. A board folder named with two spaces
        // in a row is legal on every filesystem, and nobody could ever contribute to it. This runs
        // the real validator over every board folder, which the file-rule test above does not.
        // ###########################################################################################
        [Fact]
        public void Every_published_boards_folder_names_pass_the_identity_rules()
        {
            string dataRoot = SubmissionRulesShippedDataTests.DataRoot();
            var refusals = new List<string>();
            int boards = 0;

            foreach ((string folder, _) in SubmissionRulesShippedDataTests.ShippedBoards(dataRoot))
            {
                boards++;
                string[] parts = folder.Split('/');

                var manifest = new SubmissionManifest
                {
                    SystemId = folder,
                    Manufacturer = parts[0],
                    Hardware = parts[1],
                    Board = parts[2]
                };

                refusals.AddRange(SubmissionValidator.Validate(manifest, [])
                    .Where(finding => finding.Code.StartsWith("identity.", StringComparison.Ordinal))
                    .Select(finding => $"{folder}: {finding.Code}"));
            }

            Assert.True(boards >= 10, $"Only {boards} shipped boards were found under {dataRoot}.");
            Assert.Empty(refusals);
        }

        [Fact]
        public async Task Every_file_a_published_board_cites_passes_the_content_rules()
        {
            string dataRoot = SubmissionRulesShippedDataTests.DataRoot();
            var cited = new SortedSet<string>(StringComparer.Ordinal);

            foreach ((string folder, string workbook) in SubmissionRulesShippedDataTests.ShippedBoards(dataRoot))
            {
                string cacheKey = "shipped-content:" + folder;

                try
                {
                    BoardData? board = await BoardDataReader.LoadAsync(workbook, cacheKey);
                    cited.UnionWith(SubmissionManifestBuilder.CollectReferencedFiles(board!));
                }
                finally
                {
                    BoardDataReader.ClearCache(cacheKey);
                }
            }

            var refusals = new List<string>();

            foreach (string path in cited)
            {
                string full = Path.Combine(dataRoot, path.Replace('/', Path.DirectorySeparatorChar));
                byte[] head = new byte[SubmissionContentRules.BytesNeeded(path)];
                int length;

                using (FileStream stream = File.OpenRead(full))
                    length = stream.ReadAtLeast(head, head.Length, throwOnEndOfStream: false);

                if (!SubmissionContentRules.Matches(path, head.AsSpan(0, length), out string reason))
                    refusals.Add($"{path}: {reason}");
            }

            Assert.True(cited.Count > 1000, $"Only {cited.Count} cited files were found.");
            Assert.Empty(refusals);
        }
    }
}
