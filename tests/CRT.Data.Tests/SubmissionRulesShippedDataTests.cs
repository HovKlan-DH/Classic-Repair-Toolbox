using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    // - each found only when the project owner's first real submission of a board failed. The file
    // rules added in the security review are stricter still (allowed types, cited by a row, content
    // signatures), so every one of them is run here over every board in Assets/Data, exactly as a
    // client would submit it - since 2026-10-25 as the CRT 4.0 tree holds it (the converter's own
    // plan of Assets/Data), so the files are the board's names.
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

        // ###########################################################################################
        // *** EVERY SHIPPED BOARD, CONVERTED TO THE CRT 4.0 TREE, PASSES THE SERVER'S RULES AS A 4.0
        // CRT SENDS IT (2026-10-25). *** The board's rows and every file its board file lists - by
        // name, with their checksums, its KiCad data included - judged against that same board as
        // published: an unchanged board sent again must never be refused, by the path rules, the
        // file rules or the publish plan, nor any file of it.
        // ###########################################################################################
        [Fact]
        public void No_converted_shipped_board_is_refused_by_the_file_rules_or_the_publish_plan()
        {
            DataTreeConversionPlan plan = ClassicRepairToolbox.Tests.DataTreeConversionShippedDataTests.Planned.Value;
            var refusals = new List<string>();

            Assert.Empty(plan.Problems);

            foreach (ConvertedBoard converted in plan.Boards)
            {
                BoardData board = BoardDocumentMapping.ToContent(converted.Document).Board;
                string[] parts = converted.BoardId.Split('/');

                var listed = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

                foreach (BoardDocument.StoredFileRow file in converted.Document.Files)
                    listed.TryAdd(file.Name, file);

                var published = new PublishedTreeView(
                    name => listed.TryGetValue(name, out BoardDocument.StoredFileRow? file) ? file.Sha256 : null,
                    _ => null);

                var manifest = new SubmissionManifest
                {
                    BoardId = converted.BoardId,
                    Manufacturer = parts[0],
                    Hardware = parts[1],
                    Board = parts[2],
                    Rows = new SubmissionRows
                    {
                        Schematics = [.. board.Schematics],
                        ComponentImages = [.. board.ComponentImages],
                        ComponentLocalFiles = [.. board.ComponentLocalFiles],
                        BoardLocalFiles = [.. board.BoardLocalFiles]
                    }
                };

                // Built as the client builds it: what the rows cite, then the KiCad data.
                foreach (string name in SubmissionManifestBuilder.CollectReferencedFiles(board).Concat(SubmissionKiCadFiles.Collect(converted.Document)).Distinct(StringComparer.Ordinal))
                {
                    if (!listed.TryGetValue(name, out BoardDocument.StoredFileRow? file))
                    {
                        refusals.Add($"{converted.BoardId}: [{name}] is cited but not in the board's list of files");
                        continue;
                    }

                    manifest.Files.Add(new SubmissionFile { Path = name, Sha256 = file.Sha256, SizeBytes = file.Size });
                }

                refusals.AddRange(SubmissionPathRules.ValidateManifestPaths(manifest, Path.GetTempPath())
                    .Concat(SubmissionFileRules.ValidateManifestFiles(manifest, published))
                    .Select(finding => $"{converted.BoardId}: {finding.Code} {finding.Subject}"));

                PublishPlanResult publish = PublishPlan.Build(
                    Path.GetTempPath(), manifest, "2026-October-25", DateTimeOffset.UtcNow, null, BoardDescriptorRules.BoardOrigin.Contributed, published, _ => true);

                refusals.AddRange(publish.Problems.Select(problem => $"{converted.BoardId}: {problem.Code} {problem.Subject}"));
            }

            // Anti-vacuity: an empty plan would pass without checking anything.
            Assert.True(plan.Boards.Count >= 10, $"Only {plan.Boards.Count} shipped boards were converted.");
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
                    BoardId = folder,
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

        // ###########################################################################################
        // *** EVERY KiCad FILE A SHIPPED BOARD'S SUBMISSION CARRIES PASSES FINALISE'S CONTENT CHECK
        // (code review, 2026-10-25). *** The test above reads only the files rows CITE; a board's KiCad
        // data is cited by none, and travels in every submission of it (SubmissionKiCadFiles.Collect).
        // The stragglers the converted trees keep there (KiCad-traces.json, a legacy .sch) failed the
        // check finalise runs - SubmissionContentRules.MatchesWholeFile, called here as finalise calls
        // it - so no submission of those boards could pass.
        // ###########################################################################################
        [Fact]
        public void Every_KiCad_file_a_shipped_boards_submission_carries_passes_the_content_check_finalise_runs()
        {
            string dataRoot = SubmissionRulesShippedDataTests.DataRoot();
            DataTreeConversionPlan plan = ClassicRepairToolbox.Tests.DataTreeConversionShippedDataTests.Planned.Value;
            var refusals = new List<string>();
            int checkedFiles = 0;

            foreach (ConvertedBoard converted in plan.Boards)
            {
                foreach (string name in SubmissionKiCadFiles.Collect(converted.Document))
                {
                    string full = Path.Combine(dataRoot, converted.BoardId.Replace('/', Path.DirectorySeparatorChar), name.Replace('/', Path.DirectorySeparatorChar));

                    if (!File.Exists(full))
                    {
                        refusals.Add($"{converted.BoardId}: [{name}] is listed but not found at {full}");
                        continue;
                    }

                    byte[] head = new byte[SubmissionContentRules.SampleBytes];
                    int length;

                    using (FileStream stream = File.OpenRead(full))
                        length = stream.ReadAtLeast(head, head.Length, throwOnEndOfStream: false);

                    checkedFiles++;

                    if (!SubmissionContentRules.MatchesWholeFile(name, head.AsSpan(0, length), () => new StreamReader(full), out string reason))
                        refusals.Add($"{converted.BoardId}: [{name}] {reason}");
                }
            }

            // Anti-vacuity: the shipped tree has KiCad projects, so real files were checked.
            Assert.True(checkedFiles > 10, $"Only {checkedFiles} KiCad files were checked.");
            Assert.Empty(refusals);
        }

        // ###########################################################################################
        // *** EVERY SHIPPED KiCad SCHEMATIC NAMES ITS SHEETS INSIDE ITS "KiCad data" FOLDER
        // (security review, 2026-10-11). *** KiCadSheetReferences refuses a schematic naming a sheet
        // anywhere else, or one that is not a .kicad_sch (code review, 2026-10-11), and a board's
        // KiCad data travels in its submissions - so a rule refusing a published project would make
        // that board unsubmittable. Every .kicad_sch in the tree, whether a row cites it or not.
        // ###########################################################################################
        [Fact]
        public void Every_shipped_KiCad_schematic_names_its_sheets_inside_its_KiCad_folder()
        {
            string dataRoot = SubmissionRulesShippedDataTests.DataRoot();
            var refusals = new List<string>();
            int sheetFiles = 0;

            foreach (string full in Directory.EnumerateFiles(dataRoot, "*.kicad_sch", SearchOption.AllDirectories))
            {
                string path = Path.GetRelativePath(dataRoot, full).Replace(Path.DirectorySeparatorChar, '/');
                string content = File.ReadAllText(full);

                sheetFiles += KiCadSheetReferences.Find(new StringReader(content)).Count();

                if (!KiCadSheetReferences.Check(path, new StringReader(content), out string reason))
                    refusals.Add($"{path}: {reason}");
            }

            // The shipped tree has hierarchical projects, so the rule is exercised on real sheets.
            Assert.True(sheetFiles > 0, "No shipped KiCad schematic names a child sheet.");
            Assert.Empty(refusals);
        }
    }
}
