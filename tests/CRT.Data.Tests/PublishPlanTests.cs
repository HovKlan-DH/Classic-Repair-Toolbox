using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Covers PublishPlan - what a publish will write, decided before anything is written.
//
// WHY THESE TESTS MATTER MORE THAN MOST: publishing changes what every user of a board
// downloads, and the project owner decided against retained revisions, so it cannot be undone. Every
// refusal here is a thing that would otherwise be discovered halfway through replacing a board.
//
// In the CRT 4.0 tree (2026-10-25) a publish writes two kinds of file, both at places the SERVER
// decides: the board file "<board id>/board.json", and each file's bytes stored once under their
// checksum, "Files/<ab>/<sha256>.<ext>". A submitted file's path is the board's NAME for it.
public sealed class PublishPlanTests
{
    private const string Root = "/srv/beta";

    private static readonly string HashA = new('a', 64);
    private static readonly string HashB = new('b', 64);

    // ###########################################################################################
    // A manifest whose ROWS cite every file it carries (security review, 2026-09-25). A publish
    // refuses a file no row uses - such a file is shown nowhere and was approved unseen - so each
    // file is cited as a board-level file, the simplest row that can name any file type.
    // ###########################################################################################
    private static SubmissionManifest Manifest(params SubmissionFile[] files)
    {
        var manifest = new SubmissionManifest
        {
            BoardId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Files = [.. files]
        };

        foreach (SubmissionFile file in files)
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = file.Path });

        return manifest;
    }

    private static SubmissionFile File(string name, string? hash = null, long size = 1024) => new()
    {
        Path = name,
        Sha256 = hash ?? PublishPlanTests.HashA,
        SizeBytes = size
    };

    private static PublishPlanResult Build(
        SubmissionManifest manifest,
        string revision = "r2",
        PublishedTreeView? board = null,
        IEnumerable<string>? stored = null)
    {
        var held = new HashSet<string>(stored ?? [], StringComparer.Ordinal);

        return PublishPlan.Build(
            PublishPlanTests.Root,
            manifest,
            revision,
            new DateTimeOffset(2026, 9, 21, 12, 0, 0, TimeSpan.Zero),
            ["Someone"],
            BoardDescriptorRules.BoardOrigin.Contributed,
            board,
            held.Contains);
    }

    private static bool HasCode(PublishPlanResult result, string code) =>
        result.Problems.Any(problem => problem.Code == code);

    // A published board listing these names (name -> hash), in a tree holding these folders.
    private static PublishedTreeView Board(IReadOnlyDictionary<string, string>? listed = null, params string[] folders) =>
        new(
            name => listed is not null && listed.TryGetValue(name, out string? hash) ? hash : null,
            folder =>
            {
                string prefix = folder.Length == 0 ? string.Empty : folder + "/";

                List<string> entries = folders
                    .Where(path => path.StartsWith(prefix, StringComparison.Ordinal))
                    .Select(path => path[prefix.Length..].Split('/')[0])
                    .Where(entry => entry.Length > 0)
                    .Distinct(StringComparer.Ordinal)
                    .ToList();

                return entries.Count == 0 && folder.Length > 0 ? null : entries;
            });

    // -----------------------------------------------------------------------------------
    // Where it writes
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_board_file_is_the_boards_own_board_json_under_the_data_root()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("Sheet1.png")));

        Assert.True(result.IsPlanned, string.Join(" ", result.Problems.Select(problem => problem.Message)));
        Assert.Equal(
            Path.GetFullPath(Path.Combine(PublishPlanTests.Root, "Commodore", "C64", "250407", "board.json")),
            result.Plan!.BoardFilePath);
        Assert.Equal("Commodore/C64/250407/board.json", result.Plan.BoardFileRelative);
        Assert.Equal(Path.GetDirectoryName(result.Plan.BoardFilePath), result.Plan.BoardFolder);
    }

    [Fact]
    public void A_file_is_stored_under_its_checksum_never_under_its_name()
    {
        // The whole point of the 4.0 tree's security: a submission never names a place to write.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("Images/U8.PNG")));

        PlannedFile file = Assert.Single(result.Plan!.Files);

        Assert.Equal("Images/U8.PNG", file.Name);
        Assert.Equal($"Files/aa/{PublishPlanTests.HashA}.png", file.StoredPath);
        Assert.Equal(Path.GetFullPath(Path.Combine(PublishPlanTests.Root, "Files", "aa", PublishPlanTests.HashA + ".png")), file.AbsolutePath);
    }

    [Fact]
    public void Two_names_of_the_same_bytes_are_stored_once()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File("Images/a.png"),
            PublishPlanTests.File("Images/b.png")));

        Assert.Equal(2, result.Plan!.Files.Count);
        Assert.Single(result.Plan.StoredFiles);
        Assert.Equal(1024, result.Plan.TotalFileBytes);
    }

    [Fact]
    public void Bytes_the_tree_already_holds_are_listed_but_never_written()
    {
        // A one-cell typo fix on a board of 1,200 files stores nothing at all.
        string stored = $"Files/aa/{PublishPlanTests.HashA}.png";

        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File("Images/u8.png")),
            stored: [stored]);

        Assert.True(result.IsPlanned);
        Assert.Empty(result.Plan!.Files);
        Assert.Equal("Images/u8.png", Assert.Single(result.Plan.UnchangedFiles).Name);
        Assert.Equal("Images/u8.png", Assert.Single(result.Plan.FileList).Name);
    }

    [Fact]
    public void Leaving_stored_bytes_alone_does_not_change_the_content_hash()
    {
        // Written or left alone, the file is part of the board, so the content hash - what decides
        // whether the board waits for production - must not depend on which it was.
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File("Images/u8.png"));

        PublishPlanResult unchanged = PublishPlanTests.Build(manifest, stored: [$"Files/aa/{PublishPlanTests.HashA}.png"]);
        PublishPlanResult written = PublishPlanTests.Build(manifest);

        Assert.Empty(unchanged.Plan!.Files);
        Assert.Single(written.Plan!.Files);
        Assert.Equal(written.Plan.Descriptor.ContentHash, unchanged.Plan.Descriptor.ContentHash);
    }

    [Fact]
    public void The_file_list_names_every_file_with_its_bytes_in_name_order()
    {
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(
                PublishPlanTests.File("b.png", PublishPlanTests.HashB, 7),
                PublishPlanTests.File("A.png", PublishPlanTests.HashA, 9)),
            stored: [$"Files/bb/{PublishPlanTests.HashB}.png"]);

        Assert.Collection(
            result.Plan!.FileList,
            file => Assert.Equal(("A.png", PublishPlanTests.HashA, 9L), (file.Name, file.Sha256, file.Size)),
            file => Assert.Equal(("b.png", PublishPlanTests.HashB, 7L), (file.Name, file.Sha256, file.Size)));
    }

    // -----------------------------------------------------------------------------------
    // The KiCad data (2026-09-26)
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE BOARD'S KiCad DATA IS PLANNED ALTHOUGH NO ROW CITES IT. *** PublishPlan re-runs the
    // create-time file rules as the last gate before the tree, so the same exemption must hold
    // here: a KiCad file under "KiCad data/" goes in, and one anywhere else is refused.
    // ###########################################################################################
    [Fact]
    public void The_boards_own_KiCad_data_is_planned_although_no_row_cites_it()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File("Sheet1.png"));
        manifest.Files.Add(PublishPlanTests.File("KiCad data/board.kicad_pcb", PublishPlanTests.HashB));

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.True(result.IsPlanned, string.Join("; ", result.Problems.Select(problem => problem.Code)));
        Assert.Contains(result.Plan!.Files, file => file.Name == "KiCad data/board.kicad_pcb");
    }

    [Fact]
    public void A_KiCad_file_outside_its_folder_is_refused_at_the_last_gate_too()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File("Sheet1.png"));
        manifest.Files.Add(PublishPlanTests.File("board.kicad_pcb", PublishPlanTests.HashB));

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.True(PublishPlanTests.HasCode(result, "file.type-refused"));
    }

    // A file nothing reads (a generated report) that the converted board already has stays with it
    // for as long as nobody changes it - and cannot be brought in new.
    [Fact]
    public void A_KiCad_folder_straggler_stays_only_as_the_published_board_has_it()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File("Sheet1.png"));
        manifest.Files.Add(PublishPlanTests.File("KiCad data/KiCad-traces.json", PublishPlanTests.HashB));

        PublishedTreeView same = PublishPlanTests.Board(new Dictionary<string, string> { ["KiCad data/KiCad-traces.json"] = PublishPlanTests.HashB });
        PublishedTreeView other = PublishPlanTests.Board(new Dictionary<string, string> { ["KiCad data/KiCad-traces.json"] = PublishPlanTests.HashA });

        Assert.True(PublishPlanTests.Build(manifest, board: same).IsPlanned);
        Assert.True(PublishPlanTests.HasCode(PublishPlanTests.Build(manifest, board: other), "file.type-refused"));
        Assert.True(PublishPlanTests.HasCode(PublishPlanTests.Build(manifest), "file.type-refused"));
    }

    // -----------------------------------------------------------------------------------
    // What is refused
    // -----------------------------------------------------------------------------------

    [Theory]
    [InlineData("../escape.png")]
    [InlineData("Images/../../escape.png")]
    [InlineData("/etc/passwd.png")]
    // windows-path-literal: a hostile name shaped like a drive path, refused on every platform
    [InlineData("C:/Windows/x.png")]
    public void A_name_escaping_its_folder_is_refused(string name)
    {
        // A name is where CRT puts the file in a draft and a KiCad copy, so it is still a path.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File(name)));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.path-refused"));
    }

    [Fact]
    public void A_file_with_an_unusable_hash_is_refused()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png", "not-a-hash")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.hash-invalid"));
    }

    [Fact]
    public void The_same_name_twice_is_refused()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png"), PublishPlanTests.File("a.png")));

        Assert.True(PublishPlanTests.HasCode(result, "file.duplicate"));
    }

    [Fact]
    public void Two_names_differing_only_in_case_are_refused_and_BOTH_are_named()
    {
        // One file on Windows and macOS, where CRT puts a draft's files under their names.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File("Images/U8.png"),
            PublishPlanTests.File("images/u8.png", PublishPlanTests.HashB)));

        ValidationFinding problem = Assert.Single(result.Problems, finding => finding.Code == "file.case-collision");

        Assert.Contains("Images/U8.png", problem.Message, StringComparison.Ordinal);
        Assert.Contains("images/u8.png", problem.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_board_id_that_disagrees_with_its_own_names_is_refused()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest();
        manifest.BoardId = "Commodore/C64/250425";

        Assert.True(PublishPlanTests.HasCode(PublishPlanTests.Build(manifest), "board.id-mismatch"));
    }

    [Fact]
    public void A_malformed_board_id_is_refused()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest();
        manifest.BoardId = "Commodore/../etc";

        Assert.True(PublishPlanTests.HasCode(PublishPlanTests.Build(manifest), "board.id-invalid"));
    }

    [Theory]
    [InlineData("Files")]
    [InlineData("MiniPro")]
    [InlineData("files")]
    public void A_manufacturer_named_as_one_of_the_trees_own_folders_is_refused(string manufacturer)
    {
        // A board there would sit among the stored files or the IC tests.
        SubmissionManifest manifest = PublishPlanTests.Manifest();
        manifest.Manufacturer = manufacturer;
        manifest.BoardId = $"{manufacturer}/C64/250407";

        Assert.True(PublishPlanTests.HasCode(PublishPlanTests.Build(manifest), "board.reserved-folder"));
    }

    [Fact]
    public void A_publish_with_no_revision_is_refused()
    {
        Assert.True(PublishPlanTests.HasCode(PublishPlanTests.Build(PublishPlanTests.Manifest(), revision: " "), "publish.no-revision"));
    }

    [Fact]
    public void A_file_NO_ROW_USES_is_refused()
    {
        // Nothing would show it, so no maintainer could have seen it - which is exactly how a file
        // that should never be published gets published.
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File("a.png"));
        manifest.Rows.BoardLocalFiles.Clear();

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.not-used"));
    }

    [Theory]
    [InlineData(".htaccess")]
    [InlineData("Images/.hidden/x.png")]
    [InlineData("run.exe")]
    [InlineData("drawing.svg")]
    [InlineData("board.json")]
    [InlineData("Data C64 250407.xlsx")]
    [InlineData("README")]
    public void A_file_of_a_type_board_data_never_uses_is_refused(string name)
    {
        // A .json or a workbook in particular: the board file is the publish's, from the rows.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File(name)));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.type-refused"));
    }

    [Fact]
    public void A_NEW_BOARD_whose_folder_differs_from_a_published_one_only_by_case_is_refused()
    {
        // "commodore/C64/250407" beside the real "Commodore/C64/250407": on a client the two are
        // one folder, so the "new" board's file would replace the real board's there.
        SubmissionManifest manifest = PublishPlanTests.Manifest();
        manifest.BoardId = "commodore/C64/250407";
        manifest.Manufacturer = "commodore";

        PublishPlanResult result = PublishPlanTests.Build(
            manifest,
            board: PublishPlanTests.Board(folders: "Commodore/C64/250407/board.json"));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "board.case-collision"));
    }

    [Fact]
    public void An_EXACT_published_folder_is_not_a_case_collision()
    {
        // Anti-vacuity: publishing over the board's own published file is the ordinary case.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File("u8.png")),
            board: PublishPlanTests.Board(folders: "Commodore/C64/250407/board.json"));

        Assert.True(result.IsPlanned);
    }

    [Fact]
    public void Every_problem_is_reported_at_once_rather_than_one_per_attempt()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File("../escape.png"),
            PublishPlanTests.File("a.png", "bad"),
            PublishPlanTests.File("run.exe")));

        Assert.True(result.Problems.Count >= 3);
    }

    [Fact]
    public void A_refused_plan_carries_no_plan_at_all()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("../escape.png")));

        Assert.False(result.IsPlanned);
        Assert.Null(result.Plan);
    }

    [Fact]
    public void Every_refusal_carries_a_code_and_a_subject()
    {
        // Codes follow SubmissionValidator's "subject.problem" convention so a caller can group
        // or translate a refusal without parsing prose.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("../escape.png")));

        Assert.NotEmpty(result.Problems);

        foreach (ValidationFinding problem in result.Problems)
        {
            Assert.NotEmpty(problem.Code);
            Assert.Contains('.', problem.Code);
            Assert.NotEmpty(problem.Message);
            Assert.Equal(ValidationSeverity.Error, problem.Severity);
        }
    }

    // -----------------------------------------------------------------------------------
    // The descriptor and its content hash
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_descriptor_carries_the_boards_identity_and_the_revision()
    {
        BoardDescriptor descriptor = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png"))).Plan!.Descriptor;

        Assert.Equal("Commodore/C64/250407", descriptor.BoardId);
        Assert.Equal("r2", descriptor.Revision);
        Assert.Equal(BoardDescriptorRules.BoardOrigin.Contributed, descriptor.Origin);
    }

    [Fact]
    public void The_content_hash_does_not_depend_on_the_order_files_arrive_in()
    {
        PublishPlanResult one = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png"), PublishPlanTests.File("b.png", PublishPlanTests.HashB)));
        PublishPlanResult two = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("b.png", PublishPlanTests.HashB), PublishPlanTests.File("a.png")));

        Assert.Equal(one.Plan!.Descriptor.ContentHash, two.Plan!.Descriptor.ContentHash);
    }

    [Fact]
    public void A_changed_file_moves_the_content_hash()
    {
        PublishPlanResult before = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png")));
        PublishPlanResult after = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png", PublishPlanTests.HashB)));

        Assert.NotEqual(before.Plan!.Descriptor.ContentHash, after.Plan!.Descriptor.ContentHash);
    }

    [Fact]
    public void A_changed_revision_moves_the_content_hash_even_with_identical_files()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File("a.png"));

        Assert.NotEqual(
            PublishPlanTests.Build(manifest, revision: "r2").Plan!.Descriptor.ContentHash,
            PublishPlanTests.Build(manifest, revision: "r3").Plan!.Descriptor.ContentHash);
    }

    [Fact]
    public void A_publish_with_no_files_at_all_is_still_planned()
    {
        // A board of rows alone - a new board with no pictures yet - is a board.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest());

        Assert.True(result.IsPlanned);
        Assert.Empty(result.Plan!.FileList);
    }

    // ###########################################################################################
    // *** THE BOARD FILE MUST BE FOLDED IN. *** A rows-only change - the commonest contribution -
    // stores nothing, so without the board file's hash the content hash would not move and the
    // promotion, which compares exactly this hash, would never offer the board. The board file
    // holds the highlights too, so a highlight-only change moves it the same way.
    // ###########################################################################################
    [Fact]
    public void A_rows_only_change_moves_the_content_hash_once_the_board_file_is_folded_in()
    {
        PublishPlanDetail plan = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png"))).Plan!;

        Assert.NotEqual(
            plan.DescriptorWithBoardFile(new string('1', 64)).ContentHash,
            plan.DescriptorWithBoardFile(new string('2', 64)).ContentHash);
    }

    [Fact]
    public void Folding_the_board_file_in_leaves_every_other_descriptor_field_alone()
    {
        PublishPlanDetail plan = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png"))).Plan!;
        BoardDescriptor folded = plan.DescriptorWithBoardFile(new string('1', 64));

        Assert.Equal(plan.Descriptor.BoardId, folded.BoardId);
        Assert.Equal(plan.Descriptor.Revision, folded.Revision);
        Assert.Equal(plan.Descriptor.PublishedUtc, folded.PublishedUtc);
        Assert.Equal(plan.Descriptor.Maintainers, folded.Maintainers);
        Assert.Equal(plan.Descriptor.Origin, folded.Origin);
        Assert.NotEqual(plan.Descriptor.ContentHash, folded.ContentHash);
    }

    [Fact]
    public void Folding_the_board_file_in_does_not_mutate_the_plan()
    {
        PublishPlanDetail plan = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File("a.png"))).Plan!;
        string before = plan.Descriptor.ContentHash;

        _ = plan.DescriptorWithBoardFile(new string('1', 64));

        Assert.Equal(before, plan.Descriptor.ContentHash);
    }

    [Theory]
    [InlineData("")]
    [InlineData("not-a-hash")]
    public void A_missing_board_file_hash_is_refused_rather_than_hashed_as_blank(string hash)
    {
        PublishPlanDetail plan = PublishPlanTests.Build(PublishPlanTests.Manifest()).Plan!;

        Assert.Throws<ArgumentException>(() => plan.DescriptorWithBoardFile(hash));
    }

    [Fact]
    public void The_plan_carries_the_data_root_so_the_writer_can_check_for_links()
    {
        Assert.Equal(PublishPlanTests.Root, PublishPlanTests.Build(PublishPlanTests.Manifest()).Plan!.DataRoot);
    }
}
