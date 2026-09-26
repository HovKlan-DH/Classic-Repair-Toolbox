using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Covers PublishPlan - what a publish will write, decided before anything is written.
//
// WHY THESE TESTS MATTER MORE THAN MOST: publishing changes what every user of a system
// downloads, and the project owner decided against retained revisions, so it cannot be undone. Every
// refusal here is a thing that would otherwise be discovered halfway through replacing a board.
//
// The one that carries the most weight is the GENERATION guard. Older workbook generations are
// frozen compatibility targets - "Classic-Repair-Toolbox.xlsx" still serves every app build
// before 2.0.0 - and writing one does not fail loudly. It succeeds, and older builds get data
// they cannot read.
public sealed class PublishPlanTests
{
    private const string Root = "/srv/beta";
    private const string SystemFolder = "/srv/beta/Commodore/C64/250407";
    private const string Stem = "Data C64 250407";

    // The board's own folder, relative to the data root - the shape every real submitted path has.
    private const string Own = "Commodore/C64/250407/";

    // ###########################################################################################
    // A manifest whose ROWS cite every file it carries (security review, 2026-09-25).
    //
    // A publish now refuses a file no row uses - such a file is shown nowhere and was approved
    // unseen - so a fixture carrying files and no rows no longer describes a real submission. Each
    // file is cited as a board-level file, the simplest row that can name any file type.
    // ###########################################################################################
    private static SubmissionManifest Manifest(params SubmissionFile[] files)
    {
        var manifest = new SubmissionManifest
        {
            SystemId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Files = [.. files]
        };

        foreach (SubmissionFile file in files)
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = file.Path });

        return manifest;
    }

    private static SubmissionFile File(string path, string? hash = null, long size = 1024) => new()
    {
        Path = path,
        Sha256 = hash ?? new string('a', 64),
        SizeBytes = size
    };

    private static PublishPlanResult Build(
        SubmissionManifest manifest,
        IEnumerable<string>? existing = null,
        string stem = PublishPlanTests.Stem,
        string revision = "r2",
        PublishedTreeView? tree = null) =>
        PublishPlan.Build(
            PublishPlanTests.Root,
            PublishPlanTests.SystemFolder,
            existing ?? ["Data C64 250407 v2.0.0.xlsx"],
            stem,
            manifest,
            revision,
            new DateTimeOffset(2026, 9, 21, 12, 0, 0, TimeSpan.Zero),
            ["Someone"],
            SystemDescriptorRules.SystemOrigin.Contributed,
            tree);

    private static bool HasCode(PublishPlanResult result, string code) =>
        result.Problems.Any(problem => problem.Code == code);

    // ###########################################################################################
    // *** THE BOARD'S KiCad DATA IS PLANNED FOR WRITING (2026-09-26). *** PublishPlan re-runs the
    // create-time file rules as the last gate before the tree, so the same exemption must hold
    // here: a KiCad file in the board's own folder is written although no row cites it, and one
    // anywhere else is refused - a rule loosened at create only would let a queued submission
    // publish what create no longer accepts, or the other way round.
    // ###########################################################################################
    [Fact]
    public void The_boards_own_KiCad_data_is_planned_although_no_row_cites_it()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest(
            PublishPlanTests.File("Commodore/C64/250407/Sheet1.png"));

        manifest.Files.Add(PublishPlanTests.File("Commodore/C64/250407/KiCad data/board.kicad_pcb", new string('b', 64)));

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.True(result.IsPlanned, string.Join("; ", result.Problems.Select(problem => problem.Code)));
        Assert.Contains(result.Plan!.Files, file => file.RelativePath == "Commodore/C64/250407/KiCad data/board.kicad_pcb");
    }

    [Fact]
    public void A_KiCad_file_outside_its_folder_is_refused_at_the_last_gate_too()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest(
            PublishPlanTests.File("Commodore/C64/250407/Sheet1.png"));

        manifest.Files.Add(PublishPlanTests.File("Commodore/C64/250407/board.kicad_pcb", new string('b', 64)));

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.Contains(result.Problems, problem => problem.Code == "file.type-refused");
    }

    // -----------------------------------------------------------------------------------
    // The generation guard
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_workbook_is_named_for_the_newest_generation_present()
    {
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(),
            existing: ["Data C64 250407.xlsx", "Data C64 250407 v2.0.0.xlsx"]);

        Assert.True(result.IsPlanned);
        Assert.Equal("Data C64 250407 v2.0.0.xlsx", result.Plan!.WorkbookFileName);
        Assert.Equal(Version.Parse("2.0.0"), result.Plan.TargetGeneration);
    }

    [Fact]
    public void A_newer_generation_in_the_tree_is_targeted_without_any_configuration_change()
    {
        // The project owner's rule: "always use the newest version". Adding a 3.0.0 file to the tree
        // is the ONLY action needed - nothing here is configured.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(),
            existing: ["Data C64 250407.xlsx", "Data C64 250407 v2.0.0.xlsx", "Data C64 250407 v3.0.0.xlsx"]);

        Assert.True(result.IsPlanned);
        Assert.Equal("Data C64 250407 v3.0.0.xlsx", result.Plan!.WorkbookFileName);
    }

    [Fact]
    public void A_tree_with_only_the_unversioned_original_publishes_into_it()
    {
        // The original IS a generation - the first one. A tree that has never been versioned is
        // publishing into its own newest, which is correct.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(),
            existing: ["Data C64 250407.xlsx"]);

        Assert.True(result.IsPlanned);
        Assert.Equal("Data C64 250407.xlsx", result.Plan!.WorkbookFileName);
        Assert.Null(result.Plan.TargetGeneration);
    }

    [Fact]
    public void A_brand_new_system_with_an_empty_folder_is_REFUSED_when_the_tree_has_no_generation()
    {
        // *** THIS TEST USED TO ASSERT THE OPPOSITE, AND THE OPPOSITE WAS A BUG. *** It read
        // "A_brand_new_system_with_an_empty_folder_publishes_unversioned" and pinned exactly that
        // - nothing in the folder means no generation, so the board publishes as
        // "Data C64 250407.xlsx".
        //
        // That file is the FROZEN generation serving every application build before 2.0.0, and the
        // project owner's rule is that it is never written. So every new system - the highest-value
        // kind of contribution there is - would have been published into the one place it must not
        // go, silently, and older builds would have received contributed data they cannot read.
        //
        // Found 2026-09-22 while wiring ApprovePublishFlow, whose own tests failed on the path the
        // published board landed at. The generation now comes from the TREE's master workbooks
        // when the system's own folder is empty; with no versioned master anywhere, publishing is
        // refused rather than guessed at.
        //
        // This Build() helper points at a temp root with no master workbooks in it, which is the
        // "no generation anywhere" case.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(), existing: []);

        Assert.False(result.IsPlanned);
        Assert.Contains(result.Problems, problem => problem.Code == "publish.no-generation");
    }

    [Fact]
    public void An_EXISTING_unversioned_system_is_NOT_treated_as_a_new_one()
    {
        // *** THE DISTINCTION THE FIRST VERSION OF THE NEW-SYSTEM GUARD MISSED. ***
        // ResolveNewestGeneration answers null for BOTH "the folder is empty" and "the folder
        // holds only an unversioned board", and those are opposite situations:
        //
        //   - empty            -> a new system, which must NOT land in the frozen unversioned file;
        //   - unversioned only -> an existing system that legitimately lives there.
        //
        // Conflating them made every unversioned board unpublishable, which the existing test
        // above caught immediately. The folder's EMPTINESS is what separates them, and this pins
        // that the non-empty case still plans.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(),
            existing: ["Data C64 250407.xlsx", "Images/sheet1.png"]);

        Assert.True(result.IsPlanned);
        Assert.Null(result.Plan!.TargetGeneration);
    }

    [Fact]
    public void A_submitted_workbook_is_REFUSED_rather_than_written()
    {
        // *** THE GUARD THAT PROTECTS A FROZEN GENERATION. *** The board workbook is GENERATED
        // from the submitted rows; a submission has no business carrying one. Accepting an
        // uploaded .xlsx would let a contributor overwrite any generation's workbook, including
        // the unversioned original that serves every pre-2.0.0 build, and the write would succeed
        // silently.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Data C64 250407.xlsx")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.workbook-upload"));
    }

    [Fact]
    public void A_workbook_hidden_in_a_subfolder_is_refused_too()
    {
        // The rule is about the extension anywhere in the submission, not about a name matching a
        // generation - see PublishPlan.LooksLikeWorkbook for why the broad rule was chosen.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Files/notes.xlsx")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.workbook-upload"));
    }

    // -----------------------------------------------------------------------------------
    // Path safety
    // -----------------------------------------------------------------------------------

    [Theory]
    [InlineData("../../evil.png")]
    [InlineData("Images/./../../evil.png")]
    [InlineData("/etc/passwd")]
    public void A_path_escaping_the_system_folder_is_refused(string path)
    {
        // Delegated to SubmissionPathRules, which resolves and then checks containment. Asserted
        // here anyway: this is the last gate before bytes land, and a caller that forgot to check
        // the result is the bug that already shipped once in the upload loop.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(path)));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.path-refused"));
    }

    [Fact]
    public void A_file_with_an_unusable_hash_is_refused()
    {
        // The hash is how the blob is found. Without a usable one there is nothing to copy, and
        // "content-addressed" stops meaning anything.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png", hash: "not-a-hash")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.hash-invalid"));
    }

    [Fact]
    public void The_same_path_twice_is_refused()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png"),
            PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.duplicate"));
    }

    [Fact]
    public void Two_paths_differing_only_in_case_are_refused_and_BOTH_are_named()
    {
        // Both can exist on the Linux server and only one on a Windows client, so the published
        // tree would be un-syncable for much of the audience. The message must name both
        // spellings - "already exists" about a file you can plainly see is baffling.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File(PublishPlanTests.Own + "Images/U8.png"),
            PublishPlanTests.File(PublishPlanTests.Own + "Images/u8.PNG")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.case-collision"));

        string message = result.Problems.First(problem => problem.Code == "file.case-collision").Message;
        Assert.Contains("Images/u8.PNG", message);
        Assert.Contains("Images/U8.png", message);
    }

    // -----------------------------------------------------------------------------------
    // Identity
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_system_id_that_disagrees_with_its_own_names_is_refused()
    {
        // The id is a database primary key AND resolves a folder. A mismatch means everything
        // keyed off it disagrees with what is printed on screen.
        SubmissionManifest manifest = PublishPlanTests.Manifest();
        manifest.Board = "250425";

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "system.id-mismatch"));
    }

    [Fact]
    public void A_malformed_system_id_is_refused()
    {
        SubmissionManifest manifest = PublishPlanTests.Manifest();
        manifest.SystemId = "../../etc";

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "system.id-invalid"));
    }

    [Fact]
    public void A_publish_with_no_revision_is_refused()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(), revision: "");

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "publish.no-revision"));
    }

    [Fact]
    public void A_publish_with_no_board_stem_is_refused()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(), stem: "");

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "publish.no-workbook-name"));
    }

    // -----------------------------------------------------------------------------------
    // The plan itself
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** A FILE PATH IS RESOLVED AGAINST THE DATA ROOT, NOT THE SYSTEM FOLDER (corrected
    // 2026-09-23). ***
    //
    // This test used to pass "Images/u8.png" and assert the result landed under the SYSTEM folder,
    // and it passed - because those paths are not the shape a real submission carries. A real one
    // is already data-root-relative ("Commodore/C64/250407/Sheet1.png"), so resolving it against
    // the system folder wrote it to "<root>/Commodore/C64/250407/Commodore/C64/250407/Sheet1.png".
    //
    // The first real publish duplicated 1,215 files into the board folder that way - the whole
    // board nested inside itself - and every client then re-downloaded the lot, which is how the
    // project owner noticed.
    //
    // The unrealistic input is what let it hide, so the input is corrected here too.
    // ###########################################################################################
    [Fact]
    public void A_planned_publish_resolves_a_file_to_its_DATA_ROOT_relative_location()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File("Commodore/C64/250407/Sheet1.png"),
            PublishPlanTests.File("Commodore/C64/250407/Files/manual.pdf")));

        Assert.True(result.IsPlanned);
        Assert.Equal(2, result.Plan!.Files.Count);

        foreach (PlannedFile file in result.Plan.Files)
        {
            Assert.True(Path.IsPathRooted(file.AbsolutePath));

            // Inside the tree, and resolved by SubmissionPathRules rather than re-joined by the
            // caller - a re-join reintroduces the traversal hole.
            Assert.StartsWith(
                Path.GetFullPath(PublishPlanTests.Root),
                Path.GetFullPath(file.AbsolutePath),
                StringComparison.Ordinal);

            // ###########################################################################################
            // *** AND THE SYSTEM SEGMENTS APPEAR EXACTLY ONCE. *** This is what fails against the
            // version that shipped, which produced
            // "/srv/beta/Commodore/C64/250407/Commodore/C64/250407/Sheet1.png".
            //
            // Counted WITHOUT surrounding slashes, deliberately. In that doubled path the two
            // occurrences share the slash between them, so searching for "/Commodore/C64/250407/"
            // matches only once and the assertion passes against the bug - which is precisely what
            // the first version of this test did.
            // ###########################################################################################
            string full = Path.GetFullPath(file.AbsolutePath).Replace('\\', '/');

            Assert.Equal(
                1,
                full.Split("Commodore/C64/250407", StringSplitOptions.None).Length - 1);
        }
    }

    // ###########################################################################################
    // *** A SHARED FILE LIVES OUTSIDE THE SYSTEM FOLDER, and the old base could never write it. ***
    //
    // "Commodore/Shared files/Component images/6526.png" belongs beside the MANUFACTURER, cited by
    // boards across it. Containing a publish to one board's folder made that impossible to express,
    // which is the second half of why the base was wrong rather than merely off by a prefix.
    // ###########################################################################################
    [Fact]
    public void A_manufacturer_SHARED_file_is_published_beside_the_manufacturer()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File("Commodore/Shared files/Component images/6526.png")));

        Assert.True(result.IsPlanned);

        PlannedFile file = Assert.Single(result.Plan!.Files);
        string full = Path.GetFullPath(file.AbsolutePath).Replace('\\', '/');

        Assert.EndsWith("/Commodore/Shared files/Component images/6526.png", full, StringComparison.Ordinal);

        // Emphatically NOT inside the board being published.
        Assert.DoesNotContain("/250407/", full, StringComparison.Ordinal);
    }

    // The containment that matters is still enforced - it is simply measured from the data root.
    [Fact]
    public void A_traversal_out_of_the_DATA_ROOT_is_still_refused()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File("../../etc/passwd")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.path-refused"));
    }

    [Fact]
    public void The_relative_path_is_kept_alongside_the_absolute_one()
    {
        // The descriptor's content hash is computed over RELATIVE paths. Folding the server's own
        // directory layout into a hash every client compares against would make every client
        // permanently out of date.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/u8.png")));

        PlannedFile file = Assert.Single(result.Plan!.Files);
        Assert.Equal(PublishPlanTests.Own + "Images/u8.png", file.RelativePath);
    }

    [Fact]
    public void The_descriptor_carries_the_systems_identity_and_the_revision()
    {
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/u8.png")));

        SystemDescriptor descriptor = result.Plan!.Descriptor;

        Assert.Equal("Commodore/C64/250407", descriptor.SystemId);
        Assert.Equal("Commodore", descriptor.Manufacturer);
        Assert.Equal("C64", descriptor.Hardware);
        Assert.Equal("250407", descriptor.Board);
        Assert.Equal("r2", descriptor.Revision);
        Assert.Equal(SystemDescriptorRules.SystemOrigin.Contributed, descriptor.Origin);
        Assert.Equal(["Someone"], descriptor.Maintainers);
        Assert.NotEmpty(descriptor.ContentHash);
    }

    [Fact]
    public void The_content_hash_does_not_depend_on_the_order_files_arrive_in()
    {
        // A directory walk or a database query returns whatever order it likes; two machines
        // hashing the same system must agree, or every client reads as permanently stale.
        PublishPlanResult forwards = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png", new string('1', 64)),
            PublishPlanTests.File(PublishPlanTests.Own + "Images/b.png", new string('2', 64))));

        PublishPlanResult backwards = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File(PublishPlanTests.Own + "Images/b.png", new string('2', 64)),
            PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png", new string('1', 64))));

        Assert.Equal(forwards.Plan!.Descriptor.ContentHash, backwards.Plan!.Descriptor.ContentHash);
    }

    [Fact]
    public void A_changed_file_moves_the_content_hash()
    {
        // Anti-vacuity for the test above: if the hash ignored its inputs, the ordering test
        // would pass while proving nothing.
        PublishPlanResult before = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png", new string('1', 64))));

        PublishPlanResult after = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png", new string('2', 64))));

        Assert.NotEqual(before.Plan!.Descriptor.ContentHash, after.Plan!.Descriptor.ContentHash);
    }

    [Fact]
    public void A_changed_revision_moves_the_content_hash_even_with_identical_files()
    {
        // A metadata-only republish still has to be picked up by clients.
        PublishPlanResult r2 = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png")),
            revision: "r2");

        PublishPlanResult r3 = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png")),
            revision: "r3");

        Assert.NotEqual(r2.Plan!.Descriptor.ContentHash, r3.Plan!.Descriptor.ContentHash);
    }

    [Fact]
    public void A_publish_with_no_files_at_all_is_still_planned()
    {
        // A rows-only change - the typo fix that uploads nothing - is the commonest contribution
        // there is. It must produce a plan, not a refusal.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest());

        Assert.True(result.IsPlanned);
        Assert.Empty(result.Plan!.Files);
        Assert.Equal(0, result.Plan.TotalFileBytes);
    }

    [Fact]
    public void The_total_bytes_are_summed_across_the_planned_files()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File(PublishPlanTests.Own + "Images/a.png", new string('1', 64), size: 100),
            PublishPlanTests.File(PublishPlanTests.Own + "Images/b.png", new string('2', 64), size: 250)));

        Assert.Equal(350, result.Plan!.TotalFileBytes);
    }

    [Fact]
    public void Every_problem_is_reported_at_once_rather_than_one_per_attempt()
    {
        // A contributor fixing refusals one at a time, with a full upload between each, is how a
        // contribution stops happening.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(
            PublishPlanTests.File("../escape.png"),
            PublishPlanTests.File(PublishPlanTests.Own + "Images/b.png", hash: "nope")));

        Assert.False(result.IsPlanned);
        Assert.Equal(2, result.Problems.Count);
    }

    [Fact]
    public void A_refused_plan_carries_no_plan_at_all()
    {
        // Deliberately not "a plan plus a problems list": a partially valid publish is not a
        // thing that may proceed, and a caller holding both would have to remember to check.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File("../escape.png")));

        Assert.False(result.IsPlanned);
        Assert.Null(result.Plan);
        Assert.NotEmpty(result.Problems);
    }

    // -----------------------------------------------------------------------------------
    // Folding the generated workbook into the content hash
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_rows_only_change_moves_the_content_hash_once_the_workbook_is_folded_in()
    {
        // *** THE SILENT FAILURE DescriptorWithWorkbook EXISTS TO PREVENT. *** A typo fix uploads
        // NO files, so its uploaded-file list is byte-identical to the previous publish's. If the
        // executor wrote the planned descriptor as-is, the content hash would not move and no
        // client would ever re-download the board that just changed.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest());

        string plannedHash = result.Plan!.Descriptor.ContentHash;
        string sidecar = new('9', 64);

        string withOldWorkbook = result.Plan.DescriptorWithWorkbook(new string('1', 64), sidecar).ContentHash;
        string withNewWorkbook = result.Plan.DescriptorWithWorkbook(new string('2', 64), sidecar).ContentHash;

        Assert.NotEqual(plannedHash, withOldWorkbook);
        Assert.NotEqual(withOldWorkbook, withNewWorkbook);
    }

    [Fact]
    public void A_HIGHLIGHT_ONLY_change_moves_the_content_hash_through_the_SIDECAR()
    {
        // *** THE SAME SILENT FAILURE, ONE FILE OVER, and the sidecar is the ONLY thing that can
        // catch it. *** A submission that merely MOVES A HIGHLIGHT changes no workbook row - the
        // highlights do not live in the workbook at all - and uploads no file. So the workbook
        // hash does not move either, and without the sidecar folded in the content hash would be
        // identical to the previous publish's. No client would re-download the board, and the
        // moved highlight would never reach anybody.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest());

        string workbook = new('1', 64);

        string before = result.Plan!.DescriptorWithWorkbook(workbook, new string('a', 64)).ContentHash;
        string after = result.Plan.DescriptorWithWorkbook(workbook, new string('b', 64)).ContentHash;

        Assert.NotEqual(before, after);
    }

    [Fact]
    public void Folding_the_workbook_in_leaves_every_other_descriptor_field_alone()
    {
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/u8.png")));

        SystemDescriptor planned = result.Plan!.Descriptor;
        SystemDescriptor final = result.Plan.DescriptorWithWorkbook(new string('c', 64), new string('d', 64));

        Assert.Equal(planned.SystemId, final.SystemId);
        Assert.Equal(planned.Manufacturer, final.Manufacturer);
        Assert.Equal(planned.Hardware, final.Hardware);
        Assert.Equal(planned.Board, final.Board);
        Assert.Equal(planned.Revision, final.Revision);
        Assert.Equal(planned.PublishedUtc, final.PublishedUtc);
        Assert.Equal(planned.Origin, final.Origin);
        Assert.Equal(planned.Maintainers, final.Maintainers);
    }

    [Fact]
    public void Folding_the_workbook_in_does_not_mutate_the_plan()
    {
        // SystemDescriptor is a mutable class, so a plan that quietly changed under a caller
        // holding it would defeat the whole "decide it all up front" design.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/u8.png")));

        string before = result.Plan!.Descriptor.ContentHash;
        result.Plan.DescriptorWithWorkbook(new string('c', 64), new string('d', 64));

        Assert.Equal(before, result.Plan.Descriptor.ContentHash);
    }

    [Fact]
    public void A_missing_workbook_hash_is_refused_rather_than_hashed_as_blank()
    {
        // Accepting a blank would produce a plausible-looking descriptor whose hash silently did
        // not account for the workbook at all - the exact bug this method prevents.
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest());

        string good = new('c', 64);

        Assert.Throws<ArgumentException>(() => result.Plan!.DescriptorWithWorkbook("", good));
        Assert.Throws<ArgumentException>(() => result.Plan!.DescriptorWithWorkbook("not-a-hash", good));

        // The SIDECAR hash is equally required, for the same reason: a blank one produces a
        // plausible descriptor whose hash silently does not account for the board's highlights.
        Assert.Throws<ArgumentException>(() => result.Plan!.DescriptorWithWorkbook(good, ""));
        Assert.Throws<ArgumentException>(() => result.Plan!.DescriptorWithWorkbook(good, "not-a-hash"));
    }

    [Fact]
    public void Every_refusal_carries_a_code_and_a_subject()
    {
        // Codes follow SubmissionValidator's "subject.problem" convention so a caller can group
        // or translate a refusal without parsing prose.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File("../escape.png")));

        foreach (ValidationFinding problem in result.Problems)
        {
            Assert.NotEmpty(problem.Code);
            Assert.Contains('.', problem.Code);
            Assert.NotEmpty(problem.Message);
            Assert.Equal(ValidationSeverity.Error, problem.Severity);
        }
    }

    // -----------------------------------------------------------------------------------
    // WHICH files, and WHOSE (security review, 2026-09-25)
    //
    // The same rules SubmissionFileRules applies at create, applied again at the last gate before
    // an irreversible write - a submission queued weeks ago, or by an older build, must not reach
    // the tree on the strength of a check that ran then.
    // -----------------------------------------------------------------------------------

    // A tree holding the given files (relative path -> hash). Folders are derived from the paths.
    private static PublishedTreeView Tree(params (string Path, string Hash)[] files)
    {
        var hashes = files.ToDictionary(file => file.Path, file => file.Hash, StringComparer.Ordinal);

        return new PublishedTreeView(
            path => hashes.TryGetValue(path, out string? hash) ? hash : null,
            folder =>
            {
                string prefix = folder.Length == 0 ? string.Empty : folder + "/";

                List<string> entries = hashes.Keys
                    .Where(path => path.StartsWith(prefix, StringComparison.Ordinal))
                    .Select(path => path[prefix.Length..].Split('/')[0])
                    .Distinct(StringComparer.Ordinal)
                    .ToList();

                return entries.Count == 0 && folder.Length > 0 ? null : entries;
            });
    }

    [Fact]
    public void A_file_NO_ROW_USES_is_refused()
    {
        // Nothing would show it, so no maintainer could have seen it - which is exactly how a file
        // that should never be published gets published.
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "a.png"));
        manifest.Rows.BoardLocalFiles.Clear();

        PublishPlanResult result = PublishPlanTests.Build(manifest);

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.not-used"));
    }

    [Theory]
    [InlineData("Commodore/C64/250407/.htaccess")]
    [InlineData("Commodore/C64/250407/run.exe")]
    [InlineData("Commodore/C64/250407/drawing.svg")]
    [InlineData("Commodore/C64/250407/system.json")]
    [InlineData("Commodore/C64/250407/README")]
    public void A_file_of_a_type_board_data_never_uses_is_refused(string path)
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest(PublishPlanTests.File(path)));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.type-refused"));
    }

    [Fact]
    public void ANOTHER_boards_file_is_refused_when_the_tree_cannot_be_consulted()
    {
        // Fails CLOSED: without the tree there is no way to know the file is left unchanged.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File("Commodore/C64/250425/Sheet1.png")));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.other-board"));
    }

    [Fact]
    public void ANOTHER_boards_file_that_would_CHANGE_is_refused()
    {
        // The attack this whole section exists for: a submission to one board carrying new bytes
        // for another board's file, which the publish then copied over the real one.
        string foreign = "Commodore/C64/250425/Sheet1.png";

        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(foreign, new string('1', 64))),
            tree: PublishPlanTests.Tree((foreign, new string('2', 64))));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "file.other-board"));
    }

    [Fact]
    public void ANOTHER_boards_file_cited_UNCHANGED_is_planned_but_never_written()
    {
        // Real data does this: the C128DCR 250477 board cites scope-baseline texts that live in
        // the C128 310378 folder. Refusing it would make a published board unsubmittable; writing
        // it would be pointless at best. So it is neither - but it stays part of the content.
        string foreign = "Commodore/C128/310378/Scope baseline/notes.txt";
        string hash = new('3', 64);

        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(foreign, hash)),
            tree: PublishPlanTests.Tree((foreign, hash)));

        Assert.True(result.IsPlanned);
        Assert.Empty(result.Plan!.Files);
        Assert.Equal(foreign, Assert.Single(result.Plan.UnchangedFiles).RelativePath);
    }

    [Fact]
    public void An_unchanged_foreign_file_still_counts_toward_the_content_hash()
    {
        // Leaving it out would make two different systems - one citing it, one not - hash alike.
        string foreign = "Commodore/C128/310378/Scope baseline/notes.txt";
        string hash = new('3', 64);
        PublishedTreeView tree = PublishPlanTests.Tree((foreign, hash));

        PublishPlanResult with = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(foreign, hash)), tree: tree);

        PublishPlanResult without = PublishPlanTests.Build(PublishPlanTests.Manifest(), tree: tree);

        Assert.NotEqual(without.Plan!.Descriptor.ContentHash, with.Plan!.Descriptor.ContentHash);
        Assert.NotEqual(
            without.Plan.DescriptorWithWorkbook(new string('1', 64), new string('2', 64)).ContentHash,
            with.Plan.DescriptorWithWorkbook(new string('1', 64), new string('2', 64)).ContentHash);
    }

    // ###########################################################################################
    // *** THE BOARD'S OWN AND SHARED FILES ARE NOT REWRITTEN WHEN NOTHING CHANGED (code review,
    // 2026-09-25). *** A manifest lists every file the board cites - about 1,200 for the C64
    // 250407 board - so a one-cell typo fix used to re-verify and rewrite every one of them,
    // touching each file's modified time and making the hash caches and the checksum manifest
    // re-hash the lot. A file already published byte-identical at the same path is left alone,
    // exactly as another board's file already was - and stays in the content hash.
    // ###########################################################################################
    [Theory]
    [InlineData(PublishPlanTests.Own + "Images/u8.png")]
    [InlineData("Commodore/Shared files/Component images/6526.jpg")]
    [InlineData("Generic shared files/Component images/7805.jpg")]
    public void A_file_already_published_byte_identical_is_planned_but_never_written(string path)
    {
        string hash = new('5', 64);

        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(path, hash)),
            tree: PublishPlanTests.Tree((path, hash)));

        Assert.True(result.IsPlanned, string.Join(" ", result.Problems.Select(problem => problem.Message)));
        Assert.Empty(result.Plan!.Files);
        Assert.Equal(path, Assert.Single(result.Plan.UnchangedFiles).RelativePath);
    }

    [Fact]
    public void A_file_whose_published_bytes_DIFFER_is_written()
    {
        string path = PublishPlanTests.Own + "Images/u8.png";

        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(path, new string('5', 64))),
            tree: PublishPlanTests.Tree((path, new string('6', 64))));

        Assert.Equal(path, Assert.Single(result.Plan!.Files).RelativePath);
        Assert.Empty(result.Plan.UnchangedFiles);
    }

    // Written or left alone, the file is part of what the system holds, so the content hash - what
    // every client compares to decide whether to sync - must not depend on which it was.
    [Fact]
    public void Leaving_an_unchanged_file_alone_does_not_change_the_content_hash()
    {
        string path = PublishPlanTests.Own + "Images/u8.png";
        string hash = new('5', 64);
        SubmissionManifest manifest = PublishPlanTests.Manifest(PublishPlanTests.File(path, hash));

        PublishPlanResult unchanged = PublishPlanTests.Build(manifest, tree: PublishPlanTests.Tree((path, hash)));
        PublishPlanResult written = PublishPlanTests.Build(manifest, tree: PublishPlanTests.Tree());

        Assert.Empty(unchanged.Plan!.Files);
        Assert.Single(written.Plan!.Files);
        Assert.Equal(written.Plan.Descriptor.ContentHash, unchanged.Plan.Descriptor.ContentHash);
    }

    [Fact]
    public void A_path_differing_from_a_PUBLISHED_one_only_by_case_is_refused()
    {
        // Two files on the Linux server, one on every Windows and macOS client - whichever syncs
        // last wins, so this would overwrite the published file for most users.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "images/u8.png")),
            tree: PublishPlanTests.Tree((PublishPlanTests.Own + "Images/u8.png", new string('4', 64))));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "path.case-collision"));
    }

    [Fact]
    public void A_NEW_SYSTEM_whose_folder_differs_from_a_published_one_only_by_case_is_refused()
    {
        // "commodore/C64/250407" beside the real "Commodore/C64/250407": on a client the two are
        // one folder, so the "new" system would replace the real board's files there.
        SubmissionManifest manifest = PublishPlanTests.Manifest();
        manifest.SystemId = "commodore/C64/250407";
        manifest.Manufacturer = "commodore";

        PublishPlanResult result = PublishPlanTests.Build(
            manifest,
            tree: PublishPlanTests.Tree(("Commodore/C64/250407/Sheet1.png", new string('5', 64))));

        Assert.False(result.IsPlanned);
        Assert.True(PublishPlanTests.HasCode(result, "system.case-collision"));
    }

    [Fact]
    public void An_EXACT_published_path_is_not_a_case_collision()
    {
        // Anti-vacuity: replacing a published file under its own spelling is the ordinary case.
        PublishPlanResult result = PublishPlanTests.Build(
            PublishPlanTests.Manifest(PublishPlanTests.File(PublishPlanTests.Own + "Images/u8.png")),
            tree: PublishPlanTests.Tree((PublishPlanTests.Own + "Images/u8.png", new string('6', 64))));

        Assert.True(result.IsPlanned);
        Assert.Single(result.Plan!.Files);
    }

    [Fact]
    public void The_plan_carries_the_data_root_so_the_writer_can_check_for_links()
    {
        PublishPlanResult result = PublishPlanTests.Build(PublishPlanTests.Manifest());

        Assert.Equal(PublishPlanTests.Root, result.Plan!.DataRoot);
    }
}
