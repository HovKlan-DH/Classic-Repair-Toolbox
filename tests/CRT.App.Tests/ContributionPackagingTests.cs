using Handlers.DataHandling;
using System;
using System.IO;

namespace ClassicRepairToolbox.Tests;

// Tests for ContributionPackaging - the pure half of the component contribution upload.
//
// The zip entry names this class assigns are recorded per row inside ComponentContribution.json
// (the "ZipEntry" field), and the server side of the old contribution upload used them to locate
// each submitted file exactly. That makes the naming scheme a wire contract:
// "ReferencedFiles/<SectionFolder>/<NNN>_<filename>" with one global running number and one shared
// entry per distinct source file. Change the scheme and queued submissions stop resolving.
public class ContributionPackagingTests
{
    // -------------------------------------------------------------- AssignZipEntries

    // The source paths in these tests use forward slashes on purpose: both Windows and Linux
    // parse them as path separators, while a literal @"C:\..." only parses on Windows - and the
    // CI test run happens on ubuntu-latest.
    [Fact]
    public void Entries_are_numbered_globally_across_sections_and_carry_the_section_folder()
    {
        var plan = ContributionPackaging.AssignZipEntries(new[]
        {
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = "/data/pin1.png" },
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = "/data/pin2.png" },
            new ContributionFileReference { SectionFolder = "BoardLocalFiles", ResolvedSourcePath = "/data/schematic.pdf" },
        });

        Assert.Equal(
            new[]
            {
                "ReferencedFiles/ComponentImages/001_pin1.png",
                "ReferencedFiles/ComponentImages/002_pin2.png",
                "ReferencedFiles/BoardLocalFiles/003_schematic.pdf",
            },
            plan.EntryNames);

        Assert.Equal(3, plan.Attachments.Count);
        Assert.Equal("/data/pin1.png", plan.Attachments[0].SourcePath);
        Assert.Equal("ReferencedFiles/ComponentImages/001_pin1.png", plan.Attachments[0].ZipEntryName);
    }

    // Two rows pointing at the same file must share one zip entry instead of packing the file
    // twice - and the second row's ZipEntry still tells the server exactly where the file lives,
    // even when that entry was created under another section's folder.
    [Fact]
    public void Rows_sharing_a_source_file_share_one_zip_entry()
    {
        var plan = ContributionPackaging.AssignZipEntries(new[]
        {
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = "/data/shared.png" },
            new ContributionFileReference { SectionFolder = "ComponentLocalFiles", ResolvedSourcePath = "/DATA/SHARED.PNG" },
        });

        Assert.Single(plan.Attachments);
        Assert.Equal("ReferencedFiles/ComponentImages/001_shared.png", plan.EntryNames[0]);
        Assert.Equal("ReferencedFiles/ComponentImages/001_shared.png", plan.EntryNames[1]);
    }

    // Different files that happen to have the same file name must NOT collapse into one entry.
    // The old upload matched submitted files by bare file name on the server, which made exactly
    // this case ambiguous - the per-row ZipEntry pointer exists to fix it.
    [Fact]
    public void Same_named_files_from_different_folders_get_distinct_entries()
    {
        var plan = ContributionPackaging.AssignZipEntries(new[]
        {
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = "/data/U1/pin1.png" },
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = "/data/U5/pin1.png" },
        });

        Assert.Equal(2, plan.Attachments.Count);
        Assert.Equal("ReferencedFiles/ComponentImages/001_pin1.png", plan.EntryNames[0]);
        Assert.Equal("ReferencedFiles/ComponentImages/002_pin1.png", plan.EntryNames[1]);
    }

    // A row whose file could not be resolved contributes nothing to the zip, gets an empty
    // ZipEntry, and must not consume a number - the server treats empty as "no file attached".
    [Fact]
    public void Unresolved_rows_get_an_empty_entry_and_do_not_consume_a_number()
    {
        var plan = ContributionPackaging.AssignZipEntries(new[]
        {
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = null },
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = "   " },
            new ContributionFileReference { SectionFolder = "ComponentImages", ResolvedSourcePath = "/data/pin1.png" },
        });

        Assert.Equal(new[] { "", "", "ReferencedFiles/ComponentImages/001_pin1.png" }, plan.EntryNames);
        Assert.Single(plan.Attachments);
    }

    [Fact]
    public void No_references_produce_an_empty_plan()
    {
        var plan = ContributionPackaging.AssignZipEntries(Array.Empty<ContributionFileReference>());

        Assert.Empty(plan.EntryNames);
        Assert.Empty(plan.Attachments);
    }

    // -------------------------------------------------------------- ResolveExistingFilePath

    [Fact]
    public void An_existing_absolute_path_resolves_to_itself()
    {
        using var workspace = new TempWorkspace();
        string file = Path.GetFullPath(workspace.WriteFile("external/photo.png", "x"));

        Assert.Equal(file, ContributionPackaging.ResolveExistingFilePath(workspace.Path_("data"), file));
    }

    [Fact]
    public void A_missing_absolute_path_resolves_to_null()
    {
        using var workspace = new TempWorkspace();

        Assert.Null(ContributionPackaging.ResolveExistingFilePath(
            workspace.Path_("data"),
            workspace.Path_("external", "does-not-exist.png")));
    }

    // Relative paths are the form stored in the board Excel files ("Commodore/C64/250407/x.png"),
    // always with forward slashes, and resolve against the data root.
    [Fact]
    public void A_relative_path_with_forward_slashes_resolves_inside_the_data_root()
    {
        using var workspace = new TempWorkspace();
        string expected = Path.GetFullPath(workspace.WriteFile("data/Commodore/C64/pin1.png", "x"));

        string? resolved = ContributionPackaging.ResolveExistingFilePath(
            workspace.Path_("data"),
            "Commodore/C64/pin1.png");

        Assert.Equal(expected, resolved);
    }

    [Fact]
    public void A_relative_path_missing_from_the_data_root_resolves_to_null()
    {
        using var workspace = new TempWorkspace();
        Directory.CreateDirectory(workspace.Path_("data"));

        Assert.Null(ContributionPackaging.ResolveExistingFilePath(workspace.Path_("data"), "Commodore/missing.png"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Blank_input_resolves_to_null(string? pathValue)
    {
        using var workspace = new TempWorkspace();

        Assert.Null(ContributionPackaging.ResolveExistingFilePath(workspace.Path_("data"), pathValue));
    }

    [Fact]
    public void A_relative_path_with_no_data_root_resolves_to_null()
    {
        Assert.Null(ContributionPackaging.ResolveExistingFilePath("", "Commodore/C64/pin1.png"));
    }

    // -------------------------------------------------------------- IsDisplayableImageFile

    // The contribution editor uses this to decide whether a chosen component image can actually
    // be shown. Every one of these formats is decodable by the Avalonia Bitmap the app draws
    // component images with, and all of them except .webp appear in the shipped board data.
    [Theory]
    [InlineData("pin1.png")]
    [InlineData("pin1.jpg")]
    [InlineData("pin1.jpeg")]
    [InlineData("pin1.gif")]
    [InlineData("pin1.bmp")]
    [InlineData("pin1.webp")]
    public void Every_format_the_app_can_draw_is_accepted(string fileName)
    {
        Assert.True(ContributionPackaging.IsDisplayableImageFile(fileName));
    }

    // The case that started this: an .xlsx uploads perfectly happily and then shows as an empty
    // frame. So does every other non-image, and so does a name carrying no extension at all.
    [Theory]
    [InlineData("baselines.xlsx")]
    [InlineData("datasheet.pdf")]
    [InlineData("notes.txt")]
    [InlineData("board.kicad_pcb")]
    [InlineData("README")]
    [InlineData("archive.png.zip")]
    public void A_file_the_app_cannot_draw_is_rejected(string fileName)
    {
        Assert.False(ContributionPackaging.IsDisplayableImageFile(fileName));
    }

    // .svg is in the ExternalTargetLauncher allowlist but deliberately NOT here: that allowlist
    // guards files handed to the OS shell, which has an SVG viewer, whereas a component image is
    // drawn by Avalonia Bitmap, which cannot decode SVG. Do not "fix" this by adding it.
    [Fact]
    public void An_svg_is_rejected_even_though_the_launcher_allowlist_permits_it()
    {
        Assert.False(ContributionPackaging.IsDisplayableImageFile("logo.svg"));
    }

    // Contributed file names are typed by hand, so extension case and stray whitespace are
    // expected rather than exceptional, and neither should reject a perfectly good image.
    [Theory]
    [InlineData("PIN1.PNG")]
    [InlineData("Pin1.JpG")]
    [InlineData("  pin1.png  ")]
    public void Extension_case_and_surrounding_whitespace_do_not_matter(string fileName)
    {
        Assert.True(ContributionPackaging.IsDisplayableImageFile(fileName));
    }

    // Full paths are what the file picker hands back, and relative ones are what a row loaded
    // from board data holds - both must be judged on the file name at the end.
    [Theory]
    [InlineData("/data/Commodore/C64/250407/pin1.png")]
    [InlineData("Commodore/C64/250407/pin1.png")]
    public void The_extension_is_read_from_the_end_of_a_path(string pathValue)
    {
        Assert.True(ContributionPackaging.IsDisplayableImageFile(pathValue));
    }

    // Fail closed: no file chosen is not a displayable image. The caller treats a blank row as
    // "nothing attached" separately, before it ever asks this question.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Blank_input_is_rejected(string? pathValue)
    {
        Assert.False(ContributionPackaging.IsDisplayableImageFile(pathValue));
    }

    // A directory name that happens to look like an image must not sneak through on the strength
    // of a parent folder: only the final segment carries the extension.
    [Fact]
    public void A_folder_named_like_an_image_does_not_make_the_file_inside_it_an_image()
    {
        Assert.False(ContributionPackaging.IsDisplayableImageFile("/data/pin1.png/baselines.xlsx"));
    }

    // -------------------------------------------------------------- ValidateComponentImageFile

    // Case 2 (2026-10-02). "Add new component image" creates a row with every field blank, and it
    // is perfectly normal for it to sit there while the rest of the row is filled in. What must not
    // happen is that row being saved: it would ship a component image entry declaring nothing.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_row_with_no_file_and_no_note_is_refused(string? storedPath)
    {
        Assert.Equal(
            ContributionPackaging.ComponentImageFileProblem.NoFileOrNote,
            ContributionPackaging.ValidateComponentImageFile(storedPath, note: string.Empty));
    }

    // Case 3. Spaces are not a note - the table and the server trim them away the same way, so a
    // row "kept" by a note of blanks would be the empty row of case 2 everywhere else.
    [Theory]
    [InlineData("   ")]
    [InlineData("\t")]
    public void A_row_with_no_file_and_a_note_of_only_spaces_is_refused(string note)
    {
        Assert.Equal(
            ContributionPackaging.ComponentImageFileProblem.NoFileOrNote,
            ContributionPackaging.ValidateComponentImageFile(string.Empty, note));
    }

    // Case 4. A chosen file of the wrong type is a different problem from no file at all, and the
    // editor says so differently - so the two must not collapse into one "row is bad" answer. A
    // note does not excuse it: the row still names a file that draws as an empty frame.
    [Theory]
    [InlineData("baselines.xlsx", "")]
    [InlineData("datasheet.pdf", "")]
    [InlineData("logo.svg", "")]
    [InlineData("pinout.pdf", "Compatible part-number: 1N4148")]
    public void A_file_the_app_cannot_display_is_refused_with_or_without_a_note(string storedPath, string note)
    {
        Assert.Equal(
            ContributionPackaging.ComponentImageFileProblem.NotDisplayable,
            ContributionPackaging.ValidateComponentImageFile(storedPath, note));
    }

    // Case 5. Both shapes a row can hold: an absolute path from the file picker, and the relative
    // path a row loaded from existing board data carries.
    [Theory]
    [InlineData("/pictures/scope/pin1.png")]
    [InlineData("Commodore/C64/250407/pin1.jpg")]
    public void A_displayable_file_and_no_note_is_accepted(string storedPath)
    {
        Assert.Equal(
            ContributionPackaging.ComponentImageFileProblem.None,
            ContributionPackaging.ValidateComponentImageFile(storedPath, note: string.Empty));
    }

    // The validation deliberately does NOT look at the disk. A row can name a file that has not
    // been synced locally yet and still be a submittable row; whether the file resolves is
    // ResolveExistingFilePath's job, and a missing file simply attaches nothing.
    [Fact]
    public void A_path_that_does_not_exist_on_disk_is_still_a_valid_row()
    {
        Assert.Equal(
            ContributionPackaging.ComponentImageFileProblem.None,
            ContributionPackaging.ValidateComponentImageFile("/no/such/folder/pin1.png", note: string.Empty));
    }

    // ###########################################################################################
    // Case 8 (owner report, 2026-10-02: the editor refused to save C128 components whose "Pinout"
    // image row carries only a note, while the Drafts tab's table showed nothing wrong). The editor
    // and the table must refuse the SAME rows for naming no file - CR13 on the C128DCR board (no file,
    // no note) in both, CR4 on the C128 board (a note, no file) in neither. Both sides are asked about
    // the same rows here, so this fails if either one changes its mind alone.
    // ###########################################################################################
    [Theory]
    [InlineData("CR13", "", "", true)]
    [InlineData("CR4", "", "Compatible part-number: 1N4148", false)]
    [InlineData("CR5", "", "   ", true)]
    [InlineData("CR6", "Commodore/C128/310378/cr6.png", "", false)]
    public void The_editor_and_the_table_refuse_the_same_image_rows_for_naming_no_file(
        string label, string file, string note, bool refused)
    {
        bool editorRefuses =
            ContributionPackaging.ValidateComponentImageFile(file, note) == ContributionPackaging.ComponentImageFileProblem.NoFileOrNote;

        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = label, Category = "Diode" });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = label, Name = "Pinout", File = file, Note = note });

        BoardTableSheet images = BoardTableDocument.Create(published: null, board).FindSheet(BoardWorkbookSchema.SheetComponentImages)!;
        BoardTableRow row = Assert.Single(images.Rows);
        BoardTableCell fileCell = row.Cells[images.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFile)];
        bool tableRefuses = fileCell.Problems.Any(problem => problem.Code == "file.unreferenced" && problem.Level == BoardProblemLevel.Error);

        Assert.Equal(refused, tableRefuses);
        Assert.Equal(refused, editorRefuses);
    }

    // -------------------------------------------------------------- ValidateNewComponent

    // "Add new component" opens the editor on a component that exists nowhere yet, so the board
    // label is the only thing identifying it - to the server, and to everybody reading the board
    // data afterwards. A blank one cannot be submitted.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_new_component_without_a_board_label_reports_BoardLabelMissing(string? boardLabel)
    {
        Assert.Equal(
            ContributionPackaging.NewComponentProblem.BoardLabelMissing,
            ContributionPackaging.ValidateNewComponent(boardLabel, "IC", new[] { "U1", "U2" }));
    }

    // Reusing an existing label is refused rather than merged: the server resolves the whole
    // contribution by board label, so it would diff this new component against the existing one
    // and propose deleting every image, file and link the new one did not happen to repeat.
    [Theory]
    [InlineData("U1")]
    [InlineData("u1")]
    [InlineData("  U1  ")]
    public void A_new_component_reusing_an_existing_board_label_reports_BoardLabelAlreadyExists(string boardLabel)
    {
        Assert.Equal(
            ContributionPackaging.NewComponentProblem.BoardLabelAlreadyExists,
            ContributionPackaging.ValidateNewComponent(boardLabel, "IC", new[] { "U1", "C12" }));
    }

    // The board data is read from Excel, where a label can arrive padded - the comparison has to
    // see through that too, or "U1 " on the board lets "U1" through as a new component.
    [Fact]
    public void A_padded_existing_board_label_still_counts_as_taken()
    {
        Assert.Equal(
            ContributionPackaging.NewComponentProblem.BoardLabelAlreadyExists,
            ContributionPackaging.ValidateNewComponent("U1", "IC", new[] { " U1 " }));
    }

    // A component with no category is merged into the board data and is then unreachable: the main
    // window builds its category filter from the categories present and skips blank ones
    // (ComponentListBuilder), so nothing in the UI can ever select it. This is the case that was
    // actually contributed - a component given a name and nothing else.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_new_component_without_a_category_reports_CategoryMissing(string? category)
    {
        Assert.Equal(
            ContributionPackaging.NewComponentProblem.CategoryMissing,
            ContributionPackaging.ValidateNewComponent("U99", category, new[] { "U1", "C12" }));
    }

    // The board label is judged first: a submission that is wrong in both places is told about the
    // label, which is the one that decides what the whole contribution is even about.
    [Fact]
    public void A_missing_board_label_is_reported_before_a_missing_category()
    {
        Assert.Equal(
            ContributionPackaging.NewComponentProblem.BoardLabelMissing,
            ContributionPackaging.ValidateNewComponent("", "", new[] { "U1" }));
    }

    [Fact]
    public void A_new_board_label_with_a_category_reports_None()
    {
        Assert.Equal(
            ContributionPackaging.NewComponentProblem.None,
            ContributionPackaging.ValidateNewComponent("U99", "IC", new[] { "U1", "C12" }));
    }

    // Blank rows in the board data are not labels, so they must not make a blank-looking
    // comparison succeed - and an empty or absent list simply means nothing is taken yet.
    [Fact]
    public void Blank_entries_in_the_existing_labels_are_ignored()
    {
        Assert.Equal(
            ContributionPackaging.NewComponentProblem.None,
            ContributionPackaging.ValidateNewComponent("U99", "IC", new[] { "", "   ", "U1" }));

        Assert.Equal(
            ContributionPackaging.NewComponentProblem.None,
            ContributionPackaging.ValidateNewComponent("U99", "IC", Array.Empty<string>()));

        Assert.Equal(
            ContributionPackaging.NewComponentProblem.None,
            ContributionPackaging.ValidateNewComponent("U99", "IC", null));
    }

    // -------------------------------------------------------------- BuildFeedbackText

    // The server side of the old contribution upload reformatted the summary around the literal
    // marker line "Mandatory change comment:" - the marker is part of the contract.
    [Fact]
    public void The_feedback_text_lists_the_context_and_ends_with_the_mandatory_comment_marker()
    {
        string text = ContributionPackaging.BuildFeedbackText(
            "Commodore 64",
            "250407",
            "U1 | VIC-II | 6569R5",
            "6b29fc40-ca47-1067-b31d-00dd010662da",
            "PAL",
            "Corrected the pin 5 baseline image");

        Assert.Contains("Request type: Component update", text);
        Assert.Contains("Hardware: Commodore 64", text);
        Assert.Contains("Board: 250407", text);
        Assert.Contains("Component: U1 | VIC-II | 6569R5", text);
        Assert.Contains("Component UUID v4: 6b29fc40-ca47-1067-b31d-00dd010662da", text);
        Assert.Contains("Region context: PAL", text);
        Assert.Contains("Mandatory change comment:" + Environment.NewLine + "Corrected the pin 5 baseline image", text);
    }

    // The request type is stated on every submission, not only on a deletion - a line whose
    // absence has to be noticed is a line that gets missed, and a deletion arriving as a mail that
    // reads like an ordinary edit is exactly what this summary exists to prevent.
    [Fact]
    public void A_deletion_says_so_in_the_feedback_text()
    {
        string text = ContributionPackaging.BuildFeedbackText(
            "Commodore 64",
            "250407",
            "U1 | VIC-II | 6569R5",
            "6b29fc40-ca47-1067-b31d-00dd010662da",
            "PAL",
            "This component does not exist on this board revision",
            isDeleteRequest: true);

        Assert.Contains("Request type: DELETE COMPONENT", text);

        // The marker the server formats the notification email around is unchanged.
        Assert.Contains(
            "Mandatory change comment:" + Environment.NewLine + "This component does not exist on this board revision",
            text);
    }

    // -------------------------------------------------------------- BuildDeleteComponentSummary

    // What the contributor is agreeing to lose. The sections are collapsed when the button is
    // pressed, so this sentence is the only place the cost of the deletion is visible.
    [Fact]
    public void The_delete_summary_names_the_component_and_lists_everything_that_goes_with_it()
    {
        string text = ContributionPackaging.BuildDeleteComponentSummary("U1", 4, 2, 3, 1);

        Assert.Equal(
            "The component [U1] will be removed from the board data, together with its "
            + "4 component images, 2 schematic highlights, 3 local files and 1 link.",
            text);
    }

    // Each count carries its own singular, so the sentence never reads "1 component images".
    [Fact]
    public void The_delete_summary_uses_the_singular_for_a_single_row()
    {
        string text = ContributionPackaging.BuildDeleteComponentSummary("C7", 1, 1, 1, 1);

        Assert.Equal(
            "The component [C7] will be removed from the board data, together with its "
            + "1 component image, 1 schematic highlight, 1 local file and 1 link.",
            text);
    }

    // A section with nothing in it is left out rather than reported as "0 images": the list says
    // what will be lost, and a zero is not a loss. With one item left there is no "and" either.
    [Fact]
    public void The_delete_summary_leaves_out_the_sections_that_are_empty()
    {
        Assert.Equal(
            "The component [R3] will be removed from the board data, together with its 2 links.",
            ContributionPackaging.BuildDeleteComponentSummary("R3", 0, 0, 0, 2));

        Assert.Equal(
            "The component [R3] will be removed from the board data, together with its "
            + "5 component images and 1 link.",
            ContributionPackaging.BuildDeleteComponentSummary("R3", 5, 0, 0, 1));
    }

    // A component carrying nothing else still has to produce a sentence, and one that says the
    // deletion is small rather than trailing off after "together with its".
    [Fact]
    public void A_component_carrying_nothing_else_says_so()
    {
        Assert.Equal(
            "The component [U9] will be removed from the board data. It carries no images, highlights, files or links.",
            ContributionPackaging.BuildDeleteComponentSummary("U9", 0, 0, 0, 0));
    }

    // A blank label would otherwise produce "The component [] will be removed".
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_missing_board_label_falls_back_to_naming_no_label_at_all(string? boardLabel)
    {
        string text = ContributionPackaging.BuildDeleteComponentSummary(boardLabel, 1, 0, 0, 0);

        Assert.StartsWith("This component will be removed from the board data,", text);
        Assert.DoesNotContain("[", text);
    }

    // Counts arrive from collection sizes, so a negative is not expected - but treating one as
    // "nothing here" is the only reading that cannot produce "-1 links".
    [Fact]
    public void A_negative_count_is_treated_as_nothing_rather_than_printed()
    {
        Assert.Equal(
            "The component [U1] will be removed from the board data. It carries no images, highlights, files or links.",
            ContributionPackaging.BuildDeleteComponentSummary("U1", -3, 0, 0, 0));
    }

    // -------------------------------------------------------------- TryGetDataRootRelativeFolder
    //
    // THE REGRESSION these exist for: the Contribute tab used to decide "is this picked file inside
    // the data root" with a bare dir.StartsWith(dataRoot) - no trailing separator, no normalization.
    // A SIBLING folder whose name merely begins with the root's name therefore passed, and
    // GetRelativePath turned it into a TRAVERSING relative path ("../Data-backup/Commodore") that was
    // written into the row and shipped in the uploaded payload, where the server merge concatenates
    // it onto its own data root to build a copy() target.
    //
    // Note the three-valued contract, which is what the caller branches on:
    //   a relative path -> inside, in that sub-folder
    //   ""              -> inside, directly in the root (no sub-folder to name)
    //   null            -> NOT inside; the caller keeps the user's chosen drop-down folder

    private static string Root => OperatingSystem.IsWindows()
        ? @"C:\Users\test\AppData\Roaming\Classic-Repair-Toolbox\Data"
        : "/home/test/.config/Classic-Repair-Toolbox/Data";

    private static string Sibling(string name) =>
        Path.Combine(Path.GetDirectoryName(Root)!, name);

    [Fact]
    public void A_folder_inside_the_data_root_yields_its_relative_path()
    {
        Assert.Equal(
            "Commodore/C64",
            ContributionPackaging.TryGetDataRootRelativeFolder(Root, Path.Combine(Root, "Commodore", "C64")));
    }

    // Separators are normalised to forward slashes, because that is what the payload and the
    // server's own path handling use.
    [Fact]
    public void The_relative_path_uses_forward_slashes()
    {
        string? result = ContributionPackaging.TryGetDataRootRelativeFolder(
            Root, Path.Combine(Root, "Commodore", "C64", "250407"));

        Assert.Equal("Commodore/C64/250407", result);
        Assert.DoesNotContain('\\', result!);
    }

    [Fact]
    public void The_data_root_itself_is_inside_it_with_no_sub_folder()
    {
        Assert.Equal(string.Empty, ContributionPackaging.TryGetDataRootRelativeFolder(Root, Root));
    }

    // ###########################################################################################
    // The reported bug, asserted by NAME rather than by mere difference: a sibling folder sharing
    // the root's leading characters is NOT inside it. Against the StartsWith version each of these
    // returns a "../" path instead of null.
    // ###########################################################################################
    [Theory]
    [InlineData("Data-backup")]
    [InlineData("DataOld")]
    [InlineData("Data2")]
    public void A_sibling_folder_sharing_the_roots_name_prefix_is_not_inside(string siblingName)
    {
        Assert.Null(ContributionPackaging.TryGetDataRootRelativeFolder(Root, Sibling(siblingName)));
    }

    // The same thing one level deeper - the exact shape that produced "../Data-backup/Commodore".
    [Fact]
    public void A_sub_folder_of_a_prefix_sharing_sibling_is_not_inside_either()
    {
        string? result = ContributionPackaging.TryGetDataRootRelativeFolder(
            Root, Path.Combine(Sibling("Data-backup"), "Commodore"));

        Assert.Null(result);
    }

    // Whatever this returns, it must NEVER be a traversing path - that is the property that
    // actually protects the payload, independent of which folder was picked.
    [Theory]
    [InlineData("Data-backup")]
    [InlineData("DataOld")]
    public void A_result_is_never_a_traversing_path(string siblingName)
    {
        string? result = ContributionPackaging.TryGetDataRootRelativeFolder(
            Root, Path.Combine(Sibling(siblingName), "Commodore"));

        Assert.True(result == null || !result.Contains(".."), $"got [{result}]");
    }

    [Fact]
    public void A_folder_plainly_outside_the_root_is_not_inside()
    {
        string outside = OperatingSystem.IsWindows() ? @"C:\Users\test\Desktop" : "/home/test/Desktop";

        Assert.Null(ContributionPackaging.TryGetDataRootRelativeFolder(Root, outside));
    }

    // An unnormalised root (trailing separator, or a "." segment in the middle) must still compare
    // correctly - the caller passes whatever DataManager resolved, not a canonical string.
    [Fact]
    public void An_unnormalised_root_still_matches_a_folder_inside_it()
    {
        Assert.Equal(
            "Commodore",
            ContributionPackaging.TryGetDataRootRelativeFolder(
                Root + Path.DirectorySeparatorChar, Path.Combine(Root, "Commodore")));
    }

    [Theory]
    [InlineData(null, "somewhere")]
    [InlineData("", "somewhere")]
    [InlineData("   ", "somewhere")]
    public void A_blank_root_is_not_something_to_be_inside_of(string? root, string directory)
    {
        Assert.Null(ContributionPackaging.TryGetDataRootRelativeFolder(root, directory));
    }

    // Path.GetDirectoryName returns null for a bare file name, which is what this receives then.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void A_blank_directory_is_not_inside(string? directory)
    {
        Assert.Null(ContributionPackaging.TryGetDataRootRelativeFolder(Root, directory));
    }
}
