using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The one set of rules for a board's rows (owner request, 2026-10-02: the checks shown in the
// Drafts tab's table "as either Error or Warning ... All error should be fixed before submission
// can be done") - BoardDataChecks.
//
// Three things must hold, and each has its tests here:
//
//   - WHERE: every problem names the sheet, the entry and the column it is in, so the table can
//     mark the right cell.
//   - THE SERVER IS UNCHANGED: the Submission scope is exactly what SubmissionValidator always
//     found - which SubmissionValidatorTests, untouched by the move, still hold.
//   - AN ERROR IN THE TABLE IS A REFUSAL AT SUBMIT: every error the Everything scope finds is one
//     the server's own checks find for the same board. The test that would fail if the table ever
//     called something an error that the server accepts - or the server started refusing something
//     the table calls a warning - is Every_error_the_table_finds_is_one_the_server_refuses.
// ###########################################################################################
public sealed class BoardDataChecksTests
{
    private const string Board = "Commodore/C64/250407";

    private static BoardCheckRows Rows(
        IReadOnlyList<BoardSchematicEntry>? schematics = null,
        IReadOnlyList<ComponentEntry>? components = null,
        IReadOnlyList<ComponentImageEntry>? images = null,
        IReadOnlyList<ComponentHighlightEntry>? highlights = null,
        IReadOnlyList<ComponentLocalFileEntry>? localFiles = null,
        IReadOnlyList<BoardLocalFileEntry>? boardFiles = null,
        IReadOnlyList<ComponentLinkEntry>? links = null,
        IReadOnlyList<BoardLinkEntry>? boardLinks = null) =>
        new(schematics ?? [MainSchematic], components ?? [], images ?? [], highlights ?? [], localFiles ?? [], boardFiles ?? [], links ?? [], boardLinks ?? []);

    // Every board has a "Main" schematic unless a test gives its own, so the highlights the tests
    // draw on it are not themselves an error.
    private static readonly BoardSchematicEntry MainSchematic = new() { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" };

    // With a category: a component without one is a warning of its own (component.no_category), which
    // only the tests about it want.
    private static ComponentEntry Component(string label, string region = "", string partNumber = "", string technicalName = "") =>
        new() { BoardLabel = label, Category = "IC", Region = region, PartNumber = partNumber, TechnicalNameOrValue = technicalName };

    private static ComponentHighlightEntry Highlight(string schematic, string label) =>
        new() { SchematicName = schematic, BoardLabel = label, X = "1", Y = "1", Width = "10", Height = "10" };

    private static ComponentImageEntry Image(string label, string file = $"{Board}/U1.png", string timeDiv = "", string voltsDiv = "", string triggerLevel = "") =>
        new() { BoardLabel = label, File = file, TimeDiv = timeDiv, VoltsDiv = voltsDiv, TriggerLevelVolts = triggerLevel };

    private static IReadOnlyList<BoardDataProblem> Everything(BoardCheckRows rows, params string[] files) =>
        BoardDataChecks.Check(rows, new SuppliedFileLookup([MainSchematic.SchematicImageFile, .. files]), BoardCheckScope.Everything);

    // ------------------------------------------------------------------ Where a problem is

    [Fact]
    public void A_problem_names_its_sheet_entry_and_column()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            schematics:
            [
                new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" },
                new BoardSchematicEntry { SchematicName = string.Empty, SchematicImageFile = $"{Board}/other.png" }
            ]), $"{Board}/main.png", $"{Board}/other.png");

        BoardDataProblem unnamed = Assert.Single(problems, problem => problem.Code == "schematic.unnamed");

        Assert.Equal(BoardProblemLevel.Error, unnamed.Level);
        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, unnamed.Sheet);
        Assert.Equal(1, unnamed.Index);
        Assert.Equal(BoardWorkbookSchema.ColSchematicName, unnamed.Column);
        Assert.True(unnamed.IsInSheet);
    }

    // EVERY row of a duplicate is marked in the table (owner decision, 2026-10-03: "it should show
    // all rows, and not only last, as the maintainer/contributor then has some context") - the
    // first one saying it is the first, so the two read as one problem.
    [Fact]
    public void A_duplicate_component_is_marked_on_every_row_of_it()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            components: [Component("U1"), Component("U2"), Component("U1")],
            highlights: [Highlight("Main", "U1"), Highlight("Main", "U2")]));

        List<BoardDataProblem> duplicates = problems.Where(problem => problem.Code == "component.duplicate_label").OrderBy(problem => problem.Index).ToList();

        Assert.Equal([0, 2], duplicates.Select(problem => problem.Index));
        Assert.All(duplicates, problem =>
        {
            Assert.Equal(BoardProblemLevel.Error, problem.Level);
            Assert.Equal(BoardWorkbookSchema.ColBoardLabel, problem.Column);
        });
        Assert.Contains("this is the first of them", duplicates[0].Message, StringComparison.Ordinal);
        Assert.Contains("(also [U1])", duplicates[1].Message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_duplicate_schematic_is_marked_on_every_row_of_it()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            schematics:
            [
                new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" },
                new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = $"{Board}/bottom.png" },
                new BoardSchematicEntry { SchematicName = "MAIN", SchematicImageFile = $"{Board}/main2.png" }
            ]), $"{Board}/bottom.png", $"{Board}/main2.png");

        Assert.Equal(
            [0, 2],
            problems.Where(problem => problem.Code == "schematic.duplicate").Select(problem => problem.Index).Order());
    }

    // The SERVER refuses exactly as before: once, on the later row. Only the table adds the first.
    [Fact]
    public void The_server_still_refuses_a_duplicate_on_the_later_row_alone()
    {
        BoardCheckRows rows = Rows(
            schematics:
            [
                new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" },
                new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main2.png" }
            ],
            components: [Component("U1"), Component("U1", region: "PAL"), Component("U1")]);

        IReadOnlyList<BoardDataProblem> problems = BoardDataChecks.Check(rows, files: null, BoardCheckScope.Submission);

        BoardDataProblem component = Assert.Single(problems, problem => problem.Code == "component.duplicate_label");
        Assert.Equal(2, component.Index);

        BoardDataProblem schematic = Assert.Single(problems, problem => problem.Code == "schematic.duplicate");
        Assert.Equal(1, schematic.Index);
    }

    // A highlight lives in the JSON beside the workbook - no sheet, no row.
    [Fact]
    public void A_highlight_problem_is_in_no_sheet()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            components: [Component("U1")],
            highlights: [Highlight("No such schematic", "U1")]));

        BoardDataProblem unknown = Assert.Single(problems, problem => problem.Code == "highlight.unknown_schematic");

        Assert.Null(unknown.Sheet);
        Assert.Equal(-1, unknown.Index);
        Assert.False(unknown.IsInSheet);
    }

    // ------------------------------------------------------------------ Files

    [Fact]
    public void A_file_that_is_not_there_is_an_error_on_its_file_cell()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            components: [Component("U1")],
            images: [Image("U1", $"{Board}/gone.png")],
            highlights: [Highlight("Main", "U1")]));

        BoardDataProblem missing = Assert.Single(problems, problem => problem.Code == "file.missing");

        Assert.Equal(BoardWorkbookSchema.SheetComponentImages, missing.Sheet);
        Assert.Equal(BoardWorkbookSchema.ColFile, missing.Column);
        Assert.Contains("which is not in the submission", missing.Message, StringComparison.Ordinal);
    }

    // With no lookup (the Maintainer tab's table) files are not looked for at all - but a row
    // naming no file is still one.
    [Fact]
    public void Without_a_lookup_files_are_not_looked_for()
    {
        IReadOnlyList<BoardDataProblem> problems = BoardDataChecks.Check(
            Rows(
                components: [Component("U1")],
                images: [Image("U1", $"{Board}/gone.png"), Image("U1", string.Empty)],
                highlights: [Highlight("Main", "U1")]),
            files: null,
            BoardCheckScope.Everything);

        Assert.DoesNotContain(problems, problem => problem.Code == "file.missing");
        Assert.Contains(problems, problem => problem.Code == "file.unreferenced" && problem.Index == 1);
    }

    // A component image row with a note and no file IS a note - the C128's "Pinout" rows that give
    // only a compatible part number. Spaces are no note, and a row naming a file is never one,
    // whatever its note says. The component editor asks this same question (2026-10-02).
    [Theory]
    [InlineData("", "Compatible part-number: 1N4148", true)]
    [InlineData(null, "Compatible part-number: 1N4148", true)]
    [InlineData("", "", false)]
    [InlineData("", "   ", false)]
    [InlineData(null, null, false)]
    [InlineData($"{Board}/U1.png", "Compatible part-number: 1N4148", false)]
    [InlineData($"{Board}/U1.png", "", false)]
    public void An_image_row_is_a_note_only_with_a_note_and_no_file(string? file, string? note, bool noteOnly)
    {
        Assert.Equal(noteOnly, BoardDataChecks.IsNoteOnlyImage(file, note));
    }

    // The server's own file-name rules, called - a type the server will not take, and a path with
    // a backslash - are errors in the table, on the file cell, before the file is looked for.
    [Theory]
    [InlineData($"{Board}/tool.exe", "file.type_not_allowed")]
    [InlineData($"{Board}/.htaccess.png", "file.hidden_name")]
    [InlineData($"{Board}/noextension", "file.no_type")]
    [InlineData("Commodore\\C64\\250407\\U1.png", "path.rejected")] // windows-path-literal: a backslash path is the defect under test, refused on every OS
    public void A_file_name_the_server_refuses_is_an_error(string file, string code)
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            components: [Component("U1")],
            localFiles: [new ComponentLocalFileEntry { BoardLabel = "U1", Name = "Data", File = file }],
            highlights: [Highlight("Main", "U1")]),
            file);

        BoardDataProblem problem = Assert.Single(problems, candidate => candidate.Level == BoardProblemLevel.Error);

        Assert.Equal(code, problem.Code);
        Assert.Equal(BoardWorkbookSchema.ColFile, problem.Column);
    }

    // Only the wider scope - the server's own SubmissionValidator never applied these (another
    // class of its does), so its scope must not start to.
    [Fact]
    public void The_submission_scope_does_not_apply_the_file_name_rules()
    {
        IReadOnlyList<BoardDataProblem> problems = BoardDataChecks.Check(
            Rows(localFiles: [new ComponentLocalFileEntry { BoardLabel = "U1", File = $"{Board}/tool.exe" }]),
            new SuppliedFileLookup([MainSchematic.SchematicImageFile, $"{Board}/tool.exe"]),
            BoardCheckScope.Submission);

        Assert.Empty(problems);
    }

    // ------------------------------------------------------------------ The warnings DataValidator logged

    // A comma-decimal opacity ("0,3") is what a CRT on a Danish or German machine wrote into drafts
    // before WorkbookCellText read numbers invariantly; CRT draws such a value with its default
    // opacity on EVERY machine, so the table says so (code review, 2026-10-13).
    [Theory]
    [InlineData("", "", null)]
    [InlineData("0.3", "0.65", null)]
    [InlineData("40%", " 25 % ", null)]
    [InlineData("1", "0", null)]
    [InlineData("0,3", "", "schematic.opacity")]
    [InlineData("", "0,65", "schematic.opacity")]
    [InlineData("half", "", "schematic.opacity")]
    public void An_opacity_CRT_cannot_read_is_a_warning(string schematicOpacity, string thumbnailOpacity, string? code)
    {
        var schematic = new BoardSchematicEntry
        {
            SchematicName = MainSchematic.SchematicName,
            SchematicImageFile = MainSchematic.SchematicImageFile,
            SchematicHighlightOpacity = schematicOpacity,
            ThumbnailHighlightOpacity = thumbnailOpacity,
        };

        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(schematics: [schematic]));

        if (code is null)
        {
            Assert.Empty(problems);
            return;
        }

        BoardDataProblem problem = Assert.Single(problems);
        Assert.Equal(code, problem.Code);
        Assert.Equal(BoardProblemLevel.Warning, problem.Level);
        Assert.Equal(
            schematicOpacity.Length > 0 ? BoardWorkbookSchema.ColSchematicHighlightOpacity : BoardWorkbookSchema.ColThumbnailHighlightOpacity,
            problem.Column);
    }

    [Theory]
    [InlineData("500uS", "", "", null)]
    [InlineData("5 ms", "", "", "image.time_div")]
    [InlineData("", "1V", "", null)]
    [InlineData("", "1 Volt", "", "image.volts_div")]
    [InlineData("", "", "-1.4V", null)]
    [InlineData("", "", "1,4V", "image.trigger_level")]
    public void An_oscilloscope_setting_CRT_does_not_know_is_a_warning(string timeDiv, string voltsDiv, string triggerLevel, string? code)
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(
            Rows(
                components: [Component("U1")],
                images: [Image("U1", timeDiv: timeDiv, voltsDiv: voltsDiv, triggerLevel: triggerLevel)],
                highlights: [Highlight("Main", "U1")]),
            $"{Board}/U1.png");

        if (code is null)
        {
            Assert.Empty(problems);
            return;
        }

        BoardDataProblem problem = Assert.Single(problems);
        Assert.Equal(code, problem.Code);
        Assert.Equal(BoardProblemLevel.Warning, problem.Level);
    }

    [Fact]
    public void Files_and_links_for_a_component_that_is_not_there_are_warnings()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(
            Rows(
                components: [Component("U1")],
                highlights: [Highlight("Main", "U1")],
                localFiles: [new ComponentLocalFileEntry { BoardLabel = "U9", Name = "Data", File = $"{Board}/u9.pdf" }],
                links: [new ComponentLinkEntry { BoardLabel = " u1 ", Name = "Kept", Url = "https://example.com" }, new ComponentLinkEntry { BoardLabel = "U9", Name = "Gone", Url = "https://example.com" }]),
            $"{Board}/u9.pdf");

        Assert.Equal(["file.orphan", "link.orphan"], problems.Select(problem => problem.Code));
        Assert.All(problems, problem => Assert.Equal(BoardProblemLevel.Warning, problem.Level));
        Assert.Equal(1, problems[1].Index);
    }

    // ###########################################################################################
    // A file's or a link's Region says PAL, NTSC or nothing (code review, 2026-11-02): any other word
    // shows the row in NEITHER region, so it vanishes from the pane, the popup and the "Resources" tab with
    // nothing said - and the table takes anything typed. A warning, as every new local rule is.
    // ###########################################################################################
    [Fact]
    public void A_file_or_link_of_a_region_that_does_not_exist_is_a_warning_on_its_cell()
    {
        BoardCheckRows rows = Rows(
            components: [Component("U1")],
            highlights: [Highlight("Main", "U1")],
            localFiles:
            [
                new ComponentLocalFileEntry { BoardLabel = "U1", Name = "Fine", File = $"{Board}/a.pdf", Region = "pal" },
                new ComponentLocalFileEntry { BoardLabel = "U1", Name = "Odd", File = $"{Board}/b.pdf", Region = "Europe" },
            ],
            links: [new ComponentLinkEntry { BoardLabel = "U1", Name = "Odd", Url = "https://example.com", Region = "EU" }],
            boardFiles: [new BoardLocalFileEntry { Category = "Resources", Name = "Odd", File = $"{Board}/c.pdf", Region = "Japan" }],
            boardLinks:
            [
                new BoardLinkEntry { Category = "Resources", Name = "Every region", Url = "https://example.com" },
                new BoardLinkEntry { Category = "Resources", Name = "Odd", Url = "https://example.com", Region = "SECAM" },
            ]);

        List<BoardDataProblem> problems = Everything(rows, $"{Board}/a.pdf", $"{Board}/b.pdf", $"{Board}/c.pdf")
            .Where(problem => problem.Code == "row.region")
            .ToList();

        Assert.Equal(
            [
                (BoardWorkbookSchema.SheetComponentLocalFiles, 1),
                (BoardWorkbookSchema.SheetComponentLinks, 0),
                (BoardWorkbookSchema.SheetBoardLocalFiles, 0),
                (BoardWorkbookSchema.SheetBoardLinks, 1),
            ],
            problems.Select(problem => (problem.Sheet!, problem.Index)));
        Assert.All(problems, problem =>
        {
            Assert.Equal(BoardProblemLevel.Warning, problem.Level);
            Assert.Equal(BoardWorkbookSchema.ColRegion, problem.Column);
        });
        Assert.Contains("[Europe]", problems[0].Message, StringComparison.Ordinal);
    }

    // A component nothing marks on any schematic; a highlight for a component that is not there.
    [Fact]
    public void Components_and_highlights_that_do_not_meet_are_warnings()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            components: [Component("U1"), Component("U2")],
            highlights: [Highlight("Main", "u1"), Highlight("Main", "U7")]));

        BoardDataProblem unmarked = Assert.Single(problems, problem => problem.Code == "component.no_highlight");
        Assert.Equal(1, unmarked.Index);
        Assert.Equal("U2", unmarked.Subject);

        BoardDataProblem stray = Assert.Single(problems, problem => problem.Code == "highlight.orphan");
        Assert.False(stray.IsInSheet);
    }

    // One part number on two different chips: every row involved is marked, naming the others.
    [Fact]
    public void One_part_number_on_two_different_chips_marks_every_row_involved()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(
            components:
            [
                Component("U17", partNumber: "906114-01", technicalName: "PLA"),
                Component("U16", partNumber: "906114-01", technicalName: "4066"),
                Component("U15", partNumber: "906114-01", technicalName: "PLA"),
                Component("U1", partNumber: "901225-01", technicalName: "6510")
            ],
            highlights: [Highlight("Main", "U17"), Highlight("Main", "U16"), Highlight("Main", "U15"), Highlight("Main", "U1")]));

        List<BoardDataProblem> mixed = problems.Where(problem => problem.Code == "component.part_number_mixed").ToList();

        Assert.Equal([0, 1, 2], mixed.Select(problem => problem.Index));
        Assert.All(mixed, problem => Assert.Equal(BoardWorkbookSchema.ColPartNumber, problem.Column));
        Assert.Contains("U16 (4066)", mixed[0].Message, StringComparison.Ordinal);
        Assert.DoesNotContain("U15", mixed[0].Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ An error is a refusal

    // ###########################################################################################
    // *** EVERY ERROR THE TABLE FINDS IS ONE THE SERVER REFUSES. *** A board with one defect of each
    // kind, sent as a submission would send it - its rows, and the files they name as the files it
    // carries (less the missing one). The server's three checks are run on it as the server runs
    // them; every error code the table's Everything scope reports must be among theirs. Fails if a
    // rule is added to the table as an error the server would accept.
    // ###########################################################################################
    [Fact]
    public void Every_error_the_table_finds_is_one_the_server_refuses()
    {
        var rows = new SubmissionRows
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" },
                new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main2.png" },
                new BoardSchematicEntry { SchematicName = string.Empty, SchematicImageFile = $"{Board}/x.png" }
            ],
            Components = [Component("U1"), Component("U1"), Component(string.Empty)],
            ComponentImages = [Image("U1", $"{Board}/gone.png"), Image("U1", $"{Board}/Main.PNG")],
            ComponentHighlights = [Highlight("Nowhere", "U1"), new ComponentHighlightEntry { SchematicName = "Main", BoardLabel = "U1", X = "1,5", Y = "0", Width = "1", Height = "1" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U1", Name = "Tool", File = $"{Board}/tool.exe" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Bad", Url = "ftp://example.com" }],

            // A field a row of the board's own lists needs, empty (BoardRowRequired, 2026-11-01).
            BoardLinks = [new BoardLinkEntry { Category = "Forums", Url = "https://www.lemon64.com" }],
            Credits = [new CreditEntry { NameOrHandle = "Anna" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { KiCadNetName = "/RESET" }]
        };

        // Everything the rows name, less the one that is missing - and with "Main.PNG" carried as
        // "main.png", the spelling the second image's row gets wrong.
        string[] carried = [$"{Board}/main.png", $"{Board}/main2.png", $"{Board}/x.png", $"{Board}/tool.exe"];

        var manifest = new SubmissionManifest
        {
            BoardId = Board,
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Summary = "Every defect at once.",
            Rows = rows,
            Files = carried.Select(path => new SubmissionFile { Path = path, Sha256 = new string('a', 64), SizeBytes = 1 }).ToList()
        };

        HashSet<string> serverErrors = SubmissionValidator.Validate(manifest, carried)
            .Concat(SubmissionFileRules.ValidateManifestFiles(manifest, board: null))
            .Where(finding => finding.Severity == ValidationSeverity.Error)
            .Select(finding => finding.Code)
            .ToHashSet(StringComparer.Ordinal);

        List<string> tableErrors = BoardDataChecks.Check(BoardCheckRows.From(rows), new SuppliedFileLookup(carried), BoardCheckScope.Everything)
            .Where(problem => problem.Level == BoardProblemLevel.Error)
            .Select(problem => problem.Code)
            .Distinct()
            .ToList();

        Assert.NotEmpty(tableErrors);
        Assert.All(tableErrors, code => Assert.Contains(code, serverErrors));
    }

    // ###########################################################################################
    // *** AN EMPTY FILE MARKED ON THE TABLE'S CELL IS ONE THE SERVER REFUSES (owner decision, 2026-11-02:
    // "Refuse empty files"). *** Each side asked for itself - the table of the file on this disk, the
    // server of the manifest a submit would build - and the code and the words compared: moved on one
    // side only, this fails. The server's lookup knows names, so the test above never meets this error.
    // ###########################################################################################
    [Fact]
    public void An_empty_file_the_table_marks_is_one_the_server_refuses_in_the_same_words()
    {
        using var workspace = new TempWorkspace();
        string draftFolder = Path.GetDirectoryName(workspace.WriteFile(Path.Combine("Drafts", "notes.txt"), string.Empty))!;

        var rows = new SubmissionRows
        {
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "Docs", Name = "Notes", File = "notes.txt" }],
        };

        BoardDataProblem marked = Assert.Single(
            BoardDataChecks.Check(BoardCheckRows.From(rows), new DiskFileLookup(workspace.Path_("Data"), draftFolder), BoardCheckScope.Everything),
            problem => problem.Level == BoardProblemLevel.Error);

        Assert.Equal(BoardWorkbookSchema.SheetBoardLocalFiles, marked.Sheet);
        Assert.Equal(0, marked.Index);
        Assert.Equal(BoardWorkbookSchema.ColFile, marked.Column);

        var manifest = new SubmissionManifest
        {
            BoardId = Board,
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Rows = rows,
            Files = [new SubmissionFile { Path = "notes.txt", Sha256 = new string('a', 64), SizeBytes = 0 }],
        };

        ValidationFinding refused = Assert.Single(SubmissionFileRules.ValidateManifestFiles(manifest, board: null));

        Assert.Equal(refused.Code, marked.Code);
        Assert.Equal(refused.Message, marked.Message);
    }

    // And the wider scope adds no ERROR the server's own row rules do not have - beyond the file
    // names, which the server checks in another class (the test above holds those).
    [Fact]
    public void The_wider_scope_adds_only_warnings_to_the_servers_row_rules()
    {
        BoardCheckRows rows = Rows(
            schematics: [new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" }],
            components: [Component("U1", partNumber: "1", technicalName: "A"), Component("U2", partNumber: "1", technicalName: "B")],
            images: [Image("U9", timeDiv: "fast")],
            highlights: [Highlight("Main", "U7")],
            localFiles: [new ComponentLocalFileEntry { BoardLabel = "U8", File = $"{Board}/u8.pdf" }],
            links: [new ComponentLinkEntry { BoardLabel = "U8", Url = "ftp://x" }]);

        var lookup = new SuppliedFileLookup([$"{Board}/main.png", $"{Board}/U1.png", $"{Board}/u8.pdf"]);

        IEnumerable<string> Errors(BoardCheckScope scope) =>
            BoardDataChecks.Check(rows, lookup, scope).Where(problem => problem.Level == BoardProblemLevel.Error).Select(problem => problem.Code);

        Assert.Equal(Errors(BoardCheckScope.Submission), Errors(BoardCheckScope.Everything));
        Assert.True(BoardDataChecks.Check(rows, lookup, BoardCheckScope.Everything).Count > BoardDataChecks.Check(rows, lookup, BoardCheckScope.Submission).Count);
    }

    [Fact]
    public void The_worst_of_no_problems_is_none()
    {
        Assert.Equal(BoardProblemLevel.None, BoardDataChecks.Worst([]));
        Assert.Equal(
            BoardProblemLevel.Error,
            BoardDataChecks.Worst(
            [
                new BoardDataProblem(BoardProblemLevel.Warning, "w", "", "", null, -1, null),
                new BoardDataProblem(BoardProblemLevel.Error, "e", "", "", null, -1, null)
            ]));
    }

    // ------------------------------------------------------------------ Duplicate rows (2026-10-03)

    // ###########################################################################################
    // *** "FLAGGED" IS A WARNING NOW (owner request, 2026-10-03: "Should flagged now be treated
    // as warnings?"; cases agreed with the project owner). *** Two rows with the same identity on
    // a sheet are both saved and published, but only the first is compared with the published data
    // - so a change in the others is counted nowhere and a maintainer sees nothing of it. The
    // table's violet "Flagged" rows said so; the checks say it now, as a warning on every row of
    // the set. On Components and Board schematics such a row is already the server's ERROR, so
    // those two sheets are left to it.
    // ###########################################################################################
    private const string LaterDuplicate =
        "Another row on this sheet has the same Board label, Region, Pin and Name. Only the first is compared " +
        "with the published data, so a change in this one is not shown to a maintainer. Make them differ, or delete one.";

    private const string FirstOfTwo =
        "Another row on this sheet has the same Board label, Region, Pin and Name. Only this one, the first, is " +
        "compared with the published data, so a change in the other is not shown to a maintainer. Make them differ, or delete one.";

    private static ComponentImageEntry R307Pinout(string file) =>
        new() { BoardLabel = "R307", Name = "Pinout (secondary)", File = file };

    private static List<BoardDataProblem> Duplicates(IEnumerable<BoardDataProblem> problems) =>
        problems.Where(problem => problem.Code == "row.duplicate").ToList();

    // Case 12, as the checks find it: the owner's two R307 rows.
    [Fact]
    public void Two_component_image_rows_with_the_same_identity_are_both_warned_about()
    {
        string first = "Generic shared files/Component images/resistor_3k3_5.png";
        string second = "Generic shared files/Component images/resistor.png";

        List<BoardDataProblem> duplicates = Duplicates(Everything(
            Rows(components: [Component("R307")], images: [R307Pinout(first), R307Pinout(second)]),
            first, second));

        Assert.Equal(2, duplicates.Count);
        Assert.All(duplicates, problem =>
        {
            Assert.Equal(BoardProblemLevel.Warning, problem.Level);
            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, problem.Sheet);
            Assert.Equal(BoardWorkbookSchema.ColBoardLabel, problem.Column);
        });

        Assert.Equal([0, 1], duplicates.Select(problem => problem.Index));
        Assert.Equal(FirstOfTwo, duplicates[0].Message);
        Assert.Equal(LaterDuplicate, duplicates[1].Message);
    }

    // Case 13: on Components a duplicate is the server's error, as before, and nothing more.
    [Fact]
    public void Two_components_with_the_same_label_are_errors_and_never_also_warnings()
    {
        IReadOnlyList<BoardDataProblem> problems = Everything(Rows(components: [Component("U1"), Component("U1")]));

        Assert.Contains(problems, problem =>
            problem.Code == "component.duplicate_label" && problem.Level == BoardProblemLevel.Error && problem.Index == 1);

        Assert.Empty(Duplicates(problems));
    }

    // Case 19: the launch log runs these same checks over each published board, so it reports
    // duplicates too - on every sheet the checks read, Credits and Important signals included.
    // The server's checks do not change.
    [Fact]
    public void A_duplicate_row_is_a_warning_for_the_launch_log_and_nothing_the_server_reports()
    {
        var board = new BoardData
        {
            Schematics = [MainSchematic],
            Components = [Component("R307")],
            ComponentImages = [R307Pinout($"{Board}/a.png"), R307Pinout($"{Board}/b.png")],
            Credits =
            [
                new CreditEntry { Category = "Schematics", SubCategory = "Drawn", NameOrHandle = "Dennis" },
                new CreditEntry { Category = "Schematics", SubCategory = "Drawn", NameOrHandle = "Dennis", Contact = "x@example.com" },
            ],
            KiCadImportantSignals =
            [
                new KiCadImportantSignalEntry { DisplayName = "9VAC", KiCadNetName = "9VAC" },
                new KiCadImportantSignalEntry { DisplayName = "9VAC", KiCadNetName = "9VAC" },
            ],
        };

        List<BoardDataProblem> logged = Duplicates(Everything(BoardCheckRows.From(board), $"{Board}/a.png", $"{Board}/b.png"));

        Assert.Equal(
            [BoardWorkbookSchema.SheetComponentImages, BoardWorkbookSchema.SheetComponentImages,
             BoardWorkbookSchema.SheetCredits, BoardWorkbookSchema.SheetCredits,
             BoardWorkbookSchema.SheetKiCadImportantSignals, BoardWorkbookSchema.SheetKiCadImportantSignals],
            logged.Select(problem => problem.Sheet));

        Assert.Empty(Duplicates(BoardDataChecks.Check(BoardCheckRows.From(board), null, BoardCheckScope.Submission)));
    }

    // Not a case of its own: every sheet the server does not already refuse a duplicate on warns,
    // naming that sheet's own identity columns, on the first of them.
    [Theory]
    [InlineData(BoardWorkbookSchema.SheetComponentLocalFiles, BoardWorkbookSchema.ColBoardLabel, "Board label, Name and Region")]
    [InlineData(BoardWorkbookSchema.SheetComponentLinks, BoardWorkbookSchema.ColBoardLabel, "Board label, Name and Region")]
    [InlineData(BoardWorkbookSchema.SheetBoardLocalFiles, BoardWorkbookSchema.ColCategory, "Category, Name and Region")]
    [InlineData(BoardWorkbookSchema.SheetBoardLinks, BoardWorkbookSchema.ColCategory, "Category, Name and Region")]
    [InlineData(BoardWorkbookSchema.SheetCredits, BoardWorkbookSchema.ColCategory, "Category, Sub-category and Name or handle")]
    [InlineData(BoardWorkbookSchema.SheetKiCadImportantSignals, BoardWorkbookSchema.ColDisplayName, "Display name and KiCad net name")]
    public void Each_sheet_names_its_own_identity_columns_in_the_warning(string sheet, string column, string columns)
    {
        var board = new BoardData
        {
            Schematics = [MainSchematic],
            Components = [Component("U1")],
            ComponentLocalFiles = [new() { BoardLabel = "U1", Name = "Datasheet", File = $"{Board}/u1.pdf" }, new() { BoardLabel = "U1", Name = "Datasheet", File = $"{Board}/u1b.pdf" }],
            ComponentLinks = [new() { BoardLabel = "U1", Name = "Info", Url = "https://a.example" }, new() { BoardLabel = "U1", Name = "Info", Url = "https://b.example" }],
            BoardLocalFiles = [new() { Category = "Manuals", Name = "Service", File = $"{Board}/s.pdf" }, new() { Category = "Manuals", Name = "Service", File = $"{Board}/s2.pdf" }],
            BoardLinks = [new() { Category = "Videos", Name = "Repair", Url = "https://a.example" }, new() { Category = "Videos", Name = "Repair", Url = "https://b.example" }],
            Credits = [new() { Category = "Data", SubCategory = "Typed", NameOrHandle = "Dennis" }, new() { Category = "Data", SubCategory = "Typed", NameOrHandle = "Dennis" }],
            KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1" }, new() { DisplayName = "CLK", KiCadNetName = "Net-1" }],
        };

        List<BoardDataProblem> onSheet = Duplicates(Everything(
                BoardCheckRows.From(board),
                $"{Board}/u1.pdf", $"{Board}/u1b.pdf", $"{Board}/s.pdf", $"{Board}/s2.pdf"))
            .Where(problem => problem.Sheet == sheet)
            .ToList();

        Assert.Equal(2, onSheet.Count);
        Assert.All(onSheet, problem => Assert.Equal(column, problem.Column));
        Assert.StartsWith($"Another row on this sheet has the same {columns}.", onSheet[1].Message, StringComparison.Ordinal);
    }

    // Not a case of its own: with three rows the first speaks of "the others".
    [Fact]
    public void The_first_of_three_identical_rows_speaks_of_the_others()
    {
        List<BoardDataProblem> duplicates = Duplicates(Everything(
            Rows(components: [Component("R307")], images: [R307Pinout($"{Board}/a.png"), R307Pinout($"{Board}/b.png"), R307Pinout($"{Board}/c.png")]),
            $"{Board}/a.png", $"{Board}/b.png", $"{Board}/c.png"));

        Assert.Equal(3, duplicates.Count);
        Assert.Contains("so a change in the others is not shown", duplicates[0].Message, StringComparison.Ordinal);
        Assert.Equal(LaterDuplicate, duplicates[2].Message);
    }

    // ###########################################################################################
    // *** A FIELD A ROW OF THE BOARD'S OWN LISTS NEEDS IS AN ERROR ON ITS CELL, TO THE TABLE AND TO
    // THE SERVER ALIKE (owner request, 2026-11-01: "These errors must be flagged in 'Board status' and
    // then it should not be possible to submit the draft board"). *** BoardRowRequired: a board file
    // its category and name (its file is "file.unreferenced"), a link its category, name and address,
    // a credit what it is for and the name or handle, an important signal its net and name - each
    // empty one an error naming its sheet, row and column; which part of a credit and its contact may
    // stay empty.
    // ###########################################################################################
    [Fact]
    public void A_field_a_row_of_the_boards_own_lists_needs_is_an_error_on_its_cell_in_both_scopes()
    {
        var rows = new BoardCheckRows(
            [MainSchematic], [], [], [], [],
            BoardLocalFiles:
            [
                new BoardLocalFileEntry { Category = "Resources", Name = "Manual", File = $"{Board}/manual.pdf" },
                new BoardLocalFileEntry { Name = "Notes", File = $"{Board}/notes.txt" },
            ],
            ComponentLinks: [],
            BoardLinks: [new BoardLinkEntry { Category = "Forums" }],
            Credits:
            [
                new CreditEntry { Category = "Schematic image", NameOrHandle = "Anna" },
                new CreditEntry { SubCategory = "Top", Contact = "a forum" },
            ],
            ImportantSignals: [new KiCadImportantSignalEntry { KiCadNetName = "/RESET", DisplayName = " " }]);

        foreach (BoardCheckScope scope in new[] { BoardCheckScope.Everything, BoardCheckScope.Submission })
        {
            List<(string? Sheet, int Index, string? Column)> empty = BoardDataChecks
                .Check(rows, new SuppliedFileLookup([MainSchematic.SchematicImageFile, $"{Board}/manual.pdf", $"{Board}/notes.txt"]), scope)
                .Where(problem => problem.Code == "row.required_empty")
                .Select(problem =>
                {
                    Assert.Equal(BoardProblemLevel.Error, problem.Level);
                    return (problem.Sheet, problem.Index, problem.Column);
                })
                .ToList();

            Assert.Equal(
                [
                    (BoardWorkbookSchema.SheetBoardLocalFiles, 1, BoardWorkbookSchema.ColCategory),
                    (BoardWorkbookSchema.SheetBoardLinks, 0, BoardWorkbookSchema.ColName),
                    (BoardWorkbookSchema.SheetBoardLinks, 0, BoardWorkbookSchema.ColUrl),
                    (BoardWorkbookSchema.SheetCredits, 1, BoardWorkbookSchema.ColCategory),
                    (BoardWorkbookSchema.SheetCredits, 1, BoardWorkbookSchema.ColNameOrHandle),
                    (BoardWorkbookSchema.SheetKiCadImportantSignals, 0, BoardWorkbookSchema.ColDisplayName),
                ],
                empty);
        }

        // Counted per list for Board status - and every one of those columns is one error.
        BoardRowsIncomplete incomplete = BoardRowRequired.Of(new BoardData
        {
            BoardLocalFiles = [.. rows.BoardLocalFiles],
            BoardLinks = [.. rows.BoardLinks],
            Credits = [.. rows.Credits!],
            KiCadImportantSignals = [.. rows.ImportantSignals!],
        });

        Assert.Equal(new BoardRowsIncomplete(Files: 1, Links: 1, Credits: 1, Signals: 1, Errors: 6), incomplete);
    }

    // ###########################################################################################
    // *** THE ONE RULE FOR A LINK'S ADDRESS (code review, 2026-11-01). *** The component pane marks what
    // this refuses (BoardDataChecks.IsRefusedLink): "www.lemon64.com" and "mailto:" open from their
    // "Open link" (LinkTarget), and nothing on the row said Submit would refuse them. An empty address
    // is no refused one - BoardRowRequired decides whether it must be there. And the error says what
    // to write: "CRT will not open anything else" was untrue of the mailto: its own row opens.
    // ###########################################################################################
    [Theory]
    [InlineData("https://www.lemon64.com", false)]
    [InlineData("http://www.lemon64.com", false)]
    [InlineData("HTTPS://www.lemon64.com", false)]
    [InlineData("www.lemon64.com", true)]
    [InlineData("mailto:anna@example.org", true)]
    [InlineData("ftp://example.org/x", true)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    [InlineData(null, false)]
    public void A_link_the_check_refuses_is_the_one_the_rule_names(string? address, bool refused)
    {
        Assert.Equal(refused, BoardDataChecks.IsRefusedLink(address));

        BoardCheckRows rows = Rows(boardLinks: [new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = address ?? string.Empty }]);
        List<BoardDataProblem> errors = Everything(rows).Where(problem => problem.Code == "link.not_web").ToList();

        Assert.Equal(refused ? 1 : 0, errors.Count);

        if (refused)
        {
            Assert.Contains("Write the whole address", errors[0].Message);
            Assert.DoesNotContain("will not open", errors[0].Message);
        }
    }

    // A submission sent with such a row is refused - the server's own check, as the table says.
    [Fact]
    public void A_submission_with_a_credit_missing_its_name_is_refused()
    {
        var manifest = new SubmissionManifest
        {
            BoardId = Board,
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Summary = "A credit.",
            Rows = new SubmissionRows
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" }],
                Credits = [new CreditEntry { Category = "Schematic image" }],
            },
        };

        ValidationFinding finding = Assert.Single(SubmissionValidator.Validate(manifest, [$"{Board}/main.png"]), finding => finding.Code == "row.required_empty");

        Assert.Equal(ValidationSeverity.Error, finding.Severity);
        Assert.Equal(BoardWorkbookSchema.ColNameOrHandle, finding.Subject);
        Assert.False(SubmissionValidator.CanBeQueued([finding]));
    }

    // ------------------------------------------------------------------ Unfinished work (CRT 4.0 step 1)

    // ###########################################################################################
    // *** THE COMPONENT EDITOR SAVES UNFINISHED WORK, AND THE TABLE SAYS WHAT IS LEFT (2026-10-25).
    // *** It refused a component with no category, and one with a picture CRT cannot show; it saves
    // both now (the plan: "Unfinished work is fine. Only Submit asks for problems to be fixed"), and
    // these warnings say what is left. Warnings, in the table only: the server takes both, so they
    // must never become refusals (Every_error_the_table_finds_is_one_the_server_refuses).
    // ###########################################################################################
    [Fact]
    public void A_component_with_no_category_is_an_error_in_the_table_and_to_the_server()
    {
        // An error since 2026-11-01 (owner decision: "No components must have an empty category").
        BoardCheckRows rows = Rows(components: [new ComponentEntry { BoardLabel = "U1" }, Component("U2")]);

        BoardDataProblem error = Assert.Single(Everything(rows), problem => problem.Code == "component.no_category");

        Assert.Equal(BoardProblemLevel.Error, error.Level);
        Assert.Equal(BoardWorkbookSchema.SheetComponents, error.Sheet);
        Assert.Equal(0, error.Index);
        Assert.Equal(BoardWorkbookSchema.ColCategory, error.Column);
        Assert.Contains("[U1]", error.Message, StringComparison.Ordinal);

        Assert.Contains(BoardDataChecks.Check(rows, null, BoardCheckScope.Submission), problem => problem.Code == "component.no_category");
    }

    // ###########################################################################################
    // A PIN MUST BE ON THE COMPONENT'S CASE (owner request, 2026-11-01: "if I e.g. put pin 48 on a 24 pin
    // package, this is an error") - a scope baseline's and a pinout's, to the table and the server alike.
    // ###########################################################################################
    [Fact]
    public void A_scope_baseline_on_a_pin_the_case_does_not_have_is_an_error_on_its_cell()
    {
        BoardCheckRows rows = Rows(
            components: [new ComponentEntry { BoardLabel = "U1", Category = "IC", Package = "DIP-24" }],
            images:
            [
                new ComponentImageEntry { BoardLabel = "U1", Pin = "24", TimeDiv = "1ms", File = $"{Board}/a.png" },
                new ComponentImageEntry { BoardLabel = "U1", Pin = "48", TimeDiv = "1ms", File = $"{Board}/b.png" },
            ]);

        BoardDataProblem error = Assert.Single(Everything(rows), problem => problem.Code == "pin.outside_package");

        Assert.Equal(BoardProblemLevel.Error, error.Level);
        Assert.Equal(BoardWorkbookSchema.SheetComponentImages, error.Sheet);
        Assert.Equal(1, error.Index);
        Assert.Equal(BoardWorkbookSchema.ColPin, error.Column);
        Assert.Contains("DIP-24", error.Message, StringComparison.Ordinal);
        Assert.Contains("pin 48", error.Message, StringComparison.Ordinal);

        Assert.Contains(BoardDataChecks.Check(rows, null, BoardCheckScope.Submission), problem => problem.Code == "pin.outside_package");
    }

    [Fact]
    public void A_pin_of_the_pinout_the_case_does_not_have_is_an_error_too()
    {
        BoardCheckRows rows = Rows(components: [new ComponentEntry { BoardLabel = "U1", Category = "IC", Package = "TO-92" }]) with
        {
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U1", Pin = "3" }, new ComponentPinEntry { BoardLabel = "U1", Pin = "4" }],
        };

        BoardDataProblem error = Assert.Single(BoardDataChecks.Check(rows, null, BoardCheckScope.Submission), problem => problem.Code == "pin.outside_package");

        Assert.Equal(BoardWorkbookSchema.SheetComponentPins, error.Sheet);
        Assert.Equal(1, error.Index);
    }

    // The submission's rows reach the check: a server taking only the table's word would miss them.
    [Fact]
    public void The_servers_rows_carry_the_pins_to_the_check()
    {
        var submitted = new SubmissionRows
        {
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC", Package = "DIP-8" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U1", Pin = "9" }],
        };

        Assert.Contains(
            BoardDataChecks.Check(BoardCheckRows.From(submitted), null, BoardCheckScope.Submission),
            problem => problem.Code == "pin.outside_package");
    }

    // A case CRT cannot count, no package, a pin named and not numbered: nothing to say. "2 [64K]" is how
    // four published baselines write pin 2.
    [Theory]
    [InlineData("", "99")]
    [InlineData("SIP", "99")]
    [InlineData("Other", "99")]
    [InlineData("DIP-8", "A1")]
    [InlineData("DIP-8", "2 [64K]")]
    [InlineData("DIP-8", "12V")]
    public void What_cannot_be_told_is_no_problem(string package, string pin)
    {
        BoardCheckRows rows = Rows(
            components: [new ComponentEntry { BoardLabel = "U1", Category = "IC", Package = package }],
            images: [new ComponentImageEntry { BoardLabel = "U1", Pin = pin, TimeDiv = "1ms", File = $"{Board}/a.png" }]);

        Assert.DoesNotContain(Everything(rows), problem => problem.Code == "pin.outside_package");
    }

    // Each region has a case of its own, when the component has a row of its own per region - a row for
    // every region is on all of them, so every one of them has to refuse it.
    [Fact]
    public void A_pin_is_judged_by_the_case_of_its_own_region()
    {
        ComponentEntry[] twins =
        [
            new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "PAL", Package = "DIP-40" },
            new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "NTSC", Package = "DIP-24" },
        ];

        Assert.Null(BoardDataChecks.PackageRefusing(twins, "U1", "PAL", "30"));
        Assert.Equal("DIP-24", BoardDataChecks.PackageRefusing(twins, "U1", "NTSC", "30"));
        Assert.Null(BoardDataChecks.PackageRefusing(twins, "U1", "", "30"));
        Assert.Equal("DIP-40", BoardDataChecks.PackageRefusing(twins, "u1", "", "41"));
        Assert.Null(BoardDataChecks.PackageRefusing(twins, "U99", "", "41"));
    }

    // ###########################################################################################
    // *** Security (code review, 2026-11-01): A SUBMISSION'S REGIONS REACH THE PUBLISH AS SENT. *** A JSON null made every
    // Approve of it throw. Refused where every rule is, so the table says it too. A board may name regions of its own since
    // 2026-11-09 ("x" was refused until then) - by BoardRegions.NameProblem's rule, each once.
    // ###########################################################################################
    [Theory]
    [InlineData("x|y")]
    [InlineData("pal")]
    [InlineData("")]
    [InlineData(null)]
    public void A_region_name_the_rule_refuses_is_refused(string? region)
    {
        var submitted = new SubmissionRows { Regions = ["PAL", region!] };

        BoardDataProblem problem = Assert.Single(
            BoardDataChecks.Check(BoardCheckRows.From(submitted), null, BoardCheckScope.Submission),
            found => found.Code == "regions.unknown");

        Assert.Equal(BoardProblemLevel.Error, problem.Level);

        Assert.Contains(
            BoardDataChecks.Check(BoardCheckRows.From(new BoardData { Regions = ["PAL", region!] }), null, BoardCheckScope.Everything),
            found => found.Code == "regions.unknown");
    }

    // ###########################################################################################
    // A BOARD WITH REGIONS COMES IN TWO AT LEAST (code review, 2026-11-09): one region passed and was published - a board
    // CRT's own "Board: Regions" never makes (BoardRegions.FewestRegions). Refused on both sides, as a name is.
    // ###########################################################################################
    [Fact]
    public void A_board_of_one_region_alone_is_refused()
    {
        BoardDataProblem problem = Assert.Single(
            BoardDataChecks.Check(BoardCheckRows.From(new SubmissionRows { Regions = ["EU"] }), null, BoardCheckScope.Submission),
            found => found.Code == "regions.unknown");

        Assert.Equal(BoardProblemLevel.Error, problem.Level);
        Assert.Contains("[EU] alone", problem.Message, StringComparison.Ordinal);

        Assert.Contains(
            BoardDataChecks.Check(BoardCheckRows.From(new BoardData { Regions = ["EU"] }), null, BoardCheckScope.Everything),
            found => found.Code == "regions.unknown");

        Assert.DoesNotContain(
            BoardDataChecks.Check(BoardCheckRows.From(new SubmissionRows { Regions = ["EU", "US"] }), null, BoardCheckScope.Submission),
            found => found.Code == "regions.unknown");
    }

    // Null keeps what is published, empty clears it, and PAL and NTSC are what CRT sends.
    [Fact]
    public void What_a_board_may_say_about_its_regions_is_no_problem()
    {
        foreach (List<string>? regions in new List<string>?[] { null, [], ["PAL", "NTSC"] })
        {
            Assert.DoesNotContain(
                BoardDataChecks.Check(BoardCheckRows.From(new SubmissionRows { Regions = regions }), null, BoardCheckScope.Submission),
                found => found.Code == "regions.unknown");
        }
    }

    [Fact]
    public void A_component_image_CRT_cannot_show_is_a_warning_in_the_table_and_nothing_to_the_server()
    {
        BoardCheckRows rows = Rows(
            components: [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            images: [Image("U1", $"{Board}/U1.pdf"), Image("U1", $"{Board}/U1.png")]);

        BoardDataProblem warning = Assert.Single(Everything(rows, $"{Board}/U1.pdf", $"{Board}/U1.png"), problem => problem.Code == "image.not_displayable");

        Assert.Equal(BoardProblemLevel.Warning, warning.Level);
        Assert.Equal(0, warning.Index);
        Assert.Equal(BoardWorkbookSchema.ColFile, warning.Column);
        Assert.Contains(".png", warning.Message, StringComparison.Ordinal);

        Assert.DoesNotContain(BoardDataChecks.Check(rows, null, BoardCheckScope.Submission), problem => problem.Code == "image.not_displayable");
    }

    // ------------------------------------------------------------------ CRT 4.0's own data (2026-10-25)

    private static BoardCheckRows Crt4Rows(
        IReadOnlyList<ComponentEntry>? components = null,
        IReadOnlyList<ComponentImageEntry>? images = null,
        IReadOnlyList<ComponentPinEntry>? pins = null,
        IReadOnlyList<ComponentReplacementEntry>? replacements = null) =>
        Rows(components: components ?? [Component("U8")], images: images) with
        {
            ComponentPins = pins ?? [],
            ComponentReplacements = replacements ?? [],
        };

    // ###########################################################################################
    // Every rule about CRT 4.0's data is a WARNING in the table and nothing to the server - a new
    // local rule always is, and unfinished work is fine. Each one, on the cell it is about.
    // ###########################################################################################
    [Theory]
    [InlineData("component.long_description", BoardWorkbookSchema.SheetComponents, BoardWorkbookSchema.ColDescription)]
    [InlineData("pin.orphan", BoardWorkbookSchema.SheetComponentPins, BoardWorkbookSchema.ColBoardLabel)]
    [InlineData("pin.unnumbered", BoardWorkbookSchema.SheetComponentPins, BoardWorkbookSchema.ColPin)]
    [InlineData("pin.direction", BoardWorkbookSchema.SheetComponentPins, BoardWorkbookSchema.ColDirection)]
    [InlineData("replacement.orphan", BoardWorkbookSchema.SheetComponentReplacements, BoardWorkbookSchema.ColBoardLabel)]
    [InlineData("replacement.no_part_number", BoardWorkbookSchema.SheetComponentReplacements, BoardWorkbookSchema.ColPartNumber)]
    [InlineData("replacement.kind", BoardWorkbookSchema.SheetComponentReplacements, BoardWorkbookSchema.ColKind)]
    public void Each_rule_about_CRT_4_data_is_a_warning_on_its_cell_and_nothing_to_the_server(string code, string sheet, string column)
    {
        BoardCheckRows rows = Crt4Rows(
            components: [new ComponentEntry { BoardLabel = "U8", Category = "IC", Description = new string('x', ComponentVocabulary.ShortDescriptionLimit + 1) }],
            images: [new ComponentImageEntry { BoardLabel = "U8", Pin = "2", Name = "PA0", File = $"{Board}/U1.png", TimeDiv = "1mS" }],
            pins:
            [
                new ComponentPinEntry { BoardLabel = "U99", Pin = "1", Name = "VCC", Direction = "Power" },
                new ComponentPinEntry { BoardLabel = "U8", Pin = "", Name = "RES", Direction = "Input" },
                new ComponentPinEntry { BoardLabel = "U8", Pin = "3", Name = "PA1", Direction = "Sideways" },
            ],
            replacements:
            [
                new ComponentReplacementEntry { BoardLabel = "U99", PartNumber = "6526A", Kind = "Drop-in replacement" },
                new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "", Kind = "Drop-in replacement" },
                new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "8521", Kind = "Better" },
            ]);

        BoardDataProblem warning = Assert.Single(Everything(rows, $"{Board}/U1.png"), problem => problem.Code == code);

        Assert.Equal(BoardProblemLevel.Warning, warning.Level);
        Assert.Equal(sheet, warning.Sheet);
        Assert.Equal(column, warning.Column);
        Assert.DoesNotContain(BoardDataChecks.Check(rows, null, BoardCheckScope.Submission), problem => problem.Code == code);
    }

    // Empty is fine everywhere - unfinished work - and the words compare ignoring capitals; a short
    // description of exactly the limit fits.
    [Fact]
    public void Empty_or_known_values_and_a_description_at_the_limit_are_no_problem()
    {
        BoardCheckRows rows = Crt4Rows(
            components: [new ComponentEntry { BoardLabel = "U8", Category = "IC", Description = new string('x', ComponentVocabulary.ShortDescriptionLimit) }],
            images: [new ComponentImageEntry { BoardLabel = "U8", Pin = "2", Name = "PA0", File = $"{Board}/U1.png", TimeDiv = "1mS" }],
            pins: [new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", Direction = "" }, new ComponentPinEntry { BoardLabel = "u8", Pin = "2", Direction = "ground" }],
            replacements: [new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "6526A", Kind = "" }]);

        IReadOnlyList<BoardDataProblem> problems = Everything(rows, $"{Board}/U1.png");

        Assert.DoesNotContain(problems, problem => problem.Code.StartsWith("pin.", StringComparison.Ordinal));
        Assert.DoesNotContain(problems, problem => problem.Code.StartsWith("replacement.", StringComparison.Ordinal));
        Assert.DoesNotContain(problems, problem => problem.Code == "component.long_description");
    }

    // Two pins with one number on one component are warned about, as any two rows of one identity.
    [Fact]
    public void Two_pins_with_one_number_are_a_duplicate_warning()
    {
        BoardCheckRows rows = Crt4Rows(pins:
        [
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" },
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "GND" },
        ]);

        List<BoardDataProblem> duplicates = Duplicates(Everything(rows));

        Assert.Equal(2, duplicates.Count);
        Assert.All(duplicates, problem => Assert.Equal(BoardWorkbookSchema.SheetComponentPins, problem.Sheet));
    }

    // ###########################################################################################
    // AN OSCILLOSCOPE BASELINE'S PIN, SIGNAL NAME AND IMAGE ARE MANDATORY, ITS SETTINGS NOT (owner
    // request, 2026-11-02) - errors on both sides, so the table marks the cell and the server refuses
    // it. A picture needs neither.
    // ###########################################################################################
    [Fact]
    public void An_oscilloscope_baseline_needs_its_signal_name_and_image_but_no_settings()
    {
        var board = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "PA0", File = "Scope baseline/a.png" },
                new ComponentImageEntry { BoardLabel = "U8", Pin = "2", File = "Scope baseline/b.png" },
                new ComponentImageEntry { BoardLabel = "U8", Pin = "3", Name = "PA2", Note = "A note and no image" },
                new ComponentImageEntry { BoardLabel = "U8", Name = "Top", Note = "A picture: a note is enough" },
            ],
        };

        List<BoardDataProblem> problems = BoardDataChecks.Check(BoardCheckRows.From(board), null, BoardCheckScope.Submission)
            .Where(problem => problem.Code == "baseline.required_empty")
            .ToList();

        Assert.Equal(
            [(1, BoardWorkbookSchema.ColName), (2, BoardWorkbookSchema.ColFile)],
            problems.Select(problem => (problem.Index, problem.Column!)));
        Assert.All(problems, problem => Assert.Equal(BoardProblemLevel.Error, problem.Level));
        Assert.All(problems, problem => Assert.Equal(BoardWorkbookSchema.SheetComponentImages, problem.Sheet));
    }

    // "Datasheet" of PAL and of NTSC are two rows, not a duplicate (owner request, 2026-11-02).
    [Fact]
    public void A_file_or_link_of_one_name_in_each_region_is_no_duplicate()
    {
        var board = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
            BoardLocalFiles =
            [
                new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = "a.pdf", Region = "PAL" },
                new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = "b.pdf", Region = "NTSC" },
            ],
            ComponentLinks =
            [
                new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.org/a", Region = "PAL" },
                new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.org/b", Region = "NTSC" },
            ],
        };

        Assert.DoesNotContain(
            BoardDataChecks.Check(BoardCheckRows.From(board), null, BoardCheckScope.Everything),
            problem => problem.Message.Contains("Name and Region", StringComparison.Ordinal));
    }

    // ###########################################################################################
    // ONE PHYSICAL COMPONENT, ONE WARNING (owner decision, 2026-11-04: "it should only be reported once as it
    // is ONE physical component only"). A component different per region is a row per region, and a mark
    // belongs to its label, so marking it once clears both: told once, on its first row. Capitals aside, as
    // the marks are looked up.
    // ###########################################################################################
    [Fact]
    public void A_component_with_a_row_per_region_is_told_once_that_it_is_marked_nowhere()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U19", Region = "PAL", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
                new ComponentEntry { BoardLabel = "u19", Region = "NTSC", Category = "IC" },
            ],
        };

        List<BoardDataProblem> unmarked = BoardDataChecks
            .Check(BoardCheckRows.From(board), files: null, BoardCheckScope.Everything)
            .Where(problem => problem.Code == "component.no_highlight")
            .ToList();

        Assert.Equal(["U19", "C1"], unmarked.Select(problem => problem.Subject));
        Assert.Equal([0, 1], unmarked.Select(problem => problem.Index));
    }

    // A field is named as it reads on the screen, in quotes (owner request, 2026-11-04): "Name or handle" is
    // one field, and 'has no Name or handle' read as two things it could have.
    [Fact]
    public void A_field_a_row_needs_is_named_in_quotes()
    {
        BoardDataProblem problem = Assert.Single(
            BoardDataChecks.Check(
                BoardCheckRows.From(new BoardData { Credits = [new CreditEntry { Category = "Schematic image", SubCategory = "Top" }] }),
                files: null,
                BoardCheckScope.Everything),
            found => found.Code == "row.required_empty");

        Assert.StartsWith("This credit has no \"Name or handle\".", problem.Message, StringComparison.Ordinal);
    }

}
