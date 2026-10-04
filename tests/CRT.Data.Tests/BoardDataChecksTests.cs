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

    private static ComponentEntry Component(string label, string region = "", string partNumber = "", string technicalName = "") =>
        new() { BoardLabel = label, Region = region, PartNumber = partNumber, TechnicalNameOrValue = technicalName };

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
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Bad", Url = "ftp://example.com" }]
        };

        // Everything the rows name, less the one that is missing - and with "Main.PNG" carried as
        // "main.png", the spelling the second image's row gets wrong.
        string[] carried = [$"{Board}/main.png", $"{Board}/main2.png", $"{Board}/x.png", $"{Board}/tool.exe"];

        var manifest = new SubmissionManifest
        {
            SystemId = Board,
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Summary = "Every defect at once.",
            Rows = rows,
            Files = carried.Select(path => new SubmissionFile { Path = path, Sha256 = new string('a', 64), SizeBytes = 1 }).ToList()
        };

        HashSet<string> serverErrors = SubmissionValidator.Validate(manifest, carried)
            .Concat(SubmissionFileRules.ValidateManifestFiles(manifest, tree: null))
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
    [InlineData(BoardWorkbookSchema.SheetComponentLocalFiles, BoardWorkbookSchema.ColBoardLabel, "Board label and Name")]
    [InlineData(BoardWorkbookSchema.SheetComponentLinks, BoardWorkbookSchema.ColBoardLabel, "Board label and Name")]
    [InlineData(BoardWorkbookSchema.SheetBoardLocalFiles, BoardWorkbookSchema.ColCategory, "Category and Name")]
    [InlineData(BoardWorkbookSchema.SheetBoardLinks, BoardWorkbookSchema.ColCategory, "Category and Name")]
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
}
