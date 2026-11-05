using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributeProblems - a draft's errors and warnings, each on the entry of the Contribute tab it is about
// (owner request, 2026-11-04: "move the error and warning badges to where the problem is").
// ###########################################################################################
public sealed class ContributeProblemsTests
{
    private static BoardDataProblem Problem(string? sheet, int index = 0, string code = "some.code", BoardProblemLevel level = BoardProblemLevel.Error) =>
        new(level, code, string.Empty, "What is wrong", sheet, index, null);

    private static BoardProblemsRead ChecksOf(BoardData board)
    {
        IReadOnlyList<BoardDataProblem> problems = BoardDataChecks.Check(BoardCheckRows.From(board), null, BoardCheckScope.Everything);

        return new BoardProblemsRead(BoardProblemCounts.Of(problems), problems, board);
    }

    [Theory]
    [InlineData(BoardWorkbookSchema.SheetBoardSchematics, nameof(ContributeEntry.SchematicImages))]
    [InlineData(BoardWorkbookSchema.SheetComponents, nameof(ContributeEntry.ComponentBasics))]
    [InlineData(BoardWorkbookSchema.SheetComponentLocalFiles, nameof(ContributeEntry.ComponentFiles))]
    [InlineData(BoardWorkbookSchema.SheetComponentLinks, nameof(ContributeEntry.ComponentLinks))]
    [InlineData(BoardWorkbookSchema.SheetComponentPins, nameof(ContributeEntry.ComponentPinout))]
    [InlineData(BoardWorkbookSchema.SheetComponentReplacements, nameof(ContributeEntry.ComponentReplacements))]
    [InlineData(BoardWorkbookSchema.SheetBoardLocalFiles, nameof(ContributeEntry.Files))]
    [InlineData(BoardWorkbookSchema.SheetBoardLinks, nameof(ContributeEntry.Links))]
    [InlineData(BoardWorkbookSchema.SheetCredits, nameof(ContributeEntry.Credits))]
    [InlineData(BoardWorkbookSchema.SheetKiCadImportantSignals, nameof(ContributeEntry.ImportantSignals))]
    public void A_problem_goes_on_the_entry_of_its_sheet(string sheet, string entry)
    {
        Assert.Equal(Enum.Parse<ContributeEntry>(entry), ContributeProblems.EntryOf(Problem(sheet), new BoardData()));
    }

    // A new sheet with no place here would put its problems on "Overview", saying nothing of where.
    [Fact]
    public void Every_sheet_of_a_board_is_some_entrys()
    {
        Assert.All(
            BoardWorkbookSchema.AllSheets,
            sheet => Assert.NotEqual(ContributeEntry.Overview, ContributeProblems.EntryOf(Problem(sheet.SheetName), new BoardData())));
    }

    // One sheet, two entries, told apart as everything tells them apart: a pin makes a baseline.
    [Fact]
    public void A_component_image_goes_on_pictures_or_baselines_by_what_its_row_is()
    {
        var board = new BoardData
        {
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Name = "Top" },
                new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" },
            ],
        };

        Assert.Equal(ContributeEntry.ComponentPictures, ContributeProblems.EntryOf(Problem(BoardWorkbookSchema.SheetComponentImages, 0), board));
        Assert.Equal(ContributeEntry.ComponentBaselines, ContributeProblems.EntryOf(Problem(BoardWorkbookSchema.SheetComponentImages, 1), board));
    }

    [Fact]
    public void What_is_on_no_sheet_goes_where_its_code_says_or_on_the_overview()
    {
        Assert.Equal(ContributeEntry.Regions, ContributeProblems.EntryOf(Problem(null, -1, "regions.unknown"), null));
        Assert.Equal(ContributeEntry.SchematicImages, ContributeProblems.EntryOf(Problem(null, -1, "highlight.orphan"), null));
        Assert.Equal(ContributeEntry.Overview, ContributeProblems.EntryOf(Problem(null, -1), null));
    }

    // ###########################################################################################
    // Through the real checks, a real problem of each kind: on its entry, of its component - and none
    // lost, so the badges add up to what the checks found.
    // ###########################################################################################
    [Fact]
    public void What_the_checks_find_lands_on_the_entries_and_nothing_is_lost()
    {
        var board = new BoardData
        {
            Regions = ["PAL", "PLUTO"],
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U9", FriendlyName = "No category" },
            ],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U8", Pin = "48", Name = "OOPS" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Bad", Url = "ftp://example.com" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", File = "Scope baseline/a.png" }],
            Credits = [new CreditEntry { Category = "Schematic image" }],
            BoardLinks = [new BoardLinkEntry { Name = "Forum" }],
        };

        BoardProblemsRead read = ChecksOf(board);
        List<EntryProblem> placed = ContributeProblems.Placed(read);

        Assert.Equal(read.Problems.Count, placed.Count);

        Assert.Contains(placed, item => item.Problem.Code == "regions.unknown" && item.Entry == ContributeEntry.Regions);
        Assert.Contains(placed, item => item.Problem.Code == "component.no_category" && item.Entry == ContributeEntry.ComponentBasics && item.ComponentLabel == "U9");
        Assert.Contains(placed, item => item.Problem.Code == "pin.outside_package" && item.Entry == ContributeEntry.ComponentPinout && item.ComponentLabel == "U8");
        Assert.Contains(placed, item => item.Problem.Code == "link.not_web" && item.Entry == ContributeEntry.ComponentLinks && item.ComponentLabel == "U8");
        Assert.Contains(placed, item => item.Problem.Code == "baseline.required_empty" && item.Entry == ContributeEntry.ComponentBaselines && item.ComponentLabel == "U8");
        Assert.Contains(placed, item => item.Problem.Code == "row.required_empty" && item.Entry == ContributeEntry.Credits && item.ComponentLabel.Length == 0);
        Assert.Contains(placed, item => item.Problem.Code == "row.required_empty" && item.Entry == ContributeEntry.Links);

        Dictionary<ContributeEntry, BoardProblemCounts> counts = ContributeProblems.CountsShown(placed, string.Empty);

        Assert.Equal(read.Counts, counts[ContributeEntry.Overview]);
        Assert.False(counts.ContainsKey(ContributeEntry.Name));

        List<BoardProblemCounts> below = counts.Where(count => count.Key != ContributeEntry.Overview).Select(count => count.Value).ToList();

        Assert.Equal(read.Counts.Errors, below.Sum(count => count.Errors));
        Assert.Equal(read.Counts.Warnings, below.Sum(count => count.Warnings));
    }

    // ###########################################################################################
    // With a component selected its pages have that one's alone (owner request, 2026-11-04) - "Overview" has
    // them all whatever is selected, and a board's own entries are not a component's to narrow.
    // ###########################################################################################
    private static EntryProblem Placed(ContributeEntry entry, string component, BoardProblemLevel level = BoardProblemLevel.Error) =>
        new(Problem(BoardWorkbookSchema.SheetComponents, 0, "some.code", level), entry, component);

    [Fact]
    public void With_a_component_selected_only_its_own_counts_on_its_pages()
    {
        List<EntryProblem> placed =
        [
            Placed(ContributeEntry.ComponentBasics, "U8"),
            Placed(ContributeEntry.ComponentBasics, "U9", BoardProblemLevel.Warning),
            Placed(ContributeEntry.ComponentLinks, "U9"),
            Placed(ContributeEntry.Credits, string.Empty),
        ];

        Dictionary<ContributeEntry, BoardProblemCounts> u8 = ContributeProblems.CountsShown(placed, "u8");

        Assert.Equal(new BoardProblemCounts(1, 0), u8[ContributeEntry.ComponentBasics]);
        Assert.False(u8.ContainsKey(ContributeEntry.ComponentLinks));
        Assert.Equal(new BoardProblemCounts(1, 0), u8[ContributeEntry.Credits]);
        Assert.Equal(new BoardProblemCounts(3, 1), u8[ContributeEntry.Overview]);

        Dictionary<ContributeEntry, BoardProblemCounts> none = ContributeProblems.CountsShown(placed, string.Empty);

        Assert.Equal(new BoardProblemCounts(1, 1), none[ContributeEntry.ComponentBasics]);
        Assert.Equal(new BoardProblemCounts(1, 0), none[ContributeEntry.ComponentLinks]);
    }

    // A row of no component on a component's page is in nobody's way while one is selected.
    [Fact]
    public void What_a_page_shows_goes_by_the_component_selected()
    {
        List<EntryProblem> placed =
        [
            Placed(ContributeEntry.ComponentPinout, "U8"),
            Placed(ContributeEntry.ComponentPinout, string.Empty),
            Placed(ContributeEntry.Credits, string.Empty),
        ];

        Assert.Equal(["U8"], ContributeProblems.OnEntry(placed, ContributeEntry.ComponentPinout, "U8").Select(item => item.ComponentLabel));
        Assert.Equal(2, ContributeProblems.OnEntry(placed, ContributeEntry.ComponentPinout, string.Empty).Count);
        Assert.Single(ContributeProblems.OnEntry(placed, ContributeEntry.Credits, "U8"));
        Assert.Empty(ContributeProblems.OnEntry(placed, ContributeEntry.ComponentLinks, null));
    }

    [Fact]
    public void Nothing_wrong_has_no_badge_at_all()
    {
        Assert.Empty(ContributeProblems.CountsShown([], "U8"));
    }

    // Its row in the list says all its pages together, its label in any capitals; what is of no component is
    // on no row.
    [Fact]
    public void A_component_is_counted_over_all_its_pages()
    {
        List<EntryProblem> placed =
        [
            Placed(ContributeEntry.ComponentBasics, "U8"),
            Placed(ContributeEntry.ComponentLinks, "u8", BoardProblemLevel.Warning),
            Placed(ContributeEntry.ComponentPinout, "U9"),
            Placed(ContributeEntry.ComponentPinout, string.Empty),
            Placed(ContributeEntry.Credits, string.Empty),
        ];

        Dictionary<string, BoardProblemCounts> counts = ContributeProblems.CountsByComponent(placed);

        Assert.Equal(new BoardProblemCounts(1, 1), counts["U8"]);
        Assert.Equal(new BoardProblemCounts(1, 1), counts["u8"]);
        Assert.Equal(new BoardProblemCounts(1, 0), counts["U9"]);
        Assert.Equal(2, counts.Count);
    }

    [Fact]
    public void Errors_stand_before_warnings()
    {
        var read = new BoardProblemsRead(
            new BoardProblemCounts(1, 1),
            [Problem(BoardWorkbookSchema.SheetBoardLinks, 0, "a", BoardProblemLevel.Warning), Problem(BoardWorkbookSchema.SheetBoardLinks, 0, "b")],
            new BoardData());

        Assert.Equal(["b", "a"], ContributeProblems.Placed(read).Select(item => item.Problem.Code));
    }

    // A row left behind by a component the board does not have: nothing could open it.
    [Fact]
    public void A_row_of_a_component_the_board_does_not_have_names_no_component()
    {
        var board = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U99", Pin = "1" }],
        };

        var read = new BoardProblemsRead(new BoardProblemCounts(0, 1), [Problem(BoardWorkbookSchema.SheetComponentPins, 0, "pin.orphan", BoardProblemLevel.Warning)], board);

        EntryProblem placed = Assert.Single(ContributeProblems.Placed(read));

        Assert.Equal(ContributeEntry.ComponentPinout, placed.Entry);
        Assert.Equal(string.Empty, placed.ComponentLabel);
    }

    [Fact]
    public void A_read_with_no_board_is_placed_by_its_sheets_alone()
    {
        var read = new BoardProblemsRead(new BoardProblemCounts(1, 0), [Problem(BoardWorkbookSchema.SheetComponents, 5)], null);

        EntryProblem placed = Assert.Single(ContributeProblems.Placed(read));

        Assert.Equal(ContributeEntry.ComponentBasics, placed.Entry);
        Assert.Equal(string.Empty, placed.ComponentLabel);
        Assert.Empty(ContributeProblems.Placed(BoardProblemsRead.None));
    }

    [Theory]
    [InlineData(1, 0, "1 error here")]
    [InlineData(0, 1, "1 warning here")]
    [InlineData(2, 3, "2 errors and 3 warnings here")]
    public void The_heading_says_what_is_here(int errors, int warnings, string expected)
    {
        Assert.Equal(expected, ContributeProblems.Heading(new BoardProblemCounts(errors, warnings)));
    }

    [Fact]
    public void A_panels_own_panel_says_how_many_of_each()
    {
        Assert.Equal("1 warning", ContributeProblems.Said(new BoardProblemCounts(0, 1)));
        Assert.Equal("2 errors and 3 warnings", ContributeProblems.Said(new BoardProblemCounts(2, 3)));
    }

    [Fact]
    public void Under_it_is_what_they_mean_and_on_a_components_page_of_which_components()
    {
        var withError = new BoardProblemCounts(1, 0);
        var warningsAlone = new BoardProblemCounts(0, 2);

        Assert.Contains("stops the draft being submitted", ContributeProblems.Help(withError, ofComponents: false), StringComparison.Ordinal);
        Assert.Contains("does not stop", ContributeProblems.Help(warningsAlone, ofComponents: false), StringComparison.Ordinal);
        Assert.StartsWith("Of all the board's components.", ContributeProblems.Help(withError, ofComponents: true), StringComparison.Ordinal);
        Assert.DoesNotContain("components", ContributeProblems.Help(withError, ofComponents: false), StringComparison.Ordinal);
    }

    [Fact]
    public void What_a_line_says()
    {
        Assert.Equal("Error", ContributeProblems.LevelWord(BoardProblemLevel.Error));
        Assert.Equal("Warning", ContributeProblems.LevelWord(BoardProblemLevel.Warning));
        Assert.Equal("Open [U8]", ContributeProblems.OpenLabel("U8"));
        Assert.Contains("12 more", ContributeProblems.More(12), StringComparison.Ordinal);
        Assert.Contains("1 error here", ContributeProblems.ErrorsTip(1), StringComparison.Ordinal);
        Assert.Contains("2 warnings here", ContributeProblems.WarningsTip(2), StringComparison.Ordinal);
    }

    // ###########################################################################################
    // A MESSAGE NAMING NO ROW GETS ITS ROW'S NAME (owner report, 2026-11-04). Beside the cell the table marks
    // 'This credit has no "Name or handle"' is plain; three of them in a list say nothing. A message that names
    // its row keeps its own words, and a row with nothing filled in goes by its number.
    // ###########################################################################################
    [Fact]
    public void What_is_said_of_a_problem_names_the_row_its_message_does_not()
    {
        var board = new BoardData
        {
            Credits =
            [
                new CreditEntry { Category = "Schematic image", SubCategory = "Top" },
                new CreditEntry(),
            ],
        };

        BoardDataProblem OnCredit(int index, string message) =>
            new(BoardProblemLevel.Error, "row.required_empty", string.Empty, message, BoardWorkbookSchema.SheetCredits, index, BoardWorkbookSchema.ColNameOrHandle);

        Assert.Equal("[Schematic image / Top] This credit has no \"Name or handle\".", ContributeProblems.TextOf(OnCredit(0, "This credit has no \"Name or handle\"."), board));
        Assert.Equal("Row 2: This credit has no \"Name or handle\".", ContributeProblems.TextOf(OnCredit(1, "This credit has no \"Name or handle\"."), board));
        Assert.Equal("[U2] has no category.", ContributeProblems.TextOf(OnCredit(0, "[U2] has no category."), board));

        // Past the rows, or with no board read: its message as it stands.
        Assert.Equal("Gone.", ContributeProblems.TextOf(OnCredit(5, "Gone."), board));
        Assert.Equal("Gone.", ContributeProblems.TextOf(OnCredit(0, "Gone."), null));
    }

    // ###########################################################################################
    // A ROW OF ONE REGION IS SAID TO BE OF IT, and "Open" shows that region (code review, 2026-11-04). A message
    // naming its row kept its own words, so with PAL shown "This pin of [U7] says pin 48" was about a row on no
    // page, and nothing said so. "Files" and "Links" list every row, so nothing has to be switched for them.
    // ###########################################################################################
    [Fact]
    public void What_is_said_of_a_problem_names_the_region_of_its_row_and_which_region_shows_it()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U7", FriendlyName = "VIC", Category = "IC", Region = "PAL", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U7", FriendlyName = "VIC", Category = "IC", Region = "NTSC", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC", Package = "DIP-40" },
            ],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U7", Region = "NTSC", Pin = "48", Name = "OOPS" },
                new ComponentPinEntry { BoardLabel = "U8", Pin = "48", Name = "OOPS" },
            ],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U7", Region = "NTSC", Name = "Bad", Url = "ftp://example.com" }],
        };

        List<EntryProblem> placed = ContributeProblems.Placed(ChecksOf(board));

        EntryProblem ofOneRegion = placed.Single(item => item.Problem.Code == "pin.outside_package" && item.ComponentLabel == "U7");

        Assert.Equal("NTSC", ofOneRegion.Region);
        Assert.StartsWith("NTSC: This pin of [U7]", ofOneRegion.Text, StringComparison.Ordinal);
        Assert.Equal(BoardRegions.Ntsc, ContributeProblems.RegionToOpen(ofOneRegion));

        // Of every region: nothing to say, nothing to switch.
        EntryProblem ofEvery = placed.Single(item => item.Problem.Code == "pin.outside_package" && item.ComponentLabel == "U8");

        Assert.Equal(string.Empty, ofEvery.Region);
        Assert.StartsWith("This pin of [U8]", ofEvery.Text, StringComparison.Ordinal);
        Assert.Equal(string.Empty, ContributeProblems.RegionToOpen(ofEvery));

        // Not marked is of the component, on whichever row it is told: "it is ONE physical component only".
        EntryProblem notMarked = placed.Single(item => item.Problem.Code == "component.no_highlight" && item.ComponentLabel == "U7");

        Assert.Equal(string.Empty, notMarked.Region);
        Assert.StartsWith("Component [U7]", notMarked.Text, StringComparison.Ordinal);
        Assert.Equal(string.Empty, ContributeProblems.RegionToOpen(notMarked));

        // Its page lists it whatever CRT shows - its name holds the region as it is.
        EntryProblem link = placed.Single(item => item.Problem.Code == "link.not_web");

        Assert.Contains("NTSC", link.Text, StringComparison.Ordinal);
        Assert.Equal(string.Empty, ContributeProblems.RegionToOpen(link));

        // Any other word in the cell shows the row in neither region.
        Assert.Equal(string.Empty, ContributeProblems.RegionToOpen(new EntryProblem(ofOneRegion.Problem, ContributeEntry.ComponentPinout, "U7", string.Empty, "PLUTO")));
        Assert.Equal(BoardRegions.Pal, ContributeProblems.RegionToOpen(new EntryProblem(ofOneRegion.Problem, ContributeEntry.ComponentPinout, "U7", string.Empty, "pal")));
    }
}
