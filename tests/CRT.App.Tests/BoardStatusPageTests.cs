using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardStatusPage - "Board status", a panel per part of the board (owner request, 2026-11-01: "each
// point in its own panel with a headline, then an explanation what this is about, a button for it and
// lastly also a status where it will count the data for it, if any"). What each says and how far
// each is, from the board and what is known of its KiCad data.
// ###########################################################################################
public sealed class BoardStatusPageTests
{
    private static List<BoardStatusPanel> Panels(
        BoardData board,
        IReadOnlyList<string>? files = null,
        BoardKiCadFacts? kiCad = null,
        IReadOnlyList<KiCadCalibrationEntry>? calibrations = null,
        IReadOnlyList<BoardDataProblem>? problems = null) =>
        BoardStatusPage.Panels(
            board,
            BoardReadiness.Missing(board, files is { Count: > 0 }, kiCad),
            calibrations,
            files,
            kiCad,
            problems);

    // ###########################################################################################
    // WHERE A COMPONENT IS PICKED, AND HOW ONE COMES WITH NO KiCad DATA (owner request, 2026-11-03: "it
    // should be mentioned ... that the user can use the normal category and component selector ... If this
    // board uses no KiCad, then ... add new components via the 'Add component' button and then afterwards
    // mark it for one or more schematic images").
    // ###########################################################################################
    [Fact]
    public void The_components_panel_says_where_to_pick_one_and_how_one_comes_without_KiCad_data()
    {
        string withKiCad = Part(Panels(new BoardData(), KiCadFiles(0, 1, 0)), BoardStatusPart.Components).Explanation;

        Assert.Equal(BoardStatusPage.ComponentsExplanation, withKiCad);
        Assert.Contains("\"Category\" and \"Component\" lists", withKiCad, StringComparison.Ordinal);

        string without = Part(Panels(new BoardData(), []), BoardStatusPart.Components).Explanation;

        Assert.StartsWith(BoardStatusPage.ComponentsExplanation, without, StringComparison.Ordinal);
        Assert.Contains("\"Add component\"", without, StringComparison.Ordinal);
        Assert.Contains("mark it on one or more schematic images", without, StringComparison.Ordinal);
    }

    private static BoardStatusPanel Part(IEnumerable<BoardStatusPanel> panels, BoardStatusPart part) => panels.Single(panel => panel.Part == part);

    private static string[] Lines(BoardStatusPanel panel) => panel.Lines.Select(line => new string(' ', line.Depth * 2) + line.Text).ToArray();

    // What a panel lists of what the checks found.
    private static string[] Listed(BoardStatusPanel panel) =>
        (panel.Problems ?? []).Select(item => $"{ContributeProblems.LevelWord(item.Problem.Level)} - {item.Text}").ToArray();

    private static string[] KiCadFiles(int projects, int pcbs, int sheets) =>
    [
        .. Enumerable.Range(1, projects).Select(number => $"KiCad data/board{number}.kicad_pro"),
        .. Enumerable.Range(1, pcbs).Select(number => $"KiCad data/board{number}.kicad_pcb"),
        .. Enumerable.Range(1, sheets).Select(number => $"KiCad data/sheet{number}.kicad_sch"),
    ];

    private static KiCadNetChoices Nets(int count) =>
        KiCadNetNames.Choices(Enumerable.Range(1, count).Select(number => ((string?)$"N{number}", (string?)$"N{number}", false)));

    // ###########################################################################################
    // The owner's own example: the page in the order asked for, the KiCad data's panel counting its
    // files by kind, what can be marked automatically, and its nets.
    // ###########################################################################################
    [Fact]
    public void The_panels_follow_the_owners_order_and_the_KiCad_panel_counts_what_it_holds()
    {
        var board = new BoardData
        {
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "One", KiCadNetName = "N1" }, new KiCadImportantSignalEntry { DisplayName = "One", KiCadNetName = "N2" }],
        };

        var facts = new BoardKiCadFacts(
            ["U1", "U2", "U3"],
            ["PCB top", "PCB bottom"],
            Markable: 3,
            Nets: Nets(451));

        List<BoardStatusPanel> panels = Panels(board, KiCadFiles(1, 1, 23), facts);

        Assert.Equal(
            [
                BoardStatusPart.Name, BoardStatusPart.Regions, BoardStatusPart.KiCad, BoardStatusPart.Signals, BoardStatusPart.Pictures,
                BoardStatusPart.Categories, BoardStatusPart.Files, BoardStatusPart.Links, BoardStatusPart.Credits, BoardStatusPart.Components,
            ],
            panels.Select(panel => panel.Part));

        Assert.Equal(
            [
                "25 KiCad files imported",
                "  1 project file",
                "  1 PCB file - 2 images, its top and its bottom",
                "  23 schematic files",
                "3 components can be marked automatically",
                "3 components are not on the board yet",
                "451 nets available",
                "  2 nets used in \"KiCad: Important signals\"",
            ],
            Lines(Part(panels, BoardStatusPart.KiCad)));

        Assert.Equal(BoardStatusState.Done, Part(panels, BoardStatusPart.KiCad).State);

        // A panel has a headline, what it is about and its button.
        Assert.All(panels, panel =>
        {
            Assert.False(string.IsNullOrWhiteSpace(panel.Title));
            Assert.False(string.IsNullOrWhiteSpace(panel.Explanation));
            Assert.False(string.IsNullOrWhiteSpace(panel.Action));
            Assert.DoesNotContain("...", panel.Action, StringComparison.Ordinal);
        });
    }

    // "Board name" can show no data: it is only about renaming (owner, 2026-11-01).
    [Fact]
    public void Board_name_shows_no_data()
    {
        BoardStatusPanel name = Part(Panels(new BoardData()), BoardStatusPart.Name);

        Assert.Empty(name.Lines);
        Assert.Equal(BoardStatusPage.NameTitle, name.Title);
        Assert.Equal(BoardStatusState.Done, name.State);
    }

    // ###########################################################################################
    // The KiCad project is optional: with none, its panel and the signals' are neither done nor
    // missing. While it is read no count is guessed; files CRT cannot read are an error.
    // ###########################################################################################
    [Fact]
    public void Without_KiCad_data_its_panels_are_optional_and_nothing_is_guessed_while_it_is_read()
    {
        var board = new BoardData();

        List<BoardStatusPanel> none = Panels(board, files: []);
        Assert.Equal(BoardStatusState.Optional, Part(none, BoardStatusPart.KiCad).State);
        Assert.Equal(BoardStatusState.Optional, Part(none, BoardStatusPart.Signals).State);

        // Files not counted yet.
        Assert.Equal([BoardStatusPage.ReadingKiCad], Lines(Part(Panels(board), BoardStatusPart.KiCad)));

        // Files counted, the data not read yet.
        List<BoardStatusPanel> reading = Panels(board, KiCadFiles(0, 1, 0));
        Assert.Equal(["1 KiCad file imported", "  1 PCB file - 2 images, its top and its bottom", BoardStatusPage.ReadingKiCad], Lines(Part(reading, BoardStatusPart.KiCad)));
        Assert.Equal([BoardStatusPage.ReadingKiCad], Lines(Part(reading, BoardStatusPart.Signals)));

        // Files CRT cannot read.
        List<BoardStatusPanel> unreadable = Panels(board, KiCadFiles(0, 1, 0), new BoardKiCadFacts([], [], Read: false));
        Assert.Equal(BoardStatusState.Error, Part(unreadable, BoardStatusPart.KiCad).State);
        Assert.Contains(BoardReadiness.KiCadUnreadable, Lines(Part(unreadable, BoardStatusPart.KiCad)));
    }

    // A PCB file is always read as two images, however many there are.
    [Fact]
    public void Several_PCB_files_say_two_images_each()
    {
        Assert.Contains(
            "  2 PCB files - 4 images, the top and the bottom of each",
            Lines(Part(Panels(new BoardData(), KiCadFiles(0, 2, 0)), BoardStatusPart.KiCad)));
    }

    // ###########################################################################################
    // Whether it has regions is either way an answer; nothing to do about it.
    [Fact]
    public void Regions_says_what_the_board_comes_in_and_is_never_something_to_do()
    {
        var with = new BoardData { Regions = ["PAL", "NTSC"] };

        Assert.Equal(BoardStatusState.Optional, Part(Panels(new BoardData()), BoardStatusPart.Regions).State);
        Assert.Equal(BoardStatusState.Done, Part(Panels(with), BoardStatusPart.Regions).State);
        Assert.Equal(["This board comes in PAL and NTSC versions"], Lines(Part(Panels(with), BoardStatusPart.Regions)));

        // The owner's words (2026-11-03) - PAL and NTSC in bold, TabDrafts showing them through BoldRuns.
        Assert.Equal(["No regions."], Lines(Part(Panels(new BoardData()), BoardStatusPart.Regions)));
        Assert.Equal(
            "Whether the board comes in a <b>PAL</b> and <b>NTSC</b> version, so component information can differ between them.",
            Part(Panels(with), BoardStatusPart.Regions).Explanation);
        Assert.Equal("Board: Regions", Part(Panels(with), BoardStatusPart.Regions).Title);
    }

    // A component with no category is said where it is put right, not among the components.
    [Fact]
    public void Categories_counts_them_and_says_the_error_of_a_component_with_none()
    {
        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }, new ComponentEntry { BoardLabel = "U2" }] };
        BoardDataProblem[] problems = [new BoardDataProblem(BoardProblemLevel.Error, "component.no_category", "U2", "[U2] has no category.", BoardWorkbookSchema.SheetComponents, 1, BoardWorkbookSchema.ColCategory)];

        List<BoardStatusPanel> panels = BoardStatusPage.Panels(board, BoardReadiness.Missing(board, hasKiCadFiles: false, null), null, [], null, problems);
        BoardStatusPanel categories = Part(panels, BoardStatusPart.Categories);

        Assert.Equal(BoardStatusState.Error, categories.State);
        Assert.Contains("1 category", Lines(categories));
        Assert.Contains("Error - [U2] has no category.", Listed(categories));
        Assert.DoesNotContain("Error - [U2] has no category.", Listed(Part(panels, BoardStatusPart.Components)));
    }

    // Schematic images, in the owner's words: the images KiCad knows, those added, the KiCad files mapped
    // to them (a side of a PCB counts as one, so "2 KiCad files mapped" here is one PCB file) and which are
    // aligned, and the components marked - done only when all of it is.
    // ###########################################################################################
    [Fact]
    public void Schematic_images_counts_its_way_to_done()
    {
        var board = new BoardData
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png", CadName = "PCB top" },
                new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "bottom.png", CadName = "PCB bottom" },
                new BoardSchematicEntry { SchematicName = "Photo", SchematicImageFile = "photo.png" },
                new BoardSchematicEntry { SchematicName = "Ports", SchematicImageFile = "ports.png" },
            ],
            Components = [new ComponentEntry { BoardLabel = "U1" }, new ComponentEntry { BoardLabel = "U2" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1" }],
        };

        var facts = new BoardKiCadFacts(["U1", "U2"], ["PCB top", "PCB bottom", "Sheet 1"]);
        KiCadCalibrationEntry[] calibrations = [new() { SchematicName = "Top", CadName = "PCB top" }];

        BoardStatusPanel pictures = Part(Panels(board, KiCadFiles(0, 1, 1), facts, calibrations), BoardStatusPart.Pictures);

        Assert.Equal(
            [
                "3 images known by KiCad data",
                "4 images added",
                "  2 KiCad files mapped",
                "    1 KiCad file aligned",
                "    1 KiCad file not aligned yet",
                "1 of 2 components marked",
            ],
            Lines(pictures));

        Assert.Equal(BoardStatusState.ToDo, pictures.State);
        Assert.Equal(BoardStatusState.ToDo, pictures.Lines.Single(line => line.Text == "1 KiCad file not aligned yet").Tone);

        // Everything aligned and marked: done, whatever pictures show no view (a photo of the ports).
        board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U2" });
        KiCadCalibrationEntry[] both = [.. calibrations, new() { SchematicName = "Bottom", CadName = "PCB bottom" }];

        Assert.Equal(BoardStatusState.Done, Part(Panels(board, KiCadFiles(0, 1, 1), facts, both), BoardStatusPart.Pictures).State);

        // None at all is missing.
        BoardStatusPanel none = Part(Panels(new BoardData(), files: []), BoardStatusPart.Pictures);
        Assert.Equal(BoardStatusState.ToDo, none.State);
        Assert.Contains(BoardReadiness.NoPicturesGap, Lines(none));
    }

    // ###########################################################################################
    // Important KiCad signals: the nets to pick from, the nets used, and - where names are shared - how
    // many names the Schematics tab's panel shows them as (the owner's "50 nets used in 45 unique"). A
    // signal naming a net the data does not have is said.
    // ###########################################################################################
    [Fact]
    public void Important_signals_counts_nets_and_names()
    {
        var board = new BoardData
        {
            KiCadImportantSignals =
            [
                new KiCadImportantSignalEntry { DisplayName = "9VAC", KiCadNetName = "N1" },
                new KiCadImportantSignalEntry { DisplayName = "9VAC", KiCadNetName = "N2" },
                new KiCadImportantSignalEntry { DisplayName = "Reset", KiCadNetName = "N3" },
                new KiCadImportantSignalEntry { DisplayName = "Gone", KiCadNetName = "NOPE" },
            ],
        };

        var facts = new BoardKiCadFacts([], [], Nets: Nets(5));
        BoardStatusPanel signals = Part(Panels(board, KiCadFiles(0, 1, 0), facts), BoardStatusPart.Signals);

        Assert.Equal(
            [
                "5 nets available",
                "3 nets used, shown as 2 names on the \"Schematics\" tab",
                "  1 names a net the KiCad data does not have - it will not light up",
            ],
            Lines(signals));

        Assert.Equal(BoardStatusState.Done, signals.State);

        // None yet, with nets to pick from: missing.
        BoardStatusPanel none = Part(Panels(new BoardData(), KiCadFiles(0, 1, 0), facts), BoardStatusPart.Signals);
        Assert.Equal(BoardStatusState.ToDo, none.State);
        Assert.Contains(BoardReadiness.NoSignalsGap, Lines(none));
    }

    // ###########################################################################################
    // Components (the owner left it to me, 2026-11-01): how many, how much more they hold, and what the
    // KiCad data has that the board does not. What each lacks is the tab's groups, not counted here.
    // ###########################################################################################
    [Fact]
    public void Components_counts_what_they_hold_and_what_KiCad_has_more()
    {
        var board = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1" }, new ComponentEntry { BoardLabel = "U2" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U1", Pin = "1" }, new ComponentPinEntry { BoardLabel = "U1", Pin = "2" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U1", Pin = "1", TimeDiv = "1mS" }],
            ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U1", PartNumber = "x" }],
        };

        BoardStatusPanel components = Part(Panels(board, KiCadFiles(0, 1, 0), new BoardKiCadFacts(["U1", "U2", "U3"], [])), BoardStatusPart.Components);

        Assert.Equal(
            [
                "2 components",
                "  1 with its pins listed",
                "  1 oscilloscope baseline",
                "  1 replacement part",
                "1 component in the KiCad data is not on the board yet - \"KiCad: Files\" adds it",
            ],
            Lines(components));

        Assert.Equal(BoardStatusState.ToDo, components.State);

        Assert.Equal(BoardStatusState.Done, Part(Panels(board, KiCadFiles(0, 1, 0), new BoardKiCadFacts(["U1", "U2"], [])), BoardStatusPart.Components).State);
        Assert.Equal(BoardStatusState.ToDo, Part(Panels(new BoardData()), BoardStatusPart.Components).State);
    }

    // Files and Links: how many, in how many categories. None is no fault - neither is needed.
    [Fact]
    public void Files_and_links_count_their_rows_and_categories()
    {
        var board = new BoardData
        {
            BoardLocalFiles =
            [
                new BoardLocalFileEntry { Category = "Resources", Name = "A", File = "a.pdf" },
                new BoardLocalFileEntry { Category = "resources ", Name = "B", File = "b.pdf" },
                new BoardLocalFileEntry { Category = "Troubleshooting", Name = "C", File = "c.pdf" },
            ],
            BoardLinks = [new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = "https://www.lemon64.com" }],
        };

        List<BoardStatusPanel> panels = Panels(board);

        Assert.Equal(["3 board files defined", "2 categories defined"], Lines(Part(panels, BoardStatusPart.Files)));
        Assert.Equal(["1 board link defined", "1 category defined"], Lines(Part(panels, BoardStatusPart.Links)));
        Assert.Equal(BoardStatusState.Done, Part(panels, BoardStatusPart.Files).State);

        BoardStatusPanel none = Part(Panels(new BoardData()), BoardStatusPart.Links);
        Assert.Equal(BoardStatusState.Optional, none.State);
        Assert.Equal(["No board links yet."], Lines(none));
    }

    // ###########################################################################################
    // *** EVERY ERROR AND WARNING IS SAID ON THE PANEL IT IS ABOUT (owner request, 2026-11-01: "it should
    // have the errors and warnings directly per section"). *** By the sheet it is on; a highlight, on no
    // sheet, is the images'. The table is a button away, on that sheet. Only what is about no part keeps
    // a panel of its own, first - and only while there is any.
    // ###########################################################################################
    [Fact]
    public void Every_error_and_warning_is_said_on_the_panel_it_is_about()
    {
        BoardDataProblem[] problems =
        [
            Problem(BoardProblemLevel.Warning, "image.time_div", "A setting is odd.", BoardWorkbookSchema.SheetComponentImages),
            Problem(BoardProblemLevel.Error, "file.missing", "A file is gone.", BoardWorkbookSchema.SheetComponentLocalFiles),
            Problem(BoardProblemLevel.Warning, "schematic.opacity", "An opacity is odd.", BoardWorkbookSchema.SheetBoardSchematics),
            Problem(BoardProblemLevel.Warning, "highlight.orphan", "A mark has no component.", null),
            Problem(BoardProblemLevel.Warning, "row.duplicate", "A link twice.", BoardWorkbookSchema.SheetBoardLinks),
            Problem(BoardProblemLevel.Error, "row.duplicate", "A file twice.", BoardWorkbookSchema.SheetBoardLocalFiles),
            Problem(BoardProblemLevel.Warning, "row.duplicate", "A credit twice.", BoardWorkbookSchema.SheetCredits),
            Problem(BoardProblemLevel.Warning, "row.duplicate", "A signal twice.", BoardWorkbookSchema.SheetKiCadImportantSignals),
        ];

        List<BoardStatusPanel> panels = Panels(new BoardData(), problems: problems);

        Assert.DoesNotContain(panels, panel => panel.Part == BoardStatusPart.Problems);

        // Errors first, and an error makes the panel one.
        BoardStatusPanel components = Part(panels, BoardStatusPart.Components);
        Assert.Equal(BoardStatusState.Error, components.State);
        Assert.Equal(["No components yet."], Lines(components));
        Assert.Equal(["Error - A file is gone.", "Warning - A setting is odd."], Listed(components));
        Assert.Equal(new BoardProblemCounts(1, 1), components.ProblemCounts);

        // Every line there has the way to its own component, so no table (owner request, 2026-11-04).
        Assert.Null(components.TableSheet);

        // A highlight is on no sheet, so the table cannot show it - no button for that one alone.
        BoardStatusPanel pictures = Part(panels, BoardStatusPart.Pictures);
        Assert.Contains("Warning - A mark has no component.", Listed(pictures));
        Assert.Contains("Warning - An opacity is odd.", Listed(pictures));
        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, pictures.TableSheet);

        Assert.Contains("Error - A file twice.", Listed(Part(panels, BoardStatusPart.Files)));
        Assert.Equal(BoardStatusState.Error, Part(panels, BoardStatusPart.Files).State);

        // A warning leaves nothing optional: there is something to look at.
        Assert.Contains("Warning - A link twice.", Listed(Part(panels, BoardStatusPart.Links)));
        Assert.Equal(BoardStatusState.ToDo, Part(panels, BoardStatusPart.Links).State);
        Assert.Contains("Warning - A credit twice.", Listed(Part(panels, BoardStatusPart.Credits)));
        Assert.Contains("Warning - A signal twice.", Listed(Part(panels, BoardStatusPart.Signals)));

        Assert.Null(Part(panels, BoardStatusPart.Name).TableSheet);
    }

    // What is about no part keeps a panel of its own, first on the page - there only while there is any.
    [Fact]
    public void What_is_about_no_part_has_a_panel_of_its_own_first()
    {
        List<BoardStatusPanel> panels = Panels(new BoardData(), problems: [Problem(BoardProblemLevel.Warning, "something.else", "Something else.", null)]);

        Assert.Equal(BoardStatusPart.Problems, panels[0].Part);
        Assert.Equal(BoardStatusState.ToDo, panels[0].State);
        Assert.Equal(["Warning - Something else."], Listed(panels[0]));
        Assert.Equal(BoardReadiness.TableAction, panels[0].Action);

        Assert.Equal(BoardStatusState.Error, Panels(new BoardData(), problems: [Problem(BoardProblemLevel.Error, "something.else", "Something else.", null)])[0].State);
        Assert.DoesNotContain(Panels(new BoardData()), panel => panel.Part == BoardStatusPart.Problems);
    }

    // ###########################################################################################
    // A panel names the first eight in full and counts the rest - the owner's list of nineteen sentences
    // was the complaint - and NOTHING IS LEFT OUT FOR BEING SAID ANOTHER WAY (owner request, 2026-11-04: "it
    // is not at all clear for me where those errors are, except I know they are in 'Credits' and
    // 'Components'"). A row missing a field and a component marked nowhere were left to its list's line and
    // to a group of buttons, so the total on the page added up to nothing on it.
    // ###########################################################################################
    [Fact]
    public void A_panel_lists_eight_and_counts_the_rest_and_leaves_none_out()
    {
        BoardDataProblem[] many =
        [
            .. Enumerable.Range(1, 10).Select(number => Problem(BoardProblemLevel.Warning, "image.time_div", $"Setting {number}.", BoardWorkbookSchema.SheetComponentImages)),
        ];

        BoardStatusPanel crowded = Part(Panels(new BoardData(), problems: many), BoardStatusPart.Components);

        Assert.Contains("Warning - Setting 8.", Listed(crowded));
        Assert.DoesNotContain("Warning - Setting 9.", Listed(crowded));
        Assert.Equal("and 2 more", crowded.ProblemsMore);
        Assert.Equal(new BoardProblemCounts(0, 10), crowded.ProblemCounts);

        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "U1" }] };
        BoardDataProblem unmarked = Problem(BoardProblemLevel.Warning, "component.no_highlight", "Component [U1] is not marked on any schematic.", BoardWorkbookSchema.SheetComponents);
        BoardStatusPanel components = Part(Panels(board, problems: [unmarked]), BoardStatusPart.Components);

        // Counted, and said by the one line and the buttons the tab puts in the frame: the same four lines
        // for each of eleven components say nothing that one does not.
        Assert.Equal(["1 component"], Lines(components));
        Assert.Empty(Listed(components));
        Assert.Equal(new BoardProblemCounts(0, 1), components.ProblemCounts);
        Assert.Equal(BoardStatusState.ToDo, components.State);
        Assert.Null(components.TableSheet);

        var credit = new BoardData { Credits = [new CreditEntry { Category = "Schematic image", SubCategory = "Top" }] };
        BoardDataProblem required = Problem(BoardProblemLevel.Error, "row.required_empty", "This credit has no \"Name or handle\".", BoardWorkbookSchema.SheetCredits);
        BoardStatusPanel credits = Part(Panels(credit, problems: [required]), BoardStatusPart.Credits);

        // The row is named: one of three credits saying the same would say nothing. And the line in red that
        // counted such rows goes (owner request, 2026-11-04: "it already shows the error panel").
        Assert.Equal(["1 credit defined"], Lines(credits));
        Assert.Equal(["Error - [Schematic image / Top] This credit has no \"Name or handle\"."], Listed(credits));
        Assert.Equal(BoardStatusState.Error, credits.State);
    }

    // A board file with no file is one of the fields its row needs: the check says which, as a component's
    // file, and the line in red that counted such rows has nothing to add.
    [Fact]
    public void A_board_file_with_no_file_is_named_by_the_check_alone()
    {
        var board = new BoardData { BoardLocalFiles = [new BoardLocalFileEntry { Category = "Resources", Name = "Manual" }] };
        BoardDataProblem noFile = Problem(BoardProblemLevel.Error, "file.unreferenced", "The board file does not name a file.", BoardWorkbookSchema.SheetBoardLocalFiles);

        BoardStatusPanel files = Part(Panels(board, problems: [noFile]), BoardStatusPart.Files);

        Assert.DoesNotContain(BoardReadiness.IncompleteGap(1, "file", "files"), Lines(files));
        Assert.Equal(["Error - [Resources / Manual] The board file does not name a file."], Listed(files));
        Assert.Equal(BoardStatusState.Error, files.State);

        BoardDataProblem ofComponent = Problem(BoardProblemLevel.Error, "file.unreferenced", "The component file does not name a file.", BoardWorkbookSchema.SheetComponentLocalFiles);

        Assert.Contains("Error - The component file does not name a file.", Listed(Part(Panels(new BoardData(), problems: [ofComponent]), BoardStatusPart.Components)));
    }

    // Said on the panel it is about, as every problem is.
    [Fact]
    public void A_region_that_is_none_is_said_on_the_regions_panel()
    {
        BoardDataProblem unknown = Problem(BoardProblemLevel.Error, "regions.unknown", "The board says it comes in [x].", null);

        BoardStatusPanel regions = Part(Panels(new BoardData(), problems: [unknown]), BoardStatusPart.Regions);

        Assert.Contains("Error - The board says it comes in [x].", Listed(regions));
        Assert.Equal(BoardStatusState.Error, regions.State);
    }

    // A row missing a field is an error on its list's panel even before the checks have answered.
    [Fact]
    public void A_row_missing_a_field_is_an_error_on_its_lists_panel()
    {
        BoardStatusPanel credits = Part(Panels(new BoardData { Credits = [new CreditEntry { Category = "Schematic image" }] }), BoardStatusPart.Credits);

        Assert.Equal(BoardStatusState.Error, credits.State);
        Assert.Contains(BoardReadiness.IncompleteGap(1, "credit", "credits"), Lines(credits));

        // Nobody credited is missing, not an error.
        BoardStatusPanel none = Part(Panels(new BoardData()), BoardStatusPart.Credits);
        Assert.Equal(BoardStatusState.ToDo, none.State);
        Assert.Contains(BoardReadiness.NoCreditsGap, Lines(none));
    }

    // A tab, a menu or a button named on the page is in quotes (owner request, 2026-11-01).
    [Fact]
    public void A_part_named_on_the_page_is_in_quotes()
    {
        var facts = new BoardKiCadFacts(["U1"], [], Nets: Nets(2));
        string[] kiCad = Lines(Part(Panels(new BoardData(), KiCadFiles(0, 1, 0), facts), BoardStatusPart.KiCad));

        Assert.Contains("  0 nets used in \"KiCad: Important signals\"", kiCad);
    }

    private static BoardDataProblem Problem(BoardProblemLevel level, string code, string message, string? sheet) =>
        new(level, code, "subject", message, sheet, sheet is null ? -1 : 0, null);
}
