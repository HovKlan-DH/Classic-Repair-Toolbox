using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer.Tests;

// ###########################################################################################
// Covers ReviewNotInTable - what a submission changes that its table cannot show (2026-09-26).
//
// The table is the nine workbook sheets. Highlights and calibration points are published by an
// approval too, and the automatic checks' warnings are what the machine already found; once the
// change summary went, these few lines are the only place a maintainer sees them.
// ###########################################################################################
public sealed class ReviewNotInTableTests
{
    private static string Key(params string[] parts) => string.Join(BoardDraftNaturalKeys.Separator, parts);

    private static ReviewSectionView Section(
        string name,
        IReadOnlyList<string>? added = null,
        IReadOnlyList<string>? removed = null,
        IReadOnlyList<string>? changed = null) =>
        new(name, added ?? [], removed ?? [], changed ?? [], [], new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>());

    private static IReadOnlyList<string> Texts(IReadOnlyList<ReviewNoteLine> lines) => lines.Select(line => line.Text).ToList();

    // Every sheet the table shows is left out - it is already in front of the maintainer, coloured.
    [Fact]
    public void Changes_to_the_tables_own_sheets_are_not_repeated()
    {
        IReadOnlyList<ReviewSectionView> sections = BoardWorkbookSchema.AllSheets
            .Select(sheet => Section(sheet.SheetName, added: [Key("U8")]))
            .ToList();

        Assert.Empty(ReviewNotInTable.Lines(new ReviewChangeSummaryView(false, sections), []));
    }

    // The two sections outside the workbook, each named, with its rows readable - removals first,
    // as the summary always listed them.
    [Fact]
    public void Highlights_and_calibration_points_are_listed_with_their_rows()
    {
        var changes = new ReviewChangeSummaryView(false,
        [
            Section(ReviewSummary.SectionComponentHighlights,
                added: [Key("Schematic 1", "U10")],
                removed: [Key("Schematic 1", "U5")],
                changed: [Key("Schematic 2", "U8")]),
            Section(ReviewSummary.SectionKiCadCalibrations, changed: [Key("Top")])
        ]);

        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(changes, []);

        Assert.Equal(
            [
                "Component highlights: [1] removed (Schematic 1 / U5), [1] changed (Schematic 2 / U8), [1] added (Schematic 1 / U10)",
                "KiCad calibration points: [1] changed (Top)"
            ],
            Texts(lines));

        Assert.All(lines, line => Assert.Equal(ReviewNoteKind.Change, line.Kind));
    }

    // ###########################################################################################
    // *** A NEW SYSTEM SAYS HOW MUCH, NOT WHICH (owner request, 2026-09-26). *** Everything of it is
    // added, so listing the rows listed the whole board. Highlights count COMPONENTS - one
    // component highlighted on two schematics is one - and calibration points count schematics.
    // ###########################################################################################
    [Fact]
    public void A_new_system_counts_the_components_with_highlights_and_names_none()
    {
        var changes = new ReviewChangeSummaryView(true,
        [
            Section(BoardWorkbookSchema.SheetComponents, added: [Key("A1"), Key("A2")]),
            Section(ReviewSummary.SectionComponentHighlights,
                added: [Key("1N4148", "A1"), Key("1N4148", "A2"), Key("1N754A", "A1"), Key("1N754A", "a2")]),
            Section(ReviewSummary.SectionKiCadCalibrations, added: [Key("1N4148")])
        ]);

        Assert.Equal(
            [
                "Highlights included for [2] components",
                "KiCad calibration points included for [1] schematic"
            ],
            Texts(ReviewNotInTable.Lines(changes, [])));
    }

    // ###########################################################################################
    // *** THE NUMBER IS BOLD, THE BRACKETS AND WORDS ARE NOT (owner request, 2026-09-26: "[<b>2</b>]").
    // *** So the line comes in pieces, and only the count is marked - finished text could not say
    // which characters to embolden. A finding, which counts nothing, is one plain piece.
    // ###########################################################################################
    [Fact]
    public void Only_a_counts_number_is_marked_bold()
    {
        var changes = new ReviewChangeSummaryView(true,
            [Section(ReviewSummary.SectionComponentHighlights, added: [Key("1N4148", "A1"), Key("1N754A", "A2")])]);

        ReviewNoteLine line = Assert.Single(ReviewNotInTable.Lines(changes, []));

        Assert.Equal(
            [
                new ReviewNoteRun("Highlights included for [", IsCount: false),
                new ReviewNoteRun("2", IsCount: true),
                new ReviewNoteRun("] components", IsCount: false)
            ],
            line.Runs);

        ReviewNoteLine finding = Assert.Single(ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(false, []),
            [new ReviewFindingView("a", "", "A row names no file.", IsError: true)]));

        Assert.Equal([new ReviewNoteRun(finding.Text, IsCount: false)], finding.Runs);
    }

    // ###########################################################################################
    // *** NO "(not in the table)" (owner request, 2026-09-26). *** "Highlights for 2 components
    // (not in the table)" read as if the components were missing from the table, though they are on
    // its Components sheet.
    // ###########################################################################################
    [Fact]
    public void No_line_says_it_is_not_in_the_table()
    {
        var published = new ReviewChangeSummaryView(false,
            [Section(ReviewSummary.SectionComponentHighlights, added: [Key("Schematic 1", "U10")])]);
        var newSystem = new ReviewChangeSummaryView(true,
            [Section(ReviewSummary.SectionComponentHighlights, added: [Key("Schematic 1", "U10")])]);

        Assert.All(
            ReviewNotInTable.Lines(published, []).Concat(ReviewNotInTable.Lines(newSystem, [])),
            line => Assert.DoesNotContain("not in the table", line.Text, StringComparison.OrdinalIgnoreCase));
    }

    // A new system with no highlights or calibration points has nothing to say about them.
    [Fact]
    public void A_new_system_without_highlights_has_no_line_for_them()
    {
        var changes = new ReviewChangeSummaryView(true,
        [
            Section(BoardWorkbookSchema.SheetComponents, added: [Key("A1")]),
            Section(ReviewSummary.SectionComponentHighlights),
            Section(ReviewSummary.SectionKiCadCalibrations)
        ]);

        Assert.Empty(ReviewNotInTable.Lines(changes, []));
    }

    // A long list names a few rows and counts the rest, so one line stays one line.
    [Fact]
    public void A_long_list_of_rows_is_cut_short_and_counted()
    {
        List<string> keys = Enumerable.Range(1, ReviewNotInTable.MaximumNamedRows + 3)
            .Select(i => Key("Schematic 1", $"U{i}"))
            .ToList();

        var changes = new ReviewChangeSummaryView(false, [Section(ReviewSummary.SectionComponentHighlights, changed: keys)]);

        string line = Assert.Single(ReviewNotInTable.Lines(changes, [])).Text;

        Assert.EndsWith("Schematic 1 / U6 and 3 more)", line, StringComparison.Ordinal);
        Assert.DoesNotContain("U7", line, StringComparison.Ordinal);
    }

    [Fact]
    public void The_automatic_checks_findings_are_listed_as_warnings_and_errors()
    {
        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(false, []),
            [
                new ReviewFindingView("a", "U8", "The picture is very large.", IsError: false),
                new ReviewFindingView("b", "", "A row names no file.", IsError: true)
            ]);

        Assert.Equal(
            [
                new ReviewNoteLine("Warning from the automatic checks: The picture is very large. [U8]", ReviewNoteKind.Warning),
                new ReviewNoteLine("Error from the automatic checks: A row names no file.", ReviewNoteKind.Error)
            ],
            lines);
    }

    // The server could not compare it: said, rather than an empty space where changes belong.
    [Fact]
    public void A_submission_that_could_not_be_compared_says_so()
    {
        ReviewNoteLine line = Assert.Single(ReviewNotInTable.Lines(null, []));

        Assert.Equal(ReviewNoteKind.Error, line.Kind);
        Assert.Contains("could not be compared", line.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Nothing_to_say_is_no_lines_at_all()
    {
        Assert.Empty(ReviewNotInTable.Lines(new ReviewChangeSummaryView(false, []), []));
    }

    // ------------------------------------------------------------------ Files (code review, 2026-09-26)

    private static SubmittedFileFact File(string name, string hash = "aa", string? published = null) =>
        new($"Manu1/Hardware1/Board1/{name}", hash, 10, SubmissionFileScope.Own, IsReferenced: true, PublishedSha256: published);

    // ###########################################################################################
    // *** A FILE REPLACED UNDER ITS OWN NAME COLOURS NOTHING IN THE TABLE. *** The cell's text is
    // the path, which did not change - so without this line a replaced PDF or scan was approved
    // with nothing on screen ever saying a file changed (the blind spot the retired change
    // summary's file list used to close).
    // ###########################################################################################
    [Fact]
    public void A_file_replaced_under_its_own_name_is_said_although_no_row_changed()
    {
        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(false, []),
            [],
            [
                File("manual.pdf", "new", published: "old"),
                File("Sheet1.png", "aa", published: "aa"),
                File("extra.png", "bb", published: null)
            ]);

        Assert.Equal("Files: [3] included ([1] new, [1] replaced under the same name)", Assert.Single(lines).Text);
    }

    // Nothing the approval would write: said nowhere, since every file is published as it is.
    [Fact]
    public void A_submission_changing_no_file_gets_no_files_line()
    {
        Assert.Empty(ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(false, []),
            [],
            [File("Sheet1.png", "aa", published: "aa")]));
    }

    // A new system: everything is new, so the count alone says it.
    [Fact]
    public void A_new_systems_files_are_only_counted()
    {
        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(true, []),
            [],
            [File("Sheet1.png"), File("manual.pdf", "bb")]);

        Assert.Equal("Files included: [2] files", Assert.Single(lines).Text);
    }

    // KiCad data has its own line and is not counted twice.
    [Fact]
    public void KiCad_files_are_left_to_their_own_line()
    {
        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(false, []),
            [],
            [File("Sheet1.png", "new", published: "old"), KiCad("board.kicad_pcb", "cc", published: null)]);

        Assert.Equal(
            [
                "Files: [1] included ([1] replaced under the same name)",
                "KiCad data included: [1] file ([1] new)"
            ],
            Texts(lines));
    }

    // ------------------------------------------------------------------ KiCad data (2026-09-26)

    private static SubmittedFileFact KiCad(string name, string hash = "aa", string? published = null) =>
        new($"Manu1/Hardware1/Board1/KiCad data/{name}", hash, 10, SubmissionFileScope.Own, IsReferenced: false, PublishedSha256: published);

    // ###########################################################################################
    // *** THE SUBMISSION'S KiCad DATA IS COUNTED ABOVE THE TABLE (owner request, 2026-09-26). ***
    // KiCad files travel in the submission, cited by no row - no sheet shows them, so without this
    // line they would be approved and published unseen. A new system says only how many, since
    // everything of it is new; a published board says what moved against the tree.
    // ###########################################################################################
    [Fact]
    public void A_new_systems_KiCad_data_is_counted_without_new_or_changed()
    {
        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(true, []),
            [],
            [KiCad("board.kicad_pcb"), KiCad("board.kicad_sch", "bb"), KiCad("Pages/vic.kicad_sch", "cc")]);

        ReviewNoteLine line = Assert.Single(lines);
        Assert.Equal("KiCad data included: [3] files", line.Text);
        Assert.Equal(ReviewNoteKind.Change, line.Kind);
        Assert.Equal(["3"], line.Runs.Where(run => run.IsCount).Select(run => run.Text));
    }

    [Fact]
    public void A_published_boards_KiCad_data_says_what_is_new_and_what_changed()
    {
        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(false, []),
            [],
            [
                KiCad("board.kicad_pcb", "aa", published: "aa"),
                KiCad("board.kicad_sch", "bb", published: "old"),
                KiCad("Pages/vic.kicad_sch", "cc", published: null)
            ]);

        Assert.Equal("KiCad data included: [3] files ([1] new, [1] changed)", Assert.Single(lines).Text);
    }

    // The commonest case for a published board: the folder travels untouched, said in one word so
    // the maintainer does not go looking for a change that is not there.
    [Fact]
    public void KiCad_data_carried_unchanged_says_so()
    {
        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(
            new ReviewChangeSummaryView(false, []),
            [],
            [KiCad("board.kicad_pcb", "aa", published: "aa")]);

        Assert.Equal("KiCad data included: [1] file (unchanged)", Assert.Single(lines).Text);
    }

    // No KiCad files in the submission: no KiCad line. Other files are not KiCad data - the new
    // image below is counted by the FILES line instead, which is the only line it may produce.
    [Fact]
    public void A_submission_without_KiCad_data_gets_no_KiCad_line()
    {
        var image = new SubmittedFileFact(
            "Manu1/Hardware1/Board1/Sheet1.png", "aa", 10, SubmissionFileScope.Own, IsReferenced: true, PublishedSha256: null);

        Assert.Equal(
            ["Files: [1] included ([1] new)"],
            Texts(ReviewNotInTable.Lines(new ReviewChangeSummaryView(false, []), [], [image])));
    }
}
