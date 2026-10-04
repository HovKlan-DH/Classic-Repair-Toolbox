using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

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

    // ###########################################################################################
    // *** THE OWNER'S OWN EXAMPLE (owner request, 2026-10-04) *** - the "hest" component taken off
    // the "Board layout" schematic read "Component highlights: [1] removed (Board layout / hest)".
    // Now a counted heading, and the change on a line of its own, set in under it:
    //
    //     Component highlights have [1] change:
    //       Removed component [hest] from schematic "Board layout"
    // ###########################################################################################
    [Fact]
    public void A_removed_highlight_reads_as_the_owner_wrote_it()
    {
        var changes = new ReviewChangeSummaryView(false,
            [Section(ReviewSummary.SectionComponentHighlights, removed: [Key("Board layout", "hest")])]);

        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(changes, []);

        Assert.Equal(
            [
                "Component highlights have [1] change:",
                "Removed component [hest] from schematic \"Board layout\""
            ],
            Texts(lines));

        Assert.Equal([0, 1], lines.Select(line => line.Indent));
        Assert.All(lines, line => Assert.Equal(ReviewNoteKind.Change, line.Kind));

        // The heading's count alone is bold.
        Assert.Equal(
            [
                new ReviewNoteRun("Component highlights have [", IsCount: false),
                new ReviewNoteRun("1", IsCount: true),
                new ReviewNoteRun("] change:", IsCount: false)
            ],
            lines[0].Runs);
    }

    // The two sections outside the workbook, each under its own heading, each change in words -
    // removals first, as the summary always listed them. A highlight still there says what happened
    // to its rectangle; the fields say which.
    [Fact]
    public void Highlights_and_calibration_points_are_listed_one_change_a_line_under_their_heading()
    {
        var changes = new ReviewChangeSummaryView(false,
        [
            new ReviewSectionView(
                ReviewSummary.SectionComponentHighlights,
                Added: [Key("Schematic 1", "U10")],
                Removed: [Key("Schematic 1", "U5")],
                Changed: [Key("Schematic 2", "U8"), Key("Schematic 2", "U9"), Key("Schematic 2", "U11"), Key("Schematic 2", "U12")],
                Renamed: [new ReviewRenameView(Key("Schematic 1", "U1"), Key("Schematic 1", "U2"), AlsoChanged: false)],
                FieldChanges: new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>
                {
                    [Key("Schematic 2", "U8")] = [new("X", "1", "2")],
                    [Key("Schematic 2", "U9")] = [new("Height", "1", "2")],
                    [Key("Schematic 2", "U11")] = [new("Y", "1", "2"), new("Width", "1", "2")]
                }),
            new ReviewSectionView(
                ReviewSummary.SectionKiCadCalibrations,
                Added: [Key("Bottom")],
                Removed: [],
                Changed: [Key("Top")],
                Renamed: [new ReviewRenameView(Key("Old"), Key("New"), AlsoChanged: false)],
                FieldChanges: new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>())
        ]);

        IReadOnlyList<ReviewNoteLine> lines = ReviewNotInTable.Lines(changes, []);

        Assert.Equal(
            [
                "Component highlights have [7] changes:",
                "Removed component [U5] from schematic \"Schematic 1\"",
                "Moved component [U8] on schematic \"Schematic 2\"",
                "Resized component [U9] on schematic \"Schematic 2\"",
                "Moved and resized component [U11] on schematic \"Schematic 2\"",
                "Changed component [U12] on schematic \"Schematic 2\"",
                "Added component [U10] to schematic \"Schematic 1\"",
                "and [1] more",
                "KiCad calibration points have [3] changes:",
                "Changed the calibration points of schematic \"Top\"",
                "Added calibration points to schematic \"Bottom\"",
                "Moved the calibration points of schematic \"Old\" to schematic \"New\""
            ],
            Texts(lines));

        Assert.Equal([0, 1, 1, 1, 1, 1, 1, 1, 0, 1, 1, 1], lines.Select(line => line.Indent));
    }

    // A highlight whose label changed, whose schematic was renamed, or both - one row each
    // (ReviewSummary pairs them, 2026-10-04).
    [Fact]
    public void A_renamed_highlight_says_what_changed_about_it()
    {
        var changes = new ReviewChangeSummaryView(false,
        [
            new ReviewSectionView(
                ReviewSummary.SectionComponentHighlights,
                Added: [],
                Removed: [],
                Changed: [],
                Renamed:
                [
                    new ReviewRenameView(Key("Board layout", "hest"), Key("Board layout", "U5"), AlsoChanged: false),
                    new ReviewRenameView(Key("Board", "U8"), Key("Board layout", "U8"), AlsoChanged: false),
                    new ReviewRenameView(Key("Board", "U9"), Key("Board layout", "U10"), AlsoChanged: true)
                ],
                FieldChanges: new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>())
        ]);

        Assert.Equal(
            [
                "Component highlights have [3] changes:",
                "Renamed component [hest] to [U5] on schematic \"Board layout\"",
                "Moved component [U8] from schematic \"Board\" to schematic \"Board layout\"",
                "Renamed component [U9] on schematic \"Board\" to [U10] on schematic \"Board layout\", and changed it"
            ],
            Texts(ReviewNotInTable.Lines(changes, [])));
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

    // A long list names a few changes and counts the rest, so a board relabelled wholesale does not
    // push the table off the screen. The heading still counts them all.
    [Fact]
    public void A_long_list_of_rows_is_cut_short_and_counted()
    {
        List<string> keys = Enumerable.Range(1, ReviewNotInTable.MaximumNamedRows + 3)
            .Select(i => Key("Schematic 1", $"U{i}"))
            .ToList();

        var changes = new ReviewChangeSummaryView(false, [Section(ReviewSummary.SectionComponentHighlights, changed: keys)]);

        IReadOnlyList<string> lines = Texts(ReviewNotInTable.Lines(changes, []));

        Assert.Equal($"Component highlights have [{ReviewNotInTable.MaximumNamedRows + 3}] changes:", lines[0]);
        Assert.Equal(ReviewNotInTable.MaximumNamedRows + 2, lines.Count);
        Assert.Equal("Changed component [U6] on schematic \"Schematic 1\"", lines[^2]);
        Assert.Equal("and [3] more", lines[^1]);
        Assert.DoesNotContain(lines, line => line.Contains("U7", StringComparison.Ordinal));
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

    // ###########################################################################################
    // *** NO FILE LINES HERE ANY MORE (owner request, 2026-09-30). *** "Files: [3] included ([1]
    // replaced under the same name)" and "KiCad data included: ..." were said here, for files no
    // cell shows as changed. The submission's Files view shows them, and the count on its button
    // is what still says so before anything is opened - its tests, SubmissionViewsTests, carry
    // every case these lines' tests did (a file replaced under its own name, a new one, KiCad data,
    // nothing changed). Lines() no longer even takes the files, so nothing here can bring them back.
    // ###########################################################################################
}
