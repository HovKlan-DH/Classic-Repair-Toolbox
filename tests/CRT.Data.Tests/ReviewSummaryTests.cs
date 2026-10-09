using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Covers ReviewSummary - what the Maintainer tab opens on (Phase 5, task 3).
//
// WHAT MAKES THIS WORTH TESTING HARD: the summary is the maintainer's whole view of a submission
// before they drill in. A summary that misses a change means a maintainer approves something they
// never saw, and publishing is irreversible - the project owner ruled out retained revisions, so
// there is no undo behind a wrong approval.
//
// The failure mode to beware of is the opposite one too: a summary that reports changes nobody
// made (by pairing rows on position, or by comparing a retired UUID column) is noise, and a
// maintainer who learns to scroll past noise stops reading the real entries.
public sealed class ReviewSummaryTests
{
    private static ComponentEntry Component(string label, string friendly = "PLA", string part = "906114") => new()
    {
        BoardLabel = label,
        FriendlyName = friendly,
        TechnicalNameOrValue = "906114-01",
        PartNumber = part
    };

    private static BoardData Board(params ComponentEntry[] components) => new()
    {
        RevisionDate = "2026-August-21",
        Components = [.. components]
    };

    private static ReviewSectionChange Section(ReviewChangeSummary summary, string name) =>
        summary.Sections.First(section => section.Section == name);

    private static ReviewSectionChange Components(ReviewChangeSummary summary) =>
        ReviewSummaryTests.Section(summary, BoardWorkbookSchema.SheetComponents);

    // -----------------------------------------------------------------------------------
    // The basic verbs
    // -----------------------------------------------------------------------------------

    [Fact]
    public void An_identical_board_reports_no_changes()
    {
        // The commonest thing a maintainer should never be shown: a resubmission of what is already
        // published. If this reports changes, every real change is buried in noise.
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));

        ReviewChangeSummary summary = ReviewSummary.Compare(published, submitted);

        Assert.False(summary.HasChanges);
        Assert.Equal(0, summary.TotalChanges);
        Assert.Equal("No changes", summary.Describe());
    }

    [Fact]
    public void An_added_row_is_reported_as_added()
    {
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U9"));

        ReviewSectionChange components = ReviewSummaryTests.Components(ReviewSummary.Compare(published, submitted));

        Assert.Equal(["U9"], components.Added);
        Assert.Empty(components.Changed);
        Assert.Empty(components.Removed);
    }

    [Fact]
    public void A_removed_row_is_reported_as_removed()
    {
        BoardData published = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U9"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));

        ReviewSectionChange components = ReviewSummaryTests.Components(ReviewSummary.Compare(published, submitted));

        Assert.Equal(["U9"], components.Removed);
        Assert.Empty(components.Added);
    }

    [Fact]
    public void An_edited_field_is_reported_as_changed()
    {
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8", part: "906114"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8", part: "251715-01"));

        ReviewSectionChange components = ReviewSummaryTests.Components(ReviewSummary.Compare(published, submitted));

        Assert.Equal(["U8"], components.Changed);
        Assert.Empty(components.Added);
        Assert.Empty(components.Removed);
    }

    // -----------------------------------------------------------------------------------
    // Pairing on natural keys, not position
    // -----------------------------------------------------------------------------------

    [Fact]
    public void Inserting_a_row_at_the_TOP_does_not_report_every_later_row_as_changed()
    {
        // *** THE CLASSIC DIFF BUG. *** Pairing rows by list position would report U8 and U9 as
        // changed simply because a row was inserted above them. On a board with hundreds of
        // components that is a summary a maintainer cannot use at all - and it would be worst on
        // exactly the submissions that add something, which is most of them.
        BoardData published = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U9"));

        BoardData submitted = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U1"),
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U9"));

        ReviewSectionChange components = ReviewSummaryTests.Components(ReviewSummary.Compare(published, submitted));

        Assert.Equal(["U1"], components.Added);
        Assert.Empty(components.Changed);
        Assert.Empty(components.Removed);
    }

    [Fact]
    public void Reordering_rows_alone_reports_nothing()
    {
        BoardData published = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U9"));

        BoardData submitted = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U9"),
            ReviewSummaryTests.Component("U8"));

        Assert.False(ReviewSummary.Compare(published, submitted).HasChanges);
    }

    [Fact]
    public void A_retired_UUID_is_ignored_when_comparing()
    {
        // Phase 4 retired UuidV4 as an identity and stopped writing new ones. A published row
        // carrying one and a resubmitted row without it are the SAME row; comparing the column
        // would report a change nobody made, on every row of every pre-transition board.
        BoardData published = new()
        {
            RevisionDate = "2026-August-21",
            Components = [new ComponentEntry
            {
                BoardLabel = "U8",
                FriendlyName = "PLA",
                TechnicalNameOrValue = "906114-01",
                PartNumber = "906114"
            }]
        };

        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));

        Assert.False(ReviewSummary.Compare(published, submitted).HasChanges);
    }

    [Fact]
    public void A_case_change_IS_reported()
    {
        // Ordinal and case-sensitive, like every other comparison in this project. "u8" to "U8"
        // is a real edit, and on the Linux server a file name's case genuinely matters.
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8", friendly: "pla"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8", friendly: "PLA"));

        Assert.Equal(["U8"], ReviewSummaryTests.Components(ReviewSummary.Compare(published, submitted)).Changed);
    }

    // -----------------------------------------------------------------------------------
    // Renames
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_declared_rename_is_reported_as_a_rename_and_not_as_a_delete_plus_an_add()
    {
        // Natural keys cannot see a rename - U8 to U9 reads as a delete plus an add, which is
        // true and tells the maintainer nothing. The client knows it was a rename and says so.
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U9"));

        ReviewChangeSummary summary = ReviewSummary.Compare(
            published,
            submitted,
            [new SubmissionRename { Section = BoardWorkbookSchema.SheetComponents, From = "U8", To = "U9" }]);

        ReviewSectionChange components = ReviewSummaryTests.Components(summary);

        ReviewRenamedRow renamed = Assert.Single(components.Renamed);
        Assert.Equal("U8", renamed.From);
        Assert.Equal("U9", renamed.To);
        Assert.False(renamed.AlsoChanged);

        // Never double-counted.
        Assert.Empty(components.Added);
        Assert.Empty(components.Removed);
        Assert.Equal(1, components.TotalChanges);
    }

    [Fact]
    public void A_rename_that_also_edits_a_field_says_so()
    {
        // Worth distinguishing: a maintainer who sees "renamed" assumes the rest is untouched, and
        // would not look further.
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8", part: "906114"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U9", part: "251715-01"));

        ReviewChangeSummary summary = ReviewSummary.Compare(
            published,
            submitted,
            [new SubmissionRename { Section = BoardWorkbookSchema.SheetComponents, From = "U8", To = "U9" }]);

        Assert.True(Assert.Single(ReviewSummaryTests.Components(summary).Renamed).AlsoChanged);
    }

    [Fact]
    public void A_rename_that_does_not_describe_what_actually_happened_is_ignored()
    {
        // *** A DECLARED RENAME IS UNTRUSTED INPUT. *** It arrives in the submission. A client
        // declaring a rename it did not make would otherwise hide a genuine addition or deletion
        // from the maintainer, which is the one thing this screen exists to prevent.
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U9"));

        ReviewChangeSummary summary = ReviewSummary.Compare(
            published,
            submitted,
            // Claims U8 became U9, but U8 is still there - so U9 is an ADD, not a rename.
            [new SubmissionRename { Section = BoardWorkbookSchema.SheetComponents, From = "U8", To = "U9" }]);

        ReviewSectionChange components = ReviewSummaryTests.Components(summary);

        Assert.Empty(components.Renamed);
        Assert.Equal(["U9"], components.Added);
    }

    [Fact]
    public void A_rename_declared_for_another_section_does_not_move_a_row_here()
    {
        // U9 is a different component, not U8 relabelled - or the summary would pair the two by
        // their content (the next tests), whatever was declared.
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U9", friendly: "VIC"));

        ReviewChangeSummary summary = ReviewSummary.Compare(
            published,
            submitted,
            [new SubmissionRename { Section = BoardWorkbookSchema.SheetCredits, From = "U8", To = "U9" }]);

        ReviewSectionChange components = ReviewSummaryTests.Components(summary);

        Assert.Empty(components.Renamed);
        Assert.Equal(["U9"], components.Added);
        Assert.Equal(["U8"], components.Removed);
    }

    // ###########################################################################################
    // *** A RENAME NOBODY DECLARED IS STILL ONE ROW (owner decision, 2026-10-04). *** No client
    // declares renames, so U8 relabelled U9 - nothing else of it changed - was one removed and one
    // added here while the contributor's table showed ONE row changed. The summary pairs them by the
    // table's own rule (BoardDataDiffer.PairRenamedRows).
    // ###########################################################################################
    [Fact]
    public void A_row_whose_key_alone_changed_is_reported_as_renamed_without_being_declared()
    {
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U9"));

        ReviewSectionChange components = ReviewSummaryTests.Components(ReviewSummary.Compare(published, submitted));

        ReviewRenamedRow renamed = Assert.Single(components.Renamed);
        Assert.Equal(("U8", "U9", false), (renamed.From, renamed.To, renamed.AlsoChanged));
        Assert.Empty(components.Added);
        Assert.Empty(components.Removed);
        Assert.Equal(1, components.TotalChanges);
    }

    [Fact]
    public void A_row_whose_key_AND_another_field_changed_is_still_a_removal_plus_an_addition()
    {
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U9", part: "251715-01"));

        ReviewSectionChange components = ReviewSummaryTests.Components(ReviewSummary.Compare(published, submitted));

        Assert.Empty(components.Renamed);
        Assert.Equal(["U9"], components.Added);
        Assert.Equal(["U8"], components.Removed);
    }

    // A highlight relabelled on its schematic - the same rectangle - is the case the maintainer's
    // lines above the table name ("Renamed component [U8] to [U9] on schematic ...").
    [Fact]
    public void A_highlight_relabelled_on_the_same_rectangle_is_renamed()
    {
        var published = new BoardData { ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "S", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }] };
        var submitted = new BoardData { ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "S", BoardLabel = "U9", X = "1", Y = "2", Width = "3", Height = "4" }] };

        ReviewSectionChange highlights = ReviewSummary.Compare(published, submitted).Sections
            .Single(section => section.Section == ReviewSummary.SectionComponentHighlights);

        ReviewRenamedRow renamed = Assert.Single(highlights.Renamed);
        Assert.Equal(BoardDraftNaturalKeys.ForComponentHighlight("S", "U8"), renamed.From);
        Assert.Equal(BoardDraftNaturalKeys.ForComponentHighlight("S", "U9"), renamed.To);
        Assert.Empty(highlights.Added);
        Assert.Empty(highlights.Removed);
    }

    // -----------------------------------------------------------------------------------
    // A new board
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_new_board_is_summarised_as_such_rather_than_row_by_row()
    {
        // A new board is the highest-risk submission there is (Phase 6 task 3). Listing every
        // row as an addition would be true and useless - hundreds of items long.
        BoardData submitted = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U9"));

        ReviewChangeSummary summary = ReviewSummary.Compare(published: null, submitted);

        Assert.True(summary.IsNewBoard);
        Assert.Contains("New board", summary.Describe());
        Assert.Contains("2 rows", summary.Describe());
    }

    [Fact]
    public void A_new_board_with_one_row_reads_as_one_row_not_one_rows()
    {
        BoardData submitted = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));

        Assert.Contains("1 row", ReviewSummary.Compare(published: null, submitted).Describe());
        Assert.DoesNotContain("1 rows", ReviewSummary.Compare(published: null, submitted).Describe());
    }

    // -----------------------------------------------------------------------------------
    // Every section is covered
    // -----------------------------------------------------------------------------------

    [Fact]
    public void Every_section_of_a_board_is_compared()
    {
        // A section missing from AllSections is a section whose changes are INVISIBLE to the
        // maintainer - approved without ever being shown. This asserts each one is present and can
        // actually detect a change, rather than merely being listed.
        var published = new BoardData();

        var submitted = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "a.png" }],
            Components = [ReviewSummaryTests.Component("U8")],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "Baseline", File = "a.png" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "a.pdf" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Ref", Url = "https://example.com" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "Manuals", Name = "Service", File = "b.pdf" }],
            BoardLinks = [new BoardLinkEntry { Category = "Community", Name = "Forum", Url = "https://example.com" }],
            Credits = [new CreditEntry { Category = "Photos", SubCategory = "Scans", NameOrHandle = "Someone" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "/CLK" }]
        };

        // The ELEVENTH section rides alongside rather than inside the board: BoardData has no
        // calibration section, so it is passed separately. Added 2026-09-22, when calibrations
        // began travelling with submissions - before that they would have published unreviewed,
        // which is the exact failure this test exists to catch.
        ReviewChangeSummary summary = ReviewSummary.Compare(
            published,
            submitted,
            renames: null,
            publishedCalibrations: [],
            submittedCalibrations:
            [
                new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "board.kicad_pcb" }
            ]);

        Assert.Equal(11, summary.Sections.Count);
        Assert.Equal(11, summary.ChangedSections.Count);

        foreach (ReviewSectionChange section in summary.Sections)
        {
            Assert.Single(section.Added);
        }
    }

    [Fact]
    public void A_moved_highlight_is_reported_as_changed()
    {
        // Task 3's own example is "1 highlight moved". A highlight's key is schematic + label, so
        // moving one changes its coordinates without changing its identity.
        var published = new BoardData
        {
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "10", Y = "20", Width = "30", Height = "40" }]
        };

        var submitted = new BoardData
        {
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "15", Y = "20", Width = "30", Height = "40" }]
        };

        ReviewSectionChange highlights = ReviewSummaryTests.Section(
            ReviewSummary.Compare(published, submitted),
            "Component highlights");

        Assert.Single(highlights.Changed);
    }

    // -----------------------------------------------------------------------------------
    // The one-line description
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_description_names_the_section_and_the_verb()
    {
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8", part: "changed"),
            ReviewSummaryTests.Component("U9"));

        string description = ReviewSummary.Compare(published, submitted).Describe();

        Assert.Contains("1 components added", description);
        Assert.Contains("1 components changed", description);
    }

    [Fact]
    public void A_changed_revision_date_is_mentioned()
    {
        var published = new BoardData { RevisionDate = "2026-August-21" };
        var submitted = new BoardData { RevisionDate = "2026-September-21" };

        ReviewChangeSummary summary = ReviewSummary.Compare(published, submitted);

        Assert.True(summary.HasChanges);
        Assert.NotNull(summary.RevisionDate);
        Assert.Equal("2026-August-21", summary.RevisionDate!.Before);
        Assert.Equal("2026-September-21", summary.RevisionDate.After);
        Assert.Contains("revision date changed", summary.Describe());
    }

    // -----------------------------------------------------------------------------------
    // Robustness - a maintainer opens a submission to find out what is WRONG with it
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_duplicate_key_does_not_throw()
    {
        // Duplicate board labels are a validation error SubmissionValidator already reports. A
        // maintainer opening the submission to read that report must not be met with a crash
        // instead of the answer.
        BoardData published = ReviewSummaryTests.Board(ReviewSummaryTests.Component("U8"));
        BoardData submitted = ReviewSummaryTests.Board(
            ReviewSummaryTests.Component("U8"),
            ReviewSummaryTests.Component("U8", friendly: "duplicate"));

        ReviewChangeSummary summary = ReviewSummary.Compare(published, submitted);

        Assert.NotNull(summary);
        Assert.Empty(ReviewSummaryTests.Components(summary).Added);
    }

    [Fact]
    public void Two_empty_boards_compare_cleanly()
    {
        Assert.False(ReviewSummary.Compare(new BoardData(), new BoardData()).HasChanges);
    }

    [Fact]
    public void A_null_submitted_board_is_refused()
    {
        Assert.Throws<ArgumentNullException>(() => ReviewSummary.Compare(new BoardData(), null!));
    }
}
