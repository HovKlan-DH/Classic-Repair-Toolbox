using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Covers the FIELD-LEVEL diff (NewContributeStrategy.md Phase 5, task 4's "text rows as a
// field-level diff").
//
// WHY THIS IS THE PART THAT MAKES REVIEWING POSSIBLE: "U8 changed" is not something anyone can act
// on. A maintainer given only that has to open the board in CRT, find U8, and compare it by eye
// against a submission they cannot see side by side. "Part-number: 906114 -> 251715-01" IS the
// decision - it takes a second and needs nothing else open.
//
// The field NAMES come from BoardWorkbookSchema, so a maintainer reading "Part-number" and
// someone opening the .xlsx to check are talking about the same column.
public sealed class ReviewFieldDiffTests
{
    private static ComponentEntry Component(
        string label = "U8",
        string friendly = "PLA",
        string technical = "906114-01",
        string part = "906114",
        string category = "IC",
        string region = "",
        string description = "") => new()
        {
            BoardLabel = label,
            FriendlyName = friendly,
            TechnicalNameOrValue = technical,
            PartNumber = part,
            Category = category,
            Region = region,
            Description = description
        };

    private static ReviewSectionChange Components(BoardData published, BoardData submitted) =>
        ReviewSummary.Compare(published, submitted).Sections
            .First(section => section.Section == BoardWorkbookSchema.SheetComponents);

    private static BoardData Board(params ComponentEntry[] components) => new()
    {
        RevisionDate = "2026-August-21",
        Components = [.. components]
    };

    [Fact]
    public void A_changed_row_reports_WHICH_field_moved_and_to_what()
    {
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(part: "906114"));
        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(part: "251715-01"));

        ReviewSectionChange section = ReviewFieldDiffTests.Components(published, submitted);

        ReviewFieldChange change = Assert.Single(section.FieldChanges["U8"]);

        Assert.Equal(BoardWorkbookSchema.ColPartNumber, change.Field);
        Assert.Equal("906114", change.Before);
        Assert.Equal("251715-01", change.After);
    }

    [Fact]
    public void The_field_name_is_the_WORKBOOK_COLUMN_HEADER()
    {
        // A maintainer reading this and someone opening the .xlsx must be talking about the
        // same column. Inventing a friendlier label here would create a second vocabulary for the
        // same data, and the two would drift.
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(description: "old"));
        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(description: "new"));

        ReviewFieldChange change = Assert.Single(
            ReviewFieldDiffTests.Components(published, submitted).FieldChanges["U8"]);

        // The real header, parenthetical and all - that IS the column's name in every shipped
        // workbook, and dropping the parenthetical would rename it.
        Assert.Equal("Short one-liner description (one short line only!)", change.Field);
    }

    [Fact]
    public void SEVERAL_changed_fields_are_all_reported()
    {
        BoardData published = ReviewFieldDiffTests.Board(
            ReviewFieldDiffTests.Component(part: "906114", category: "IC"));

        BoardData submitted = ReviewFieldDiffTests.Board(
            ReviewFieldDiffTests.Component(part: "251715-01", category: "PLA"));

        IReadOnlyList<ReviewFieldChange> changes =
            ReviewFieldDiffTests.Components(published, submitted).FieldChanges["U8"];

        Assert.Equal(2, changes.Count);
        Assert.Contains(changes, change => change.Field == BoardWorkbookSchema.ColPartNumber);
        Assert.Contains(changes, change => change.Field == BoardWorkbookSchema.ColCategory);
    }

    [Fact]
    public void UNCHANGED_fields_are_not_reported()
    {
        // The whole value of this screen is that it shows only what moved. Listing every field of
        // a changed row would bury the edit among six identical lines.
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(part: "906114"));
        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(part: "251715-01"));

        ReviewFieldChange change = Assert.Single(
            ReviewFieldDiffTests.Components(published, submitted).FieldChanges["U8"]);

        Assert.Equal(BoardWorkbookSchema.ColPartNumber, change.Field);
    }

    [Fact]
    public void A_CLEARED_field_is_reported_with_an_empty_after()
    {
        // *** DELETING INFORMATION IS THE EDIT HARDEST TO NOTICE, and the one a maintainer most
        // needs shown. *** Omitting a cleared field because its new value is blank would hide
        // exactly the change that deserves a second look.
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(description: "Decodes the memory map"));
        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(description: ""));

        ReviewFieldChange change = Assert.Single(
            ReviewFieldDiffTests.Components(published, submitted).FieldChanges["U8"]);

        Assert.Equal("Decodes the memory map", change.Before);
        Assert.Equal(string.Empty, change.After);
    }

    [Fact]
    public void A_field_FILLED_IN_for_the_first_time_is_reported_with_an_empty_before()
    {
        // Part-number, not Region as this test first used: since 2026-09-24 a component's REGION
        // is part of its identity (BoardDraftNaturalKeys.ForComponent), so filling one in makes a
        // different row - see A_component_given_a_region_is_a_removal_plus_an_addition below. The
        // rule under test is about any ordinary field.
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(part: ""));
        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(part: "906114"));

        ReviewFieldChange change = Assert.Single(
            ReviewFieldDiffTests.Components(published, submitted).FieldChanges["U8"]);

        Assert.Equal(string.Empty, change.Before);
        Assert.Equal("906114", change.After);
    }

    // ###########################################################################################
    // *** A REGIONAL VARIANT IS A ROW OF ITS OWN (2026-09-24). *** A component stored once per
    // region - U8 for PAL, U8 for NTSC - used to key on the label alone, so a submission adding the
    // NTSC row beside an existing PAL one showed the maintainer NOTHING: the two collapsed into one.
    // Hiding an added row from the person approving it is what this screen exists to prevent.
    // ###########################################################################################
    [Fact]
    public void A_regional_variant_added_beside_an_existing_one_is_shown_to_the_maintainer_as_added()
    {
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(region: "PAL"));
        BoardData submitted = ReviewFieldDiffTests.Board(
            ReviewFieldDiffTests.Component(region: "PAL"),
            ReviewFieldDiffTests.Component(region: "NTSC"));

        ReviewSectionChange section = ReviewFieldDiffTests.Components(published, submitted);

        Assert.Single(section.Added);
        Assert.Empty(section.Removed);
        Assert.Empty(section.Changed);
    }

    [Fact]
    public void A_component_given_a_region_and_nothing_else_is_one_row_renamed()
    {
        // Region is part of a component's identity, so giving it one changes its key - as changing
        // its label does. It was a removal plus an addition until 2026-10-04; with nothing else of
        // the row changed it is now the same row, renamed (BoardDataDiffer.PairRenamedRows).
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(region: ""));
        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(region: "PAL"));

        ReviewSectionChange section = ReviewFieldDiffTests.Components(published, submitted);

        ReviewRenamedRow renamed = Assert.Single(section.Renamed);
        Assert.Equal("U8", renamed.From);
        Assert.Equal(BoardDraftNaturalKeys.ForComponent("U8", "PAL"), renamed.To);
        Assert.False(renamed.AlsoChanged);
        Assert.Empty(section.Added);
        Assert.Empty(section.Removed);
    }

    [Fact]
    public void A_component_given_a_region_AND_a_new_part_number_is_a_removal_plus_an_addition()
    {
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(region: ""));
        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component(region: "PAL", part: "251715-01"));

        ReviewSectionChange section = ReviewFieldDiffTests.Components(published, submitted);

        Assert.Single(section.Added);
        Assert.Equal(["U8"], section.Removed);
    }

    [Fact]
    public void An_ADDED_row_carries_no_field_diff()
    {
        // There is nothing to diff against - every field is new. The row itself is the change,
        // and listing seven "(blank) -> value" lines would be noise.
        BoardData published = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component("U8"));
        BoardData submitted = ReviewFieldDiffTests.Board(
            ReviewFieldDiffTests.Component("U8"),
            ReviewFieldDiffTests.Component("U9"));

        ReviewSectionChange section = ReviewFieldDiffTests.Components(published, submitted);

        Assert.Equal(["U9"], section.Added);
        Assert.False(section.FieldChanges.ContainsKey("U9"));
    }

    [Fact]
    public void A_REMOVED_row_carries_no_field_diff()
    {
        BoardData published = ReviewFieldDiffTests.Board(
            ReviewFieldDiffTests.Component("U8"),
            ReviewFieldDiffTests.Component("U9"));

        BoardData submitted = ReviewFieldDiffTests.Board(ReviewFieldDiffTests.Component("U8"));

        ReviewSectionChange section = ReviewFieldDiffTests.Components(published, submitted);

        Assert.Equal(["U9"], section.Removed);
        Assert.False(section.FieldChanges.ContainsKey("U9"));
    }

    [Fact]
    public void A_highlight_that_MOVED_reports_its_coordinates()
    {
        // Task 4's own example is "1 highlight moved". Until the schematic can be drawn, the
        // coordinates are what tells a maintainer how far it moved - "10 -> 15" is a nudge,
        // "10 -> 900" is a different part of the board entirely.
        var published = new BoardData
        {
            ComponentHighlights = [new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1", BoardLabel = "U8", X = "10", Y = "20", Width = "30", Height = "40"
            }]
        };

        var submitted = new BoardData
        {
            ComponentHighlights = [new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1", BoardLabel = "U8", X = "900", Y = "20", Width = "30", Height = "40"
            }]
        };

        ReviewSectionChange section = ReviewSummary.Compare(published, submitted).Sections
            .First(candidate => candidate.Section == "Component highlights");

        ReviewFieldChange change = Assert.Single(section.FieldChanges.Values.Single());

        Assert.Equal("X", change.Field);
        Assert.Equal("10", change.Before);
        Assert.Equal("900", change.After);
    }

    [Fact]
    public void A_retired_UUID_produces_no_field_change()
    {
        // UuidV4 is excluded from the comparison entirely, so it must not appear as a field
        // change either - otherwise every pre-transition board would report one on every row.
        var published = new BoardData
        {
            Components = [new ComponentEntry
            {
                BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "906114"
            }]
        };

        var submitted = new BoardData
        {
            Components = [new ComponentEntry
            {
                BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "251715-01"
            }]
        };

        ReviewFieldChange change = Assert.Single(
            ReviewFieldDiffTests.Components(published, submitted).FieldChanges["U8"]);

        Assert.Equal(BoardWorkbookSchema.ColPartNumber, change.Field);
    }

    [Fact]
    public void Every_section_can_produce_a_field_diff()
    {
        // A section whose rows changed but which reports no field detail would leave a maintainer
        // back at "something changed". This walks one edit through each of the ten - nine of which
        // CHANGE a row. Important signals cannot: since 2026-09-26 both of their columns are the
        // row's key (one display name covers several nets), so a new net is the row RENAMED (it was
        // a removal plus an addition until 2026-10-04, when a row whose key alone changed became one
        // row) - pinned at the end.
        var published = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "S", SchematicImageFile = "a.png", CadName = "before" }],
            Components = [ReviewFieldDiffTests.Component()],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "N", File = "a.png", Note = "before" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "S", BoardLabel = "U8", X = "1" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "N", File = "before.pdf" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "N", Url = "https://before" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "C", Name = "N", File = "before.pdf" }],
            BoardLinks = [new BoardLinkEntry { Category = "C", Name = "N", Url = "https://before" }],
            Credits = [new CreditEntry { Category = "C", SubCategory = "S", NameOrHandle = "N", Contact = "before" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "D", KiCadNetName = "before" }]
        };

        var submitted = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "S", SchematicImageFile = "a.png", CadName = "after" }],
            Components = [ReviewFieldDiffTests.Component(part: "changed")],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "N", File = "a.png", Note = "after" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "S", BoardLabel = "U8", X = "2" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "N", File = "after.pdf" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "N", Url = "https://after" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "C", Name = "N", File = "after.pdf" }],
            BoardLinks = [new BoardLinkEntry { Category = "C", Name = "N", Url = "https://after" }],
            Credits = [new CreditEntry { Category = "C", SubCategory = "S", NameOrHandle = "N", Contact = "after" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "D", KiCadNetName = "after" }]
        };

        ReviewChangeSummary summary = ReviewSummary.Compare(published, submitted);

        Assert.Equal(10, summary.ChangedSections.Count);

        foreach (ReviewSectionChange section in summary.ChangedSections.Where(section => section.Section != BoardWorkbookSchema.SheetKiCadImportantSignals))
        {
            Assert.Single(section.Changed);
            Assert.NotEmpty(section.FieldChanges);
            Assert.Single(section.FieldChanges.Values.Single());
        }

        ReviewSectionChange signals = summary.ChangedSections.Single(section => section.Section == BoardWorkbookSchema.SheetKiCadImportantSignals);
        Assert.Empty(signals.Changed);
        Assert.Empty(signals.Removed);
        Assert.Empty(signals.Added);
        Assert.Single(signals.Renamed);
    }
}
