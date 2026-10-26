using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Bringing a draft up to date with the published board - the three-way merge (code review,
// 2026-10-13).
//
// *** THE BUG THIS EXISTS FOR. *** A draft is a complete copy of the board and a submission is the
// board's complete intended state, so a draft started in September did not hold what October
// published - and sending it showed the maintainer every row published since as a deletion by
// this contributor. The drift window promised "your edits are still applied on top of it", which
// was not true. These pin what IS true now: what the contributor changed is carried onto the newer
// board, and nothing else of theirs is.
// ###########################################################################################
public sealed class DraftRebaseTests
{
    private static ComponentEntry Component(string label, string description = "", string region = "") =>
        new() { BoardLabel = label, Description = description, Region = region };

    private static ComponentHighlightEntry Highlight(string schematic, string label, string x = "1") =>
        new() { SchematicName = schematic, BoardLabel = label, X = x, Y = "1", Width = "10", Height = "10" };

    private static BoardSchematicEntry Schematic(string name, string file) =>
        new() { SchematicName = name, SchematicImageFile = file };

    private static BoardData Board(
        IEnumerable<ComponentEntry>? components = null,
        IEnumerable<ComponentHighlightEntry>? highlights = null,
        IEnumerable<BoardSchematicEntry>? schematics = null,
        string revision = "2026-09-01") => new()
    {
        RevisionDate = revision,
        HardwareName = "Commodore 64",
        BoardName = "250407",
        Schematics = (schematics ?? [Schematic("Sheet 1", "b/Sheet1.png")]).ToList(),
        Components = (components ?? []).ToList(),
        ComponentHighlights = (highlights ?? []).ToList(),
    };

    private static DraftRebaseResult Merge(BoardData based, BoardData draft, BoardData published) =>
        DraftRebase.Merge(based, draft, published, [], [], []);

    private static string[] Labels(BoardData board) => board.Components.Select(c => c.BoardLabel).ToArray();

    private static string DescriptionOf(BoardData board, string label) =>
        board.Components.Single(c => c.BoardLabel == label).Description;

    // ###########################################################################################
    // CRT 4.0's pins and replacement parts (2026-10-25) merge three ways like every other section:
    // the contributor's pin edit stays, the pin and replacement the update added arrive.
    // ###########################################################################################
    [Fact]
    public void Pins_and_replacements_merge_three_ways()
    {
        BoardData based = Board(components: [Component("U8")]);
        based.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" });

        BoardData draft = Board(components: [Component("U8")]);
        draft.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "GND" });

        BoardData published = Board(components: [Component("U8")], revision: "2026-10-01");
        published.ComponentPins.AddRange([
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" },
            new ComponentPinEntry { BoardLabel = "U8", Pin = "2", Name = "PA0" }]);
        published.ComponentReplacements.Add(new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "6526A" });

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["1 GND", "2 PA0"], result.Board.ComponentPins.Select(pin => $"{pin.Pin} {pin.Name}"));
        Assert.Equal("6526A", Assert.Single(result.Board.ComponentReplacements).PartNumber);
    }

    [Fact]
    public void A_stale_draft_with_one_change_keeps_everything_published_since()
    {
        // The review's scenario: one fix in September, forty components published in October.
        BoardData based = Board([Component("U8", "CPU"), Component("U9", "VIC")]);
        BoardData draft = Board([Component("U8", "6510 CPU"), Component("U9", "VIC")]);
        BoardData published = Board(
            [Component("U8", "CPU"), Component("U9", "VIC"), .. Enumerable.Range(1, 40).Select(i => Component($"C{i}", "Cap"))],
            revision: "2026-10-01");

        DraftRebaseResult result = Merge(based, draft, published);

        // Against the published board, the draft now says ONE thing: the contributor's fix.
        Assert.Equal(1, BoardDataDiffer.CountChanges(published, result.Board));
        Assert.Equal("6510 CPU", DescriptionOf(result.Board, "U8"));
        Assert.Equal(42, result.Board.Components.Count);
        Assert.Equal(1, result.ChangesKept);
        Assert.Empty(result.Conflicts);
        Assert.Equal("2026-10-01", result.Board.RevisionDate);
    }

    // ###########################################################################################
    // *** A ROW THE UPDATE RELABELLED, EDITED BY THE CONTRIBUTOR (code review, 2026-10-13). *** The
    // merge paired only the contributor's renames: "U8" came back beside the update's "U08", with a
    // false "The update removed this row" - and sending it re-added U8. Now the contributor's edit
    // lands on U08, with nothing said, since the two changed different things.
    // ###########################################################################################
    [Fact]
    public void An_edit_to_a_row_the_update_relabelled_lands_on_the_new_label()
    {
        BoardData based = Board([Component("U8", "CPU"), Component("U9", "VIC")]);
        BoardData draft = Board([Component("U8", "6510 CPU"), Component("U9", "VIC")]);
        BoardData published = Board([Component("U08", "CPU"), Component("U9", "VIC")], revision: "2026-10-01");

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["U08", "U9"], Labels(result.Board));
        Assert.Equal("6510 CPU", DescriptionOf(result.Board, "U08"));
        Assert.Empty(result.Conflicts);
        Assert.Equal(1, BoardDataDiffer.CountChanges(published, result.Board));
    }

    // Left alone, removed, or with highlights the contributor added: each follows the relabel.
    [Fact]
    public void A_relabel_by_the_update_is_followed_by_the_contributors_removal_and_highlights()
    {
        BoardData based = Board(
            [Component("U8", "CPU"), Component("U9", "VIC")],
            [Highlight("Sheet 1", "U8", x: "1")]);

        // The contributor removed U9 and marked U8 a second time.
        BoardData draft = Board(
            [Component("U8", "CPU")],
            [Highlight("Sheet 1", "U8", x: "1"), Highlight("Sheet 1", "U8", x: "50")]);

        // The update relabelled both, and its highlight followed U8.
        BoardData published = Board(
            [Component("U08", "CPU"), Component("U09", "VIC")],
            [Highlight("Sheet 1", "U08", x: "1")],
            revision: "2026-10-01");

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["U08"], Labels(result.Board));
        Assert.Equal(
            ["U08@1", "U08@50"],
            result.Board.ComponentHighlights.Select(h => $"{h.BoardLabel}@{h.X}").OrderBy(text => text, System.StringComparer.Ordinal).ToArray());
    }

    [Fact]
    public void A_row_the_contributor_left_alone_follows_the_update_changed_or_removed()
    {
        BoardData based = Board([Component("U8", "CPU"), Component("U9", "VIC"), Component("U10", "SID")]);
        BoardData draft = Board([Component("U8", "CPU"), Component("U9", "VIC"), Component("U10", "SID")]);
        BoardData published = Board([Component("U8", "CPU 6510"), Component("U10", "SID")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["U8", "U10"], Labels(result.Board));
        Assert.Equal("CPU 6510", DescriptionOf(result.Board, "U8"));
        Assert.Equal(0, result.ChangesKept);
    }

    [Fact]
    public void A_row_the_contributor_removed_stays_removed_and_one_they_added_is_kept_in_place()
    {
        BoardData based = Board([Component("U1"), Component("U2"), Component("U3")]);
        BoardData draft = Board([Component("U1"), Component("U1A", "added after U1"), Component("U3")]);
        BoardData published = Board([Component("U0", "new first"), Component("U1"), Component("U2"), Component("U3")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["U0", "U1", "U1A", "U3"], Labels(result.Board));
        Assert.Equal(2, result.ChangesKept);
        Assert.Empty(result.Conflicts);
    }

    [Fact]
    public void When_both_sides_changed_a_row_the_contributors_version_is_kept_and_said()
    {
        BoardData based = Board([Component("U8", "CPU")]);
        BoardData draft = Board([Component("U8", "Mine")]);
        BoardData published = Board([Component("U8", "Theirs")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal("Mine", DescriptionOf(result.Board, "U8"));
        DraftRebaseConflict conflict = Assert.Single(result.Conflicts);
        Assert.Equal("Components", conflict.Section);
        Assert.Contains("U8", conflict.Row);
        Assert.Contains("both changed", conflict.What);
    }

    [Fact]
    public void The_same_change_on_both_sides_is_no_conflict()
    {
        BoardData based = Board([Component("U8", "CPU")]);
        BoardData draft = Board([Component("U8", "6510")]);
        BoardData published = Board([Component("U8", "6510")]);

        Assert.Empty(Merge(based, draft, published).Conflicts);
    }

    [Fact]
    public void A_row_the_update_removed_but_the_contributor_changed_is_kept_and_said()
    {
        BoardData based = Board([Component("U8", "CPU"), Component("U9")]);
        BoardData draft = Board([Component("U8", "Mine"), Component("U9")]);
        BoardData published = Board([Component("U9")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["U8", "U9"], Labels(result.Board));
        Assert.Contains("removed", Assert.Single(result.Conflicts).What);
    }

    [Fact]
    public void A_row_the_update_changed_but_the_contributor_removed_stays_removed_and_said()
    {
        BoardData based = Board([Component("U8", "CPU"), Component("U9")]);
        BoardData draft = Board([Component("U9")]);
        BoardData published = Board([Component("U8", "Theirs"), Component("U9")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["U9"], Labels(result.Board));
        Assert.Contains("stays removed", Assert.Single(result.Conflicts).What);
    }

    [Fact]
    public void A_renamed_row_is_one_rename_carried_onto_the_update()
    {
        // A key edit with nothing else changed is ONE row (owner decision, 2026-10-04) - so the
        // update's other rows stay, and the renamed row takes the old one's place.
        BoardData based = Board([Component("U8", "CPU"), Component("U9", "VIC")]);
        BoardData draft = Board([Component("U08", "CPU"), Component("U9", "VIC")]);
        BoardData published = Board([Component("U8", "CPU"), Component("U9", "VIC II")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["U08", "U9"], Labels(result.Board));
        Assert.Equal("VIC II", DescriptionOf(result.Board, "U9"));
        Assert.Equal(1, result.ChangesKept);
        Assert.Empty(result.Conflicts);
    }

    [Fact]
    public void Duplicate_rows_are_kept_pair_by_pair()
    {
        // Two "C1" rows are a warning in the table, not an error - a merge must not lose one.
        BoardData based = Board([Component("C1", "first"), Component("C1", "second")]);
        BoardData draft = Board([Component("C1", "first"), Component("C1", "second, corrected")]);
        BoardData published = Board([Component("C1", "first"), Component("C1", "second"), Component("C2")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal(["first", "second, corrected", ""], result.Board.Components.Select(c => c.Description).ToArray());
    }

    [Fact]
    public void Keys_compare_ignoring_case_as_the_count_does()
    {
        BoardData based = Board([Component("U8", "CPU")]);
        BoardData draft = Board([Component("u8", "CPU")]);
        BoardData published = Board([Component("U8", "CPU 6510")]);

        // A capitalisation-only key edit with nothing else changed is the same row (IndexByKey);
        // its values are compared exactly, so "u8" vs "U8" is the contributor's change.
        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Single(result.Board.Components);
        Assert.Equal("u8", result.Board.Components[0].BoardLabel);
    }

    [Fact]
    public void Highlights_merge_rectangle_by_rectangle()
    {
        BoardData based = Board(highlights: [Highlight("Sheet 1", "U8", "1"), Highlight("Sheet 1", "U9", "5")]);
        BoardData draft = Board(highlights: [Highlight("Sheet 1", "U8", "1"), Highlight("Sheet 1", "U9", "5"), Highlight("Sheet 1", "U10", "9")]);
        BoardData published = Board(highlights: [Highlight("Sheet 1", "U8", "2"), Highlight("Sheet 1", "U9", "5")]);

        DraftRebaseResult result = Merge(based, draft, published);

        // The update moved U8's rectangle (the contributor had not touched it); the contributor
        // added U10's. Both are in.
        Assert.Equal(
            ["U8@2", "U9@5", "U10@9"],
            result.Board.ComponentHighlights.Select(h => $"{h.BoardLabel}@{h.X}").ToArray());
    }

    [Fact]
    public void A_contributors_highlight_follows_a_schematic_the_update_renamed()
    {
        BoardData based = Board(schematics: [Schematic("Sheet 1", "b/Sheet1.png")]);
        BoardData draft = Board(
            schematics: [Schematic("Sheet 1", "b/Sheet1.png")],
            highlights: [Highlight("Sheet 1", "U8")]);
        BoardData published = Board(schematics: [Schematic("Main board", "b/Sheet1.png")]);

        DraftRebaseResult result = Merge(based, draft, published);

        Assert.Equal("Main board", Assert.Single(result.Board.ComponentHighlights).SchematicName);
    }

    [Fact]
    public void Calibrations_merge_three_ways_and_follow_a_renamed_schematic()
    {
        static KiCadCalibrationEntry Calibration(string schematic, double scale) =>
            new() { SchematicName = schematic, CadName = "pcb", ScaleX = scale, ScaleY = scale };

        BoardData based = Board(schematics: [Schematic("Sheet 1", "b/Sheet1.png"), Schematic("Sheet 2", "b/Sheet2.png")]);
        BoardData published = Board(schematics: [Schematic("Top", "b/Sheet1.png"), Schematic("Sheet 2", "b/Sheet2.png")]);

        DraftRebaseResult result = DraftRebase.Merge(
            based,
            based,
            published,
            baseCalibrations: [Calibration("Sheet 1", 1.0), Calibration("Sheet 2", 1.0)],
            draftCalibrations: [Calibration("Sheet 1", 1.5), Calibration("Sheet 2", 1.0)],
            publishedCalibrations: [Calibration("Top", 1.0), Calibration("Sheet 2", 2.0)]);

        // The contributor recalibrated Sheet 1 (now "Top"); the update recalibrated Sheet 2.
        Assert.Equal(
            ["Sheet 2=2", "Top=1.5"],
            result.Calibrations.Select(c => FormattableString.Invariant($"{c.SchematicName}={c.ScaleX}")).OrderBy(s => s, StringComparer.Ordinal).ToArray());
    }

    [Fact]
    public void An_untouched_draft_becomes_the_published_board()
    {
        BoardData based = Board([Component("U8", "CPU")]);
        BoardData published = Board([Component("U8", "CPU 6510"), Component("U9")], revision: "2026-10-02");

        DraftRebaseResult result = Merge(based, based, published);

        Assert.Equal(0, BoardDataDiffer.CountChanges(published, result.Board));
        Assert.Equal(0, result.ChangesKept);
    }
}
