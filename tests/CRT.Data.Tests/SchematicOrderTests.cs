using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SchematicOrder - a board's schematics in the order the contributor dragged them into in the
// Drafts tab's "Schematic images" window (2026-09-27).
//
// The order arrives as NAMES because the draft is re-read at save time and may have changed since
// the window read it; these pin that nothing is lost or invented when the two disagree.
// ###########################################################################################
public sealed class SchematicOrderTests
{
    private static BoardSchematicEntry Schematic(string name) =>
        new() { SchematicName = name, SchematicImageFile = $"Commodore/C128/310378 Open128/{name}.png" };

    // An order given by name alone - no image file to match.
    private static List<(string SchematicName, string? ImageFile)> ByName(params string[] names) =>
        names.Select(name => (name, (string?)null)).ToList();

    private static List<string> Names(IEnumerable<BoardSchematicEntry> schematics) =>
        schematics.Select(entry => entry.SchematicName).ToList();

    private static readonly List<BoardSchematicEntry> Board =
        [Schematic("PCB; Top"), Schematic("PCB; Bottom"), Schematic("Clock"), Schematic("CPU 8502")];

    [Fact]
    public void The_schematics_follow_the_order_given()
    {
        List<BoardSchematicEntry> ordered = SchematicOrder.Apply(Board, ByName("Clock", "PCB; Top", "CPU 8502", "PCB; Bottom"));

        Assert.Equal(["Clock", "PCB; Top", "CPU 8502", "PCB; Bottom"], Names(ordered));
    }

    // Only the ORDER changes: the rows are the board's own, with every value they carry.
    [Fact]
    public void The_rows_themselves_are_kept_not_rebuilt()
    {
        List<BoardSchematicEntry> ordered = SchematicOrder.Apply(Board, ByName("Clock", "PCB; Top", "CPU 8502", "PCB; Bottom"));

        Assert.Same(Board[2], ordered[0]);
        Assert.Same(Board[0], ordered[1]);
    }

    // A schematic added to the draft outside the window (in Excel, say) is not dropped.
    [Fact]
    public void A_schematic_the_order_does_not_name_keeps_its_place_after_the_named_ones()
    {
        List<BoardSchematicEntry> ordered = SchematicOrder.Apply(Board, ByName("CPU 8502", "Clock"));

        Assert.Equal(["CPU 8502", "Clock", "PCB; Top", "PCB; Bottom"], Names(ordered));
    }

    // And one removed outside the window is not invented.
    [Fact]
    public void A_name_no_schematic_has_is_skipped()
    {
        List<BoardSchematicEntry> ordered = SchematicOrder.Apply(Board, ByName("Clock", "Gone", "PCB; Top", "CPU 8502", "PCB; Bottom"));

        Assert.Equal(["Clock", "PCB; Top", "CPU 8502", "PCB; Bottom"], Names(ordered));
    }

    [Fact]
    public void Names_match_ignoring_case_and_outer_spaces_like_every_schematic_name()
    {
        List<BoardSchematicEntry> ordered = SchematicOrder.Apply(Board, ByName("  clock ", "cpu 8502", "PCB; TOP", "pcb; bottom"));

        Assert.Equal(["Clock", "CPU 8502", "PCB; Top", "PCB; Bottom"], Names(ordered));
    }

    // A hand-edited workbook can carry two rows with one name; both survive, in the order given.
    [Fact]
    public void Two_schematics_sharing_a_name_both_survive()
    {
        BoardSchematicEntry first = Schematic("Top");
        BoardSchematicEntry second = Schematic("Top");

        List<BoardSchematicEntry> ordered = SchematicOrder.Apply([first, Schematic("Bottom"), second], ByName("Bottom", "Top", "Top"));

        Assert.Equal(3, ordered.Count);
        Assert.Equal("Bottom", ordered[0].SchematicName);
        Assert.Same(first, ordered[1]);
        Assert.Same(second, ordered[2]);
    }

    // ###########################################################################################
    // Two rows sharing a name, told apart by their image files: the one DRAGGED to the top is the
    // one put there. By name alone the first of the name always took the first place, so dragging
    // the second above it saved nothing (code review, 2026-10-10).
    // ###########################################################################################
    [Fact]
    public void Two_schematics_sharing_a_name_go_where_their_image_files_say()
    {
        BoardSchematicEntry first = new() { SchematicName = "Top", SchematicImageFile = "Commodore/C64/250407/top-a.png" };
        BoardSchematicEntry second = new() { SchematicName = "Top", SchematicImageFile = "Commodore/C64/250407/top-b.png" };

        List<BoardSchematicEntry> ordered = SchematicOrder.Apply(
            [first, second],
            [("Top", "Commodore\\C64\\250407\\TOP-B.png"), ("Top", "Commodore/C64/250407/top-a.png")]);

        Assert.Same(second, ordered[0]);
        Assert.Same(first, ordered[1]);
    }

    [Fact]
    public void No_order_leaves_the_schematics_as_they_were()
    {
        Assert.Equal(Names(Board), Names(SchematicOrder.Apply(Board, ByName())));
        Assert.Empty(SchematicOrder.Apply([], ByName("Clock")));
    }
}
