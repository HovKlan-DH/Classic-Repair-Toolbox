using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardData's "same board, one thing different" copies.
//
// *** WHY BY REFLECTION. *** The Schematic images window had its own copy that listed the sections
// by hand and forgot HardwareName and BoardName, so every image import or removal wrote the draft
// back without its "# Hardware:" / "# Board:" caption (2026-09-27). A test naming the properties
// would have forgotten them the same way; walking every property cannot, and it also fails the day
// someone adds a property to BoardData and not to these copies.
// ###########################################################################################
public sealed class BoardDataTests
{
    // A board with a DIFFERENT, non-default value in every property, so a property the copy leaves
    // out shows up as a default where the original's value should be.
    private static BoardData EveryPropertySet() => new()
    {
        RevisionDate = "2026-September-27",
        HardwareName = "C128",
        BoardName = "310378 Open128",
        Schematics = [new BoardSchematicEntry { SchematicName = "Clock" }],
        Components = [new ComponentEntry { BoardLabel = "U1" }],
        ComponentImages = [new ComponentImageEntry { BoardLabel = "U1" }],
        ComponentHighlights = [new ComponentHighlightEntry { BoardLabel = "U1" }],
        ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U1" }],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1" }],
        BoardLocalFiles = [new BoardLocalFileEntry { Name = "Manual" }],
        BoardLinks = [new BoardLinkEntry { Name = "Forum" }],
        Credits = [new CreditEntry { NameOrHandle = "Someone" }],
        KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "9VAC" }],
    };

    [Fact]
    public void The_fixture_really_sets_every_property()
    {
        BoardData blank = new();
        BoardData full = EveryPropertySet();

        foreach (PropertyInfo property in typeof(BoardData).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            Assert.False(
                Equals(property.GetValue(blank), property.GetValue(full)) || IsEmptyList(property.GetValue(full)),
                $"The fixture leaves [{property.Name}] at its default - give it a value, or a copy that drops it passes.");
        }
    }

    [Fact]
    public void WithSchematics_replaces_the_schematics_and_keeps_everything_else()
    {
        BoardData board = EveryPropertySet();
        List<BoardSchematicEntry> replacement = [new BoardSchematicEntry { SchematicName = "CPU 8502" }];

        BoardData copy = board.WithSchematics(replacement);

        Assert.Same(replacement, copy.Schematics);
        AssertSameExcept(board, copy, nameof(BoardData.Schematics));
    }

    private static void AssertSameExcept(BoardData original, BoardData copy, string changed)
    {
        foreach (PropertyInfo property in typeof(BoardData).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (property.Name == changed)
            {
                continue;
            }

            Assert.True(
                Equals(property.GetValue(original), property.GetValue(copy)),
                $"The copy dropped or changed [{property.Name}].");
        }
    }

    private static bool IsEmptyList(object? value) =>
        value is System.Collections.ICollection collection && collection.Count == 0;
}
