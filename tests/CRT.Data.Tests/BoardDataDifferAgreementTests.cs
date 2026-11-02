using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// THE CONTRIBUTOR'S COUNT AND THE MAINTAINER'S SUMMARY, ON ONE INPUT (code review, 2026-10-13).
//
// "N rows changed" on the Drafts tab is BoardDataDiffer; the maintainer's summary, the queue line
// and the board's history are ReviewSummary. Two code paths that must give the same number - and
// a test per side, each asserting its own half, passed happily while they disagreed: ReviewSummary
// keyed rows by exact case, the differ ignoring it, so a Credits row differing only in capitals was
// "0 rows changed" for the contributor and "1 added" for the maintainer. Every case here goes
// through BOTH, and fails if only one side moves.
// ###########################################################################################
public sealed class BoardDataDifferAgreementTests
{
    private static ComponentEntry Component(string label, string description = "", string region = "") =>
        new() { BoardLabel = label, Description = description, Region = region };

    private static CreditEntry Credit(string name, string contact = "") =>
        new() { Category = "Hardware", SubCategory = "-", NameOrHandle = name, Contact = contact };

    private static ComponentHighlightEntry Rectangle(string x, string label = "U8") =>
        new() { SchematicName = "Sheet 1", BoardLabel = label, X = x, Y = "2", Width = "3", Height = "4" };

    private static BoardData Board(IEnumerable<ComponentEntry>? components = null, IEnumerable<CreditEntry>? credits = null, IEnumerable<ComponentHighlightEntry>? rectangles = null) => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "b/Sheet1.png" }],
        Components = (components ?? []).ToList(),
        Credits = (credits ?? []).ToList(),
        ComponentHighlights = (rectangles ?? []).ToList(),
    };

    public static TheoryData<string, BoardData, BoardData> Cases => new()
    {
        { "nothing changed", Board([Component("U8", "CPU")]), Board([Component("U8", "CPU")]) },
        { "a value changed", Board([Component("U8", "CPU")]), Board([Component("U8", "6510")]) },
        { "a row added", Board([Component("U8")]), Board([Component("U8"), Component("U9")]) },
        { "a row removed", Board([Component("U8"), Component("U9")]), Board([Component("U8")]) },
        { "a key edited alone is one rename", Board([Component("U8", "CPU")]), Board([Component("U08", "CPU")]) },
        { "a key AND a value edited", Board([Component("U8", "CPU")]), Board([Component("U08", "6510")]) },
        { "capitals only in a key", Board([Component("U8", "CPU")]), Board([Component("u8", "CPU")]) },
        { "trailing space in a value", Board([Component("U8", "CPU")]), Board([Component("U8", "CPU ")]) },
        { "a regional twin", Board([Component("U8", "CPU", "PAL")]), Board([Component("U8", "CPU", "PAL"), Component("U8", "CPU", "NTSC")]) },
        {
            "a duplicate differing only in capitals (the review's case)",
            Board(credits: [Credit("Dennis", "a@example.org")]),
            Board(credits: [Credit("Dennis", "a@example.org"), Credit("dennis", "b@example.org")])
        },

        // Several rectangles of one component on one schematic (code review, 2026-11-02): only the first
        // of them was compared, on both sides.
        { "a second rectangle unchanged", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("1"), Rectangle("10")]) },
        { "a second rectangle moved", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("1"), Rectangle("20")]) },
        { "both rectangles moved", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("5"), Rectangle("20")]) },
        { "a third rectangle added", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("1"), Rectangle("10"), Rectangle("20")]) },
        { "a second rectangle deleted", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("10")]) },
        { "two rectangles in another order", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("10"), Rectangle("1")]) },
        { "a component relabelled with its two rectangles", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("1", "U08"), Rectangle("10", "U08")]) },
        { "a rectangle whose label reads like the next of a key", Board(rectangles: [Rectangle("1"), Rectangle("10")]), Board(rectangles: [Rectangle("99", "U8␞1"), Rectangle("1"), Rectangle("10")]) },
    };

    // The same row under a plain label and under one holding that character: one change either way, on both
    // sides - and the very key it has, nothing cut off it (code review, 2026-11-02, Security).
    [Theory]
    [InlineData("U9")]
    [InlineData("U8␞1")]
    public void A_rectangle_added_is_one_change_on_both_sides_whatever_its_label_holds(string label)
    {
        BoardData published = Board(rectangles: [Rectangle("1"), Rectangle("10")]);
        BoardData submitted = Board(rectangles: [Rectangle("99", label), Rectangle("1"), Rectangle("10")]);

        string key = BoardDraftNaturalKeys.ForComponentHighlight("Sheet 1", label);

        Assert.Equal(key, Assert.Single(BoardDataDiffer.Compare(published, submitted)).NaturalKey);
        Assert.Equal(key, Assert.Single(ReviewSummary.Compare(published, submitted).Sections.Single(section => section.Section == ReviewSummary.SectionComponentHighlights).Added));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void The_count_and_the_summary_agree(string because, BoardData published, BoardData draft)
    {
        int count = BoardDataDiffer.CountChanges(published, draft);
        int summary = ReviewSummary.Compare(published, draft).TotalChanges;

        Assert.True(count == summary, $"{because}: the contributor's count says {count}, the maintainer's summary {summary}");
    }

    // ###########################################################################################
    // The summary lists a row's fields BY HAND (it names them by their workbook columns); the
    // differ walks every property. A column added to an entry type is compared by the differ the
    // day it exists and was silently ignored by the summary until somebody edited its list. Every
    // readable property of every row type is one of the summary's fields.
    // ###########################################################################################
    [Theory]
    [InlineData(typeof(BoardSchematicEntry))]
    [InlineData(typeof(ComponentEntry))]
    [InlineData(typeof(ComponentImageEntry))]
    [InlineData(typeof(ComponentHighlightEntry))]
    [InlineData(typeof(ComponentLocalFileEntry))]
    [InlineData(typeof(ComponentLinkEntry))]
    [InlineData(typeof(BoardLocalFileEntry))]
    [InlineData(typeof(BoardLinkEntry))]
    [InlineData(typeof(CreditEntry))]
    [InlineData(typeof(KiCadImportantSignalEntry))]
    [InlineData(typeof(KiCadCalibrationEntry))]
    [InlineData(typeof(ComponentPinEntry))]
    [InlineData(typeof(ComponentReplacementEntry))]
    public void The_summary_compares_every_property_of_a_row(Type rowType)
    {
        object row = Activator.CreateInstance(rowType)!;
        int properties = rowType.GetProperties(BindingFlags.Public | BindingFlags.Instance).Count(property => property.CanRead);

        Assert.Equal(properties, ReviewSummary.FieldsOf(row).Count);
    }
}
