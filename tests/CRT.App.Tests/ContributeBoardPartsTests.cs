using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributeBoardParts - what the Contribute tab's left column says about a board (CRT 4.0 step 2,
// 2026-10-25 - the plan's mock: "Pictures 12 ... Important signals 4 ... Files and links 6 ...
// Credits 3"), and its lines for when there is nothing to list.
// ###########################################################################################
public sealed class ContributeBoardPartsTests
{
    [Fact]
    public void A_boards_parts_are_counted_files_and_links_together()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top" }, new BoardSchematicEntry { SchematicName = "Bottom" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "CLK", KiCadNetName = "CLK" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Name = "Service manual" }],
            BoardLinks = [new BoardLinkEntry { Name = "Forum" }, new BoardLinkEntry { Name = "Video" }],
            Credits = [new CreditEntry { NameOrHandle = "Anna" }],
        };

        Assert.Equal(new ContributeBoardParts(2, 1, 3, 1), ContributeBoardParts.Of(board));
    }

    [Fact]
    public void No_board_has_none_of_anything()
    {
        Assert.Equal(new ContributeBoardParts(0, 0, 0, 0), ContributeBoardParts.Of(null));
    }

    // "none", not a bare 0, which reads like a fault.
    [Theory]
    [InlineData(0, "none")]
    [InlineData(-1, "none")]
    [InlineData(1, "1")]
    [InlineData(12, "12")]
    public void A_count_reads_as_the_column_shows_it(int count, string expected)
    {
        Assert.Equal(expected, ContributeBoardParts.CountText(count));
    }

    [Theory]
    [InlineData(0, "No drafts yet")]
    [InlineData(1, "My drafts (1)")]
    [InlineData(3, "My drafts (3)")]
    public void My_drafts_says_how_many_there_are(int drafts, string expected)
    {
        Assert.Equal(expected, ContributeBoardParts.MyDraftsLine(drafts));
    }

    // An empty list says why, and what adds the first one (the plan: "Every empty area says what goes
    // there and carries the button that adds it").
    [Fact]
    public void An_empty_component_list_says_why()
    {
        Assert.Equal("No board is shown.", ContributeBoardParts.NoComponentsLine(hasBoard: false, searching: false));
        Assert.Equal("No component matches the search.", ContributeBoardParts.NoComponentsLine(hasBoard: true, searching: true));
        Assert.Contains("\"Add component\"", ContributeBoardParts.NoComponentsLine(hasBoard: true, searching: false), StringComparison.Ordinal);
    }

    // The header of a board with no draft says what a change does - and that nothing is shared.
    [Fact]
    public void A_board_with_no_draft_says_what_the_first_change_does()
    {
        string line = ContributeBoardParts.NoDraftLine(hasBoard: true);

        Assert.Contains("first change makes your own draft", line, StringComparison.Ordinal);
        Assert.Contains("nothing is shared until you submit", line, StringComparison.Ordinal);
        Assert.Contains("Pick a hardware and a board", ContributeBoardParts.NoDraftLine(hasBoard: false), StringComparison.Ordinal);
    }
}
