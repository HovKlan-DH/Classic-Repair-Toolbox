using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributeBoardParts - what the Contribute tab's left column says about a board (CRT 4.0 step 2,
// 2026-10-25 - the plan's mock: "Pictures 12 ... Important signals 4 ... Files and links 6 ...
// Credits 3"), and its lines for when there is nothing to list.
// ###########################################################################################
public sealed class ContributeBoardPartsTests
{
    // Files and links counted apart since 2026-10-29 - an entry each in the column (owner request).
    [Fact]
    public void A_boards_parts_are_counted_files_and_links_apart()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top" }, new BoardSchematicEntry { SchematicName = "Bottom" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "CLK", KiCadNetName = "CLK" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Name = "Service manual" }],
            BoardLinks = [new BoardLinkEntry { Name = "Forum" }, new BoardLinkEntry { Name = "Video" }],
            Credits = [new CreditEntry { NameOrHandle = "Anna" }],
        };

        Assert.Equal(new ContributeBoardParts(2, 1, 1, 2, 1), ContributeBoardParts.Of(board));
    }

    [Fact]
    public void No_board_has_none_of_anything()
    {
        Assert.Equal(new ContributeBoardParts(0, 0, 0, 0, 0), ContributeBoardParts.Of(null));
    }

    // The entries added on 2026-11-01: how many categories, and what the board comes in.
    [Fact]
    public void The_categories_are_counted_and_the_regions_named()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "NTSC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            ],
        };

        Assert.Equal(2, ContributeBoardParts.Of(board).Categories);
        Assert.Equal("PAL, NTSC", ContributeBoardParts.RegionsText(board));
        Assert.Equal("none", ContributeBoardParts.RegionsText(new BoardData()));
        Assert.Equal("none", ContributeBoardParts.RegionsText(null));
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

    // ###########################################################################################
    // "KiCad data" in the column counts the board's KiCad files (owner request, 2026-10-30) - the names
    // under "KiCad data/", by the rule Submit sends them by, each once however it is spelt; nothing
    // else the board lists.
    // ###########################################################################################
    [Fact]
    public void The_KiCad_files_counted_are_the_names_under_KiCad_data_each_once()
    {
        Assert.Equal(3, ContributeBoardParts.KiCadFileCount(
        [
            "KiCad data/board.kicad_pcb",
            "KiCad data/board.kicad_sch",
            "KiCad data/Pages/video.kicad_sch",
            "kicad data/board.kicad_pcb",
            "Board local files/manual.pdf",
            "Images/top.png",
        ]));

        Assert.Equal(0, ContributeBoardParts.KiCadFileCount([]));
        Assert.Equal("none", ContributeBoardParts.CountText(ContributeBoardParts.KiCadFileCount(["Images/top.png"])));
    }

    // ###########################################################################################
    // THE NOTICE OVER A DRAFT (owner request, 2026-10-30: "clearly stating that this is a local draft
    // and that it must be submitted for a board maintainer to review it, before others can see it") -
    // all three said, for an existing board and a new one, in plain ASCII.
    // ###########################################################################################
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void The_draft_notice_says_it_is_local_and_must_be_submitted_and_reviewed(bool isNewBoard)
    {
        string notice = ContributeBoardParts.LocalDraftNotice(isNewBoard);

        Assert.Contains("local draft", notice, StringComparison.Ordinal);
        Assert.Contains("only on this computer", notice, StringComparison.Ordinal);
        Assert.Contains("submit", notice, StringComparison.Ordinal);
        Assert.Contains("maintainer", notice, StringComparison.Ordinal);
        Assert.Contains("reviewed", notice, StringComparison.Ordinal);
        Assert.All(notice, character => Assert.True(character < 128, $"[{character}] is not plain ASCII"));
    }
}
