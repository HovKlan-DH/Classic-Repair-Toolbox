using CRT;

namespace ClassicRepairToolbox.Tests;

// Main.FormatBoardKeyAsFullNames - the hardware and board names shown on a workbook card in the
// Workbooks tab, on two lines ("Commodore 128" above "310378 (C128 & C128D)").
//
// WHAT THIS COVERS AND WHAT IT CANNOT. The method prefers the entry resolved from
// DataManager.HardwareBoards, whose setter is private, so a test cannot populate it - that branch is
// verified by running the app. What IS covered is the FALLBACK, which splits the BoardKey itself,
// and that is the branch worth pinning: a workbook whose board was later removed from
// classic-repair-toolbox.dk resolves no entry, and before this method existed such a card fell back
// to showing a raw key with a pipe in it.
//
// The key's shape is "HardwareName|BoardName" - see Main.FindEntryForBoardKey, which builds exactly
// that string to match against.
//
// Referenced as CRT.Main rather than Main: the app also has a "Main" NAMESPACE (the Main/ folder),
// which shadows the class name here.
//
// No collection attribute: this touches no static state (DataManager.HardwareBoards is only READ,
// and an empty list is precisely the no-entry case these tests drive).
public sealed class BoardKeyFullNameTests
{
    // The fallback, and the shape every real board key has. Both halves come back separately,
    // because the card renders them on their own lines.
    [Fact]
    public void A_board_key_is_split_into_its_hardware_and_board_halves()
    {
        var (hardware, board) = CRT.Main.FormatBoardKeyAsFullNames("Commodore 128|310378 (C128 & C128D)");

        Assert.Equal("Commodore 128", hardware);
        Assert.Equal("310378 (C128 & C128D)", board);
    }

    // The board half may itself contain spaces, brackets and ampersands - "310378 (C128 & C128D)" is
    // a real board name. Only the FIRST pipe separates, so nothing in the board name can be
    // mistaken for the separator.
    [Fact]
    public void Only_the_first_separator_splits_the_key()
    {
        var (hardware, board) = CRT.Main.FormatBoardKeyAsFullNames("Commodore 64|250469 (short board)");

        Assert.Equal("Commodore 64", hardware);
        Assert.Equal("250469 (short board)", board);
    }

    // A key with no separator at all (blank, or hand-edited into the workbook's index.json) must
    // still render something rather than blanking the card's board line. It goes on the first line
    // and the second is omitted by the caller.
    [Fact]
    public void A_key_with_no_separator_is_shown_whole_on_the_first_line()
    {
        var (hardware, board) = CRT.Main.FormatBoardKeyAsFullNames("SomethingUnexpected");

        Assert.Equal("SomethingUnexpected", hardware);
        Assert.Equal(string.Empty, board);
    }

    // A malformed key must not throw or produce a half-empty pair that renders as a stray line. A
    // separator at either end leaves one side empty, so the whole string is kept together instead.
    [Theory]
    [InlineData("|310378")]
    [InlineData("Commodore 128|")]
    public void A_key_with_an_empty_half_is_not_split(string boardKey)
    {
        var (hardware, board) = CRT.Main.FormatBoardKeyAsFullNames(boardKey);

        Assert.Equal(boardKey, hardware);
        Assert.Equal(string.Empty, board);
    }

    [Fact]
    public void A_blank_or_null_key_yields_blanks_rather_than_throwing()
    {
        Assert.Equal((string.Empty, string.Empty), CRT.Main.FormatBoardKeyAsFullNames(string.Empty));
        Assert.Equal((string.Empty, string.Empty), CRT.Main.FormatBoardKeyAsFullNames(null!));
    }

    // The point of the change: what a card shows is now the names the two drop-downs use, NOT the
    // short folder-derived "C128/310378" that FormatBoardKeyForDisplay produces for the worklog
    // picker. Both formatters still exist on purpose - the short one is right in a dropdown row
    // where several boards sit side by side.
    [Fact]
    public void The_card_names_are_not_the_short_folder_derived_label()
    {
        var (hardware, board) = CRT.Main.FormatBoardKeyAsFullNames("Commodore 128|310378 (C128 & C128D)");

        Assert.DoesNotContain("/", hardware);
        Assert.DoesNotContain("/", board);
        Assert.DoesNotContain("|", hardware);
        Assert.DoesNotContain("|", board);
    }
}
