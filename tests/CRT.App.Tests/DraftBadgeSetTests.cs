using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftBadgeSet - which Hardware and Board drop-down entries carry the "Draft" chip (maintainer
// request, 2026-09-24). A board carries it when it has a local draft; a hardware when ANY of its
// boards does.
// ###########################################################################################
public sealed class DraftBadgeSetTests
{
    private static HardwareBoardEntry Entry(string hardware, string board) => new()
    {
        HardwareName = hardware,
        BoardName = board,
        ExcelDataFile = $"Commodore/{hardware}/{board}/Data {hardware} {board}.xlsx",
    };

    [Fact]
    public void A_drafted_board_and_its_hardware_both_carry_the_chip()
    {
        var badges = DraftBadgeSet.From(new[] { Entry("C64", "250407") });

        Assert.True(badges.HardwareHasDraft("C64"));
        Assert.True(badges.BoardHasDraft("C64", "250407"));
    }

    // One drafted board is enough to mark its hardware - but only that board, not its siblings.
    [Fact]
    public void Only_the_DRAFTED_board_of_a_hardware_carries_the_chip()
    {
        var badges = DraftBadgeSet.From(new[] { Entry("C64", "250407") });

        Assert.False(badges.BoardHasDraft("C64", "250425"));
    }

    [Fact]
    public void A_hardware_with_no_drafted_board_carries_no_chip()
    {
        var badges = DraftBadgeSet.From(new[] { Entry("C64", "250407") });

        Assert.False(badges.HardwareHasDraft("VIC-20"));
    }

    // ###########################################################################################
    // A board name is not unique across hardware - so a draft of C64 "250407" must not badge a
    // board of the same name under another hardware. This is why the question is always asked
    // about a board OF a hardware.
    // ###########################################################################################
    [Fact]
    public void The_same_board_name_under_ANOTHER_hardware_carries_no_chip()
    {
        var badges = DraftBadgeSet.From(new[] { Entry("C64", "250407") });

        Assert.False(badges.BoardHasDraft("C128", "250407"));
    }

    // The drop-downs match names with OrdinalIgnoreCase, so the chip must too.
    [Fact]
    public void Names_match_ignoring_case_and_surrounding_whitespace()
    {
        var badges = DraftBadgeSet.From(new[] { Entry(" Commodore 64 ", "250407") });

        Assert.True(badges.HardwareHasDraft("commodore 64"));
        Assert.True(badges.BoardHasDraft("COMMODORE 64", " 250407 "));
    }

    [Fact]
    public void No_drafts_means_no_chips_anywhere()
    {
        Assert.False(DraftBadgeSet.Empty.HardwareHasDraft("C64"));
        Assert.False(DraftBadgeSet.From(null).BoardHasDraft("C64", "250407"));
        Assert.False(DraftBadgeSet.From(Array.Empty<HardwareBoardEntry>()).HardwareHasDraft("C64"));
    }

    [Fact]
    public void Blank_names_are_never_badged()
    {
        var badges = DraftBadgeSet.From(new[] { Entry("", "250407"), Entry("C64", "") });

        Assert.False(badges.HardwareHasDraft(string.Empty));
        Assert.False(badges.HardwareHasDraft(null));
        Assert.False(badges.BoardHasDraft("C64", string.Empty));

        // The entry with a blank board still marks its hardware - it IS a drafted system of it.
        Assert.True(badges.HardwareHasDraft("C64"));
    }
}
