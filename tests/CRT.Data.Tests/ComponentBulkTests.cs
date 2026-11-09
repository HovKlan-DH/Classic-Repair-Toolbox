using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentBulk - several components edited as one on the "Contribute" tab (owner request, 2026-11-08: select "R21" and
// "R22", change "Category" or "Friendly name" and both get it; on "Replacement parts" only the lines all of them have).
// A change reaches every component picked and nothing else, the rows only some have stay as they are, and a change that
// changes nothing is no change - the same board back, no step in the draft's history.
// ###########################################################################################
public sealed class ComponentBulkTests
{
    private static readonly string[] Both = ["R21", "R22"];

    private static BoardData Board() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "R21", Category = "Resistor", FriendlyName = "Pull-up", TechnicalNameOrValue = "4K7" },
            new ComponentEntry { BoardLabel = "R22", Category = "Resistor", FriendlyName = "Pull-down", TechnicalNameOrValue = "4K7" },
            new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" },
        ],
        ComponentReplacements =
        [
            new ComponentReplacementEntry { BoardLabel = "R21", PartNumber = "MF 1%", Kind = "Drop-in" },
            new ComponentReplacementEntry { BoardLabel = "R21", PartNumber = "Only R21" },
            new ComponentReplacementEntry { BoardLabel = "R22", PartNumber = "Only R22" },
            new ComponentReplacementEntry { BoardLabel = "R22", PartNumber = "MF 1%", Kind = "Drop-in" },
            new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "MF 1%", Kind = "Drop-in" },
        ],
        ComponentLinks =
        [
            new ComponentLinkEntry { BoardLabel = "R21", Name = "Colors", Url = "https://example.org/colors" },
            new ComponentLinkEntry { BoardLabel = "R22", Name = "Colors", Url = "https://example.org/colors" },
        ],
    };

    private static List<string> PartsOf(BoardData board, string label) =>
        ComponentPaneEdits.ReplacementsOf(board, label, string.Empty).Select(row => row.PartNumber).ToList();

    // ---- fields -------------------------------------------------------------------------------------

    [Fact]
    public void A_field_they_share_is_shown_and_one_they_differ_in_is_none()
    {
        BoardData board = Board();

        Assert.Equal("4K7", ComponentBulk.CommonValue(board, Both, string.Empty, ComponentField.TechnicalNameOrValue));
        Assert.Equal("Resistor", ComponentBulk.CommonValue(board, Both, string.Empty, ComponentField.Category));
        Assert.Null(ComponentBulk.CommonValue(board, Both, string.Empty, ComponentField.FriendlyName));
        Assert.Equal(["Pull-up", "Pull-down"], ComponentBulk.ValuesOf(board, Both, string.Empty, ComponentField.FriendlyName));
    }

    // ###########################################################################################
    // VALUES ARE COMPARED AS AN EDIT KEEPS THEM (code review, 2026-11-09): a space left at the end by Excel, or a category in
    // other capitals - one category everywhere else in CRT - read as two values, the box empty and its list naming the same
    // text twice. Values that truly differ still do.
    // ###########################################################################################
    [Fact]
    public void Values_that_read_the_same_are_one_value()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "R21", Category = "IC", Description = "10k 1%", FriendlyName = "A" },
                new ComponentEntry { BoardLabel = "R22", Category = "ic", Description = "10k 1% ", FriendlyName = "a" },
            ],
        };

        Assert.Equal("IC", ComponentBulk.CommonValue(board, Both, string.Empty, ComponentField.Category));
        Assert.Equal("10k 1%", ComponentBulk.CommonValue(board, Both, string.Empty, ComponentField.Description));
        Assert.Null(ComponentBulk.CommonValue(board, Both, string.Empty, ComponentField.FriendlyName));
    }

    [Fact]
    public void A_field_typed_is_set_on_every_one_picked_and_no_other()
    {
        BoardData board = ComponentBulk.WithField(Board(), Both, string.Empty, ComponentField.FriendlyName, "Bus pull");

        Assert.Equal(["Bus pull", "Bus pull", "CPU"], board.Components.Select(row => row.FriendlyName));
        Assert.Equal(["R21", "R22", "U8"], board.Components.Select(row => row.BoardLabel));
    }

    [Fact]
    public void A_field_set_to_what_they_all_have_is_no_change()
    {
        BoardData board = Board();

        Assert.Same(board, ComponentBulk.WithField(board, Both, string.Empty, ComponentField.Category, "Resistor"));
    }

    [Fact]
    public void Make_all_like_one_gives_the_others_its_fields_but_never_its_label()
    {
        BoardData board = ComponentBulk.FieldsMadeLike(Board(), "R21", Both, string.Empty);

        ComponentEntry r22 = board.Components.Single(row => row.BoardLabel == "R22");
        Assert.Equal("Pull-up", r22.FriendlyName);
        Assert.Equal("R22", r22.BoardLabel);
        Assert.Equal("CPU", board.Components.Single(row => row.BoardLabel == "U8").FriendlyName);
    }

    // "Basics" made like one leaves the package, which "Pinout" shows (code review, 2026-11-09: it went too, the note under
    // the button saying only "Basics" changes).
    [Fact]
    public void Make_all_like_one_on_Basics_leaves_the_package()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "R21", Category = "Resistor", FriendlyName = "Pull-up", Package = "Axial" },
                new ComponentEntry { BoardLabel = "R22", Category = "Resistor", FriendlyName = "Pull-down", Package = "Radial" },
            ],
        };

        ComponentEntry r22 = ComponentBulk.FieldsMadeLike(board, "R21", Both, string.Empty).Components.Single(row => row.BoardLabel == "R22");

        Assert.Equal("Pull-up", r22.FriendlyName);
        Assert.Equal("Radial", r22.Package);
        Assert.DoesNotContain(ComponentField.Package, ComponentBulk.BasicsFields);
    }

    // ---- lists --------------------------------------------------------------------------------------

    [Fact]
    public void The_common_rows_are_those_every_one_has_by_what_they_hold_not_their_label()
    {
        BoardData board = Board();

        Assert.Equal(["MF 1%"], ComponentBulk.CommonRows(board, Both, string.Empty, ComponentBulk.Replacements).Select(row => row.PartNumber));
        Assert.Equal(
            [("R21", "Only R21"), ("R22", "Only R22")],
            ComponentBulk.RowsOfSome(board, Both, string.Empty, ComponentBulk.Replacements).Select(each => (each.Label, each.Row.PartNumber)));
        Assert.Empty(ComponentBulk.RowsOfSome(board, Both, string.Empty, ComponentBulk.Links));
        Assert.NotEmpty(ComponentBulk.RowsOfSome(board, Both, string.Empty, ComponentBulk.Replacements));
    }

    // A row one has twice is common twice only where every other has it twice too.
    [Fact]
    public void A_row_twice_is_common_as_often_as_every_one_has_it()
    {
        BoardData board = Board().WithComponentSections(componentLinks:
        [
            new ComponentLinkEntry { BoardLabel = "R21", Name = "Colors", Url = "https://example.org/colors" },
            new ComponentLinkEntry { BoardLabel = "R21", Name = "Colors", Url = "https://example.org/colors" },
            new ComponentLinkEntry { BoardLabel = "R22", Name = "Colors", Url = "https://example.org/colors" },
        ]);

        Assert.Single(ComponentBulk.CommonRows(board, Both, string.Empty, ComponentBulk.Links));
        Assert.Equal("R21", Assert.Single(ComponentBulk.RowsOfSome(board, Both, string.Empty, ComponentBulk.Links)).Label);
    }

    [Fact]
    public void A_common_row_changed_changes_on_every_one_and_their_own_rows_stay_where_they_are()
    {
        BoardData board = Board();
        List<ComponentReplacementEntry> before = ComponentBulk.CommonRows(board, Both, string.Empty, ComponentBulk.Replacements);

        var changed = new ComponentReplacementEntry { BoardLabel = "R21", PartNumber = "MF 0.5%", Kind = "Drop-in" };
        BoardData after = ComponentBulk.WithCommonRows(board, Both, string.Empty, ComponentBulk.Replacements, before, [(changed, 0)]);

        Assert.Equal(["MF 0.5%", "Only R21"], PartsOf(after, "R21"));
        Assert.Equal(["Only R22", "MF 0.5%"], PartsOf(after, "R22"));
        Assert.Equal(["MF 1%"], PartsOf(after, "U8"));
    }

    [Fact]
    public void A_common_row_removed_goes_from_every_one_and_a_row_added_comes_to_every_one()
    {
        BoardData board = Board();
        List<ComponentReplacementEntry> before = ComponentBulk.CommonRows(board, Both, string.Empty, ComponentBulk.Replacements);

        var added = new ComponentReplacementEntry { PartNumber = "CF 5%" };
        BoardData after = ComponentBulk.WithCommonRows(board, Both, string.Empty, ComponentBulk.Replacements, before, [(added, -1)]);

        Assert.Equal(["Only R21", "CF 5%"], PartsOf(after, "R21"));
        Assert.Equal(["Only R22", "CF 5%"], PartsOf(after, "R22"));
        Assert.Equal(["R21", "R22"], after.ComponentReplacements.Where(row => row.PartNumber == "CF 5%").Select(row => row.BoardLabel));
    }

    [Fact]
    public void The_common_rows_handed_back_as_they_were_are_no_change()
    {
        BoardData board = Board();
        List<ComponentReplacementEntry> before = ComponentBulk.CommonRows(board, Both, string.Empty, ComponentBulk.Replacements);

        Assert.Same(board, ComponentBulk.WithCommonRows(board, Both, string.Empty, ComponentBulk.Replacements, before, before.Select((row, at) => (row, at)).ToList()));
    }

    // In a region the rows of the other region stay as they are, on every one.
    [Fact]
    public void An_edit_in_a_region_leaves_the_other_regions_rows_alone()
    {
        BoardData board = new()
        {
            Regions = ["PAL", "NTSC"],
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC" },
            ],
            ComponentReplacements =
            [
                new ComponentReplacementEntry { BoardLabel = "U1", Region = "NTSC", PartNumber = "NTSC part" },
                new ComponentReplacementEntry { BoardLabel = "U2", Region = "NTSC", PartNumber = "NTSC part" },
            ],
        };

        BoardData after = ComponentBulk.WithCommonRows(board, ["U1", "U2"], "PAL", ComponentBulk.Replacements, [], [(new ComponentReplacementEntry { PartNumber = "Any" }, -1)]);

        Assert.Equal(2, after.ComponentReplacements.Count(row => row.PartNumber == "NTSC part"));
        Assert.Equal(2, after.ComponentReplacements.Count(row => row.PartNumber == "Any"));
    }

    // ###########################################################################################
    // A PIN ADDED TO SEVERAL JOINS EACH ONE'S OWN SET (code review, 2026-11-09): the first one's region went to all - U5's
    // "PAL" (a component of PAL alone) gave U6, whose pins are for every region, a pin NTSC never showed.
    // ###########################################################################################
    [Fact]
    public void A_pin_added_to_several_joins_each_ones_own_set()
    {
        var board = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components =
            [
                new ComponentEntry { BoardLabel = "U5", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U6", Category = "IC" },
                new ComponentEntry { BoardLabel = "U7", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U7", Category = "IC", Region = "NTSC" },
            ],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "A", Region = "PAL" },
                new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "A", Region = "NTSC" },
            ],
        };

        BoardData after = ComponentBulk.WithCommonRows(
            board,
            ["U5", "U6", "U7"],
            "PAL",
            ComponentBulk.Pins,
            [],
            [(new ComponentPinEntry { Pin = "9", Name = "CLK", Region = "PAL" }, -1)]);

        string RegionOf(string label) => after.ComponentPins.Single(pin => pin.BoardLabel == label && pin.Pin == "9").Region;

        Assert.Equal("PAL", RegionOf("U5"));
        Assert.Equal(string.Empty, RegionOf("U6"));
        Assert.Equal("PAL", RegionOf("U7"));

        // A component of one region keeps its pins in it whatever region CRT shows, its pins told apart or not (second code
        // review, 2026-11-09: NTSC-only B's new pin went to PAL with PAL shown - a region B is not in).
        BoardData ntscOnly = board.WithComponentSections(
            components: [.. board.Components, new ComponentEntry { BoardLabel = "B", Category = "IC", Region = "NTSC" }],
            componentPins: [.. board.ComponentPins, new ComponentPinEntry { BoardLabel = "B", Pin = "1", Name = "A", Region = "NTSC" }]);

        Assert.Equal("NTSC", ComponentPaneEdits.RegionForNewPin(ntscOnly, "B", "PAL"));
        Assert.Equal("PAL", ComponentPaneEdits.RegionForNewPin(ntscOnly, "U7", "PAL"));
        Assert.Equal(string.Empty, ComponentPaneEdits.RegionForNewPin(ntscOnly, "U6", "PAL"));
    }

    [Fact]
    public void Make_all_like_one_gives_the_others_its_rows_in_place_of_theirs()
    {
        BoardData board = ComponentBulk.MadeLike(Board(), "R21", Both, ComponentBulk.Replacements);

        Assert.Equal(["MF 1%", "Only R21"], PartsOf(board, "R22"));
        Assert.Equal(["MF 1%", "Only R21"], PartsOf(board, "R21"));
        Assert.Equal(["MF 1%"], PartsOf(board, "U8"));
        Assert.Empty(ComponentBulk.RowsOfSome(board, Both, string.Empty, ComponentBulk.Replacements));

        // Already alike: the same board, no step.
        Assert.Same(board, ComponentBulk.MadeLike(board, "R21", Both, ComponentBulk.Replacements));
    }

    // ###########################################################################################
    // "MAKE ALL LIKE" GIVES THE OTHERS THE PAGE IN EVERY REGION, AS THE ONE PICKED FILES IT (second code review, 2026-11-09):
    // read in the region shown and written over what the others showed there, U1's PAL pins took the place of U2's pins for
    // every region - U2 had none in NTSC any more - and the other way round NTSC listed U2's own and U1's.
    // ###########################################################################################
    [Fact]
    public void Make_all_like_one_gives_the_others_the_page_in_every_region_as_that_one_files_it()
    {
        static List<string> PinsIn(BoardData board, string label, string region) =>
            ComponentPaneEdits.PinsOf(board, label, region).Select(pin => $"{pin.Pin}:{pin.Name}").ToList();

        var board = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC" },
            ],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "PAL clock", Region = "PAL" },
                new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "NTSC clock", Region = "NTSC" },
                new ComponentPinEntry { BoardLabel = "U2", Pin = "1", Name = "Clock" },
            ],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U2", Pin = "1", Name = "Clock", File = "Scope baseline/u2.png" }],
        };

        // U1's pins told apart per region, to U2 whose pins are for every region: each region shows U1's.
        BoardData likeU1 = ComponentBulk.MadeLike(board, "U1", ["U1", "U2"], ComponentBulk.Pins);
        Assert.Equal(["1:PAL clock"], PinsIn(likeU1, "U2", "PAL"));
        Assert.Equal(["1:NTSC clock"], PinsIn(likeU1, "U2", "NTSC"));

        // ... and the other way round: U1's pins are U2's, for every region, nothing of NTSC's left beside them.
        BoardData likeU2 = ComponentBulk.MadeLike(board, "U2", ["U1", "U2"], ComponentBulk.Pins);
        Assert.Equal(["1:Clock"], PinsIn(likeU2, "U1", "PAL"));
        Assert.Equal(["1:Clock"], PinsIn(likeU2, "U1", "NTSC"));

        // Pictures leave the baselines - another page - as they are.
        BoardData pictures = ComponentBulk.MadeLike(board, "U1", ["U1", "U2"], ComponentBulk.Pictures);
        Assert.Single(pictures.ComponentImages, image => image.BoardLabel == "U2" && image.Pin == "1");
    }
}
