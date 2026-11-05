using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentClone - "Clone component from another board" (owner request, 2026-11-04: "It will not change the
// component name, but all other data should be identical to the one it copies from").
// ###########################################################################################
public sealed class ComponentCloneTests
{
    private static readonly IReadOnlyDictionary<string, string> NoFiles = new Dictionary<string, string>();

    // A chip that has something of everything, one row for every region.
    private static BoardData Source() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U7", FriendlyName = "PLA", TechnicalNameOrValue = "906114", PartNumber = "906114-01", Category = "Chips", Package = "DIP-28", Description = "Glue logic" },
            new ComponentEntry { BoardLabel = "U9", FriendlyName = "Other" },
        ],
        ComponentPins = [new ComponentPinEntry { BoardLabel = "U7", Pin = "28", Name = "VCC", Direction = "Power" }],
        ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U7", PartNumber = "PLAnkton", Kind = "Drop-in replacement", Note = "Socket it" }],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U7", Name = "Datasheet", Url = "https://example.org/pla.pdf" }],
        ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Notes", File = "Component local files/pla.txt" }],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U7", Name = "Top side", File = "Images/pla.png", Note = "As fitted" },
            new ComponentImageEntry { BoardLabel = "U7", Pin = "1", Name = "I0", File = "Scope baseline/i0.png", TimeDiv = "1us", VoltsDiv = "1V", TriggerLevelVolts = "2", MachineState = "READY" },
        ],
    };

    private static BoardData Target() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" },
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "Old name", PartNumber = "OLD", Category = "IC", Package = "DIP-40", Description = "Old" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "OLD" },
            new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "VSS" },
        ],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Old link", Url = "https://example.org/old" }],
        ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
    };

    private static Dictionary<string, string> SameNames(params string[] names) =>
        names.ToDictionary(name => name, name => name, StringComparer.OrdinalIgnoreCase);

    private static BoardData Cloned(BoardData? target = null, BoardData? source = null, string shown = "PAL") =>
        ComponentClone.Apply(target ?? Target(), "U8", source ?? Source(), "U7", shown, SameNames("Component local files/pla.txt", "Images/pla.png", "Scope baseline/i0.png"));

    [Fact]
    public void Every_field_and_every_list_comes_under_this_components_label()
    {
        BoardData board = Cloned();

        ComponentEntry row = Assert.Single(board.Components, entry => entry.BoardLabel == "U8");
        Assert.Equal("PLA", row.FriendlyName);
        Assert.Equal("906114", row.TechnicalNameOrValue);
        Assert.Equal("906114-01", row.PartNumber);
        Assert.Equal("Chips", row.Category);
        Assert.Equal("DIP-28", row.Package);
        Assert.Equal("Glue logic", row.Description);

        Assert.Equal("VCC", Assert.Single(board.ComponentPins, pin => pin.BoardLabel == "U8").Name);
        Assert.Equal("Socket it", Assert.Single(board.ComponentReplacements, part => part.BoardLabel == "U8").Note);
        Assert.Equal("Datasheet", Assert.Single(board.ComponentLinks, link => link.BoardLabel == "U8").Name);
        Assert.Equal("Component local files/pla.txt", Assert.Single(board.ComponentLocalFiles, file => file.BoardLabel == "U8").File);

        // A picture and a baseline, every value of each.
        Assert.Equal(2, board.ComponentImages.Count(image => image.BoardLabel == "U8"));
        Assert.Equal("As fitted", board.ComponentImages.Single(image => image.File == "Images/pla.png").Note);
        Assert.Equal("READY", board.ComponentImages.Single(image => image.Pin == "1").MachineState);
    }

    // What it held of those goes - and no other component, its label or where it is marked moves.
    [Fact]
    public void What_it_held_goes_while_its_label_its_marks_and_the_other_components_stay()
    {
        BoardData board = Cloned();

        Assert.Equal(["U1", "U8"], board.Components.Select(row => row.BoardLabel));
        Assert.DoesNotContain(board.ComponentPins, pin => pin.Name == "OLD");
        Assert.DoesNotContain(board.ComponentLinks, link => link.Name == "Old link");

        Assert.Equal("VSS", Assert.Single(board.ComponentPins, pin => pin.BoardLabel == "U1").Name);
        Assert.Equal("CPU", board.Components.Single(row => row.BoardLabel == "U1").FriendlyName);

        ComponentHighlightEntry mark = Assert.Single(board.ComponentHighlights);
        Assert.Equal("U8", mark.BoardLabel);
        Assert.Equal("1", mark.X);
    }

    // Nothing new: no step for the draft's history.
    [Fact]
    public void Cloned_again_it_is_the_same_board()
    {
        BoardData once = Cloned();

        Assert.Same(once, Cloned(once));
    }

    [Fact]
    public void A_component_either_board_does_not_have_changes_nothing()
    {
        BoardData target = Target();

        Assert.Same(target, ComponentClone.Apply(target, "U99", Source(), "U7", "PAL", NoFiles));
        Assert.Same(target, ComponentClone.Apply(target, "U8", Source(), "U99", "PAL", NoFiles));
        Assert.Null(ComponentClone.FactsOf(target, "U8", Source(), "U99", "PAL"));
        Assert.Empty(ComponentClone.FilesCited(target, "U99", Source(), "U7", "PAL"));
    }

    // ---- regions -----------------------------------------------------------------------------------

    // A chip told apart per region, its rows saying so - pins and files too.
    private static BoardData Regional() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U7", Region = "PAL", FriendlyName = "VIC PAL", Category = "IC" },
            new ComponentEntry { BoardLabel = "U7", Region = "NTSC", FriendlyName = "VIC NTSC", Category = "IC" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U7", Region = "PAL", Pin = "1", Name = "P" },
            new ComponentPinEntry { BoardLabel = "U7", Region = "NTSC", Pin = "1", Name = "N" },
            new ComponentPinEntry { BoardLabel = "U7", Pin = "40", Name = "VCC" },
        ],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U7", Region = "NTSC", Name = "NTSC notes", Url = "https://example.org/n" }],
    };

    private static BoardData WithRegions(BoardData board) => board.WithRegions(["PAL", "NTSC"]);

    [Fact]
    public void In_every_region_on_both_boards_its_rows_come_as_they_are()
    {
        BoardData board = ComponentClone.Apply(WithRegions(Target()), "U8", Regional(), "U7", "PAL", NoFiles);

        Assert.Equal(["PAL:VIC PAL", "NTSC:VIC NTSC"], board.Components.Where(row => row.BoardLabel == "U8").Select(row => $"{row.Region}:{row.FriendlyName}"));
        Assert.Equal([":VCC", "NTSC:N", "PAL:P"], board.ComponentPins.Where(pin => pin.BoardLabel == "U8").Select(pin => $"{pin.Region}:{pin.Name}").Order(StringComparer.Ordinal));
        Assert.Equal("NTSC", Assert.Single(board.ComponentLinks, link => link.BoardLabel == "U8").Region);
    }

    // A PAL-only chip there must not take this component out of NTSC, where it stands too.
    [Fact]
    public void Of_one_region_there_and_every_region_here_it_comes_for_every_region()
    {
        var source = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U7", Region = "PAL", FriendlyName = "PAL chip" }],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U7", Region = "PAL", Pin = "1", Name = "A" },
                new ComponentPinEntry { BoardLabel = "U7", Region = "NTSC", Pin = "1", Name = "Stray" },
            ],
        };

        BoardData board = ComponentClone.Apply(WithRegions(Target()), "U8", source, "U7", "NTSC", NoFiles);

        Assert.Equal(string.Empty, Assert.Single(board.Components, row => row.BoardLabel == "U8").Region);

        ComponentPinEntry pin = Assert.Single(board.ComponentPins, row => row.BoardLabel == "U8");
        Assert.Equal("A", pin.Name);
        Assert.Equal(string.Empty, pin.Region);
    }

    [Fact]
    public void Of_one_region_here_it_takes_what_that_region_shows_there_and_stays_in_it()
    {
        BoardData target = WithRegions(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Region = "NTSC", FriendlyName = "Mine" }],
        });

        BoardData board = ComponentClone.Apply(target, "U8", Regional(), "U7", "PAL", NoFiles);

        ComponentEntry row = Assert.Single(board.Components);
        Assert.Equal("NTSC", row.Region);
        Assert.Equal("VIC NTSC", row.FriendlyName);

        // Its own region's stays that region's, and what was for every region stays so.
        Assert.Equal(["NTSC:N", ":VCC"], board.ComponentPins.Select(pin => $"{pin.Region}:{pin.Name}"));
    }

    // ###########################################################################################
    // A plain chip there, its row for every region, must not take this component into the other region
    // (code review, 2026-11-04): U8 of NTSC would be listed and shown under PAL too. Its pins and the rest
    // for every region stay so.
    // ###########################################################################################
    [Fact]
    public void Of_one_region_here_a_chip_for_every_region_there_leaves_it_in_its_region()
    {
        BoardData target = WithRegions(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Region = "NTSC", FriendlyName = "Mine" }],
        });

        BoardData board = ComponentClone.Apply(target, "U8", Source(), "U7", "PAL", SameNames("Component local files/pla.txt", "Images/pla.png", "Scope baseline/i0.png"));

        ComponentEntry row = Assert.Single(board.Components);
        Assert.Equal("NTSC", row.Region);
        Assert.Equal("PLA", row.FriendlyName);

        Assert.Equal(string.Empty, Assert.Single(board.ComponentPins).Region);
        Assert.Equal(string.Empty, Assert.Single(board.ComponentLinks).Region);

        // Nothing new the second time, so no step.
        Assert.Same(board, ComponentClone.Apply(board, "U8", Source(), "U7", "PAL", SameNames("Component local files/pla.txt", "Images/pla.png", "Scope baseline/i0.png")));
    }

    // Rows of both kinds there - for every region, and of this one - are one row here, the one that region shows.
    [Fact]
    public void Of_one_region_here_it_never_gets_two_rows()
    {
        BoardData target = WithRegions(new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Region = "PAL", FriendlyName = "Mine" }],
        });

        var source = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U7", FriendlyName = "Every region" },
                new ComponentEntry { BoardLabel = "U7", Region = "PAL", FriendlyName = "PAL" },
            ],
        };

        ComponentEntry row = Assert.Single(ComponentClone.Apply(target, "U8", source, "U7", "PAL", NoFiles).Components);

        Assert.Equal("PAL", row.Region);
        Assert.Equal(ComponentPaneEdits.RowShown(source, "U7", "PAL")!.FriendlyName, row.FriendlyName);
    }

    // On a board with no regions too: rows for every region and of the region shown there are one row here, or
    // two rows with the same label and no region, a duplicate warning (code review, 2026-11-04).
    [Fact]
    public void On_a_board_with_no_regions_it_never_gets_two_rows()
    {
        Assert.False(BoardRegions.Has(Target()));

        var source = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U7", FriendlyName = "Every region" },
                new ComponentEntry { BoardLabel = "U7", Region = "PAL", FriendlyName = "PAL" },
            ],
        };

        ComponentEntry row = Assert.Single(ComponentClone.Apply(Target(), "U8", source, "U7", "PAL", NoFiles).Components, entry => entry.BoardLabel == "U8");

        Assert.Equal(string.Empty, row.Region);
        Assert.Equal(ComponentPaneEdits.RowShown(source, "U7", "PAL")!.FriendlyName, row.FriendlyName);
    }

    // A row saying a region would switch this board's PAL/NTSC buttons on.
    [Fact]
    public void On_a_board_with_no_regions_nothing_says_a_region()
    {
        Assert.False(BoardRegions.Has(Target()));

        BoardData board = ComponentClone.Apply(Target(), "U8", Regional(), "U7", "NTSC", NoFiles);

        Assert.False(BoardRegions.Has(board));
        Assert.Equal("VIC NTSC", Assert.Single(board.Components, row => row.BoardLabel == "U8").FriendlyName);
        Assert.Equal(["N", "VCC"], board.ComponentPins.Where(pin => pin.BoardLabel == "U8").Select(pin => pin.Name));
        Assert.Equal("NTSC notes", Assert.Single(board.ComponentLinks, link => link.BoardLabel == "U8").Name);
    }

    // ---- files ---------------------------------------------------------------------------------------

    [Fact]
    public void A_file_stands_under_the_name_it_was_given_here()
    {
        var names = new Dictionary<string, string>(SameNames("Component local files/pla.txt", "Scope baseline/i0.png"), StringComparer.OrdinalIgnoreCase)
        {
            ["Images/pla.png"] = "Images/pla (2).png",
        };

        BoardData board = ComponentClone.Apply(Target(), "U8", Source(), "U7", "PAL", names);

        Assert.Contains(board.ComponentImages, image => image.File == "Images/pla (2).png");
        Assert.DoesNotContain(board.ComponentImages, image => image.File == "Images/pla.png");
    }

    // Left in, the row would cite a file the draft does not have - an error at Submit, with no way to it.
    [Fact]
    public void A_row_citing_a_file_that_did_not_come_is_left_out_and_a_row_with_no_file_is_not()
    {
        BoardData source = Source().WithComponentSections(
            componentImages: [.. Source().ComponentImages, new ComponentImageEntry { BoardLabel = "U7", Name = "A note", Note = "No file here" }]);

        BoardData board = ComponentClone.Apply(Target(), "U8", source, "U7", "PAL", SameNames("Images/pla.png"));

        Assert.Equal(["Top side", "A note"], board.ComponentImages.Where(image => image.BoardLabel == "U8").Select(image => image.Name));
        Assert.Empty(board.ComponentLocalFiles);
    }

    [Fact]
    public void The_files_cited_are_each_named_once_as_the_taken_rows_cite_them()
    {
        BoardData source = Regional().WithComponentSections(
            componentImages:
            [
                new ComponentImageEntry { BoardLabel = "U7", Region = "PAL", File = "Images/./same.png" },
                new ComponentImageEntry { BoardLabel = "U7", Region = "PAL", File = "Images/same.png" },
                new ComponentImageEntry { BoardLabel = "U7", Region = "NTSC", File = "Images/ntsc.png" },
                new ComponentImageEntry { BoardLabel = "U9", File = "Images/another.png" },
            ]);

        Assert.Equal(["Images/same.png", "Images/ntsc.png"], ComponentClone.FilesCited(WithRegions(Target()), "U8", source, "U7", "PAL"));

        // On a board with no regions only the region shown comes, so only its file.
        Assert.Equal(["Images/same.png"], ComponentClone.FilesCited(Target(), "U8", source, "U7", "PAL"));
    }

    private static BoardDocument.StoredFileRow Listed(string name, string sha) => new() { Name = name, Sha256 = sha };

    // Under a name no file here has: a copy under a name in use would replace it for every row citing it.
    [Fact]
    public void A_file_comes_under_a_free_name_unless_the_very_same_file_is_here_already()
    {
        var there = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase)
        {
            ["Images/same.png"] = Listed("Images/same.png", "AAA"),
            ["Images/taken.png"] = Listed("Images/taken.png", "BBB"),
            ["Images/new.png"] = Listed("Images/new.png", "CCC"),
        };

        var here = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase)
        {
            ["Images/same.png"] = Listed("Images/same.png", "aaa"),
            ["Images/taken.png"] = Listed("Images/taken.png", "other bytes"),
        };

        List<ComponentCloneFile> files = ComponentClone.NamesHere(
            ["Images/same.png", "Images/taken.png", "Images/new.png"],
            there,
            here,
            ["Images/same.png", "Images/taken.png"]);

        Assert.Equal(new ComponentCloneFile("Images/same.png", "Images/same.png", CopyIn: false), files[0]);

        Assert.True(files[1].CopyIn);
        Assert.NotEqual("Images/taken.png", files[1].NameHere);
        Assert.EndsWith(".png", files[1].NameHere, StringComparison.Ordinal);

        Assert.Equal(new ComponentCloneFile("Images/new.png", "Images/new.png", CopyIn: true), files[2]);
    }

    // ###########################################################################################
    // A file that needed a name of its own the first time is found again by its bytes: cloned again, it is the
    // same file under the same name, so no copy and no step (code review, 2026-11-04: each clone picked a new
    // name, "Images/x (3).png", and said "Cloned" for no change).
    // ###########################################################################################
    [Fact]
    public void A_file_given_a_name_of_its_own_is_found_again_by_its_bytes_so_cloning_again_is_no_step()
    {
        var there = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase)
        {
            ["Images/x.png"] = Listed("Images/x.png", "AAA"),
        };

        var here = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase)
        {
            ["Images/x.png"] = Listed("Images/x.png", "other bytes"),
        };

        ComponentCloneFile first = Assert.Single(ComponentClone.NamesHere(["Images/x.png"], there, here, ["Images/x.png"]));

        Assert.True(first.CopyIn);
        Assert.NotEqual("Images/x.png", first.NameHere);

        here[first.NameHere] = Listed(first.NameHere, "aaa");

        Assert.Equal(
            new ComponentCloneFile("Images/x.png", first.NameHere, CopyIn: false),
            Assert.Single(ComponentClone.NamesHere(["Images/x.png"], there, here, ["Images/x.png", first.NameHere])));

        var source = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U7", FriendlyName = "PLA" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U7", Name = "Top", File = "Images/x.png" }],
        };

        var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["Images/x.png"] = first.NameHere };

        BoardData once = ComponentClone.Apply(Target(), "U8", source, "U7", "PAL", names);

        Assert.Equal(first.NameHere, Assert.Single(once.ComponentImages).File);
        Assert.Same(once, ComponentClone.Apply(once, "U8", source, "U7", "PAL", names));
    }

    // The folder says what kind of file it is: the same bytes filed as something else are not this file. Nor
    // does a file with no checksum to go by match another with none.
    [Fact]
    public void The_same_bytes_in_another_folder_or_no_checksum_at_all_are_copied_in()
    {
        var there = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase)
        {
            ["Images/x.png"] = Listed("Images/x.png", "AAA"),
            ["Images/none.png"] = Listed("Images/none.png", string.Empty),
        };

        var here = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase)
        {
            ["Scope baseline/x.png"] = Listed("Scope baseline/x.png", "AAA"),
            ["Images/blank.png"] = Listed("Images/blank.png", string.Empty),
        };

        List<ComponentCloneFile> files = ComponentClone.NamesHere(["Images/x.png", "Images/none.png"], there, here, here.Keys);

        Assert.All(files, file => Assert.True(file.CopyIn));
    }

    // ---- what the panel says ---------------------------------------------------------------------------

    [Fact]
    public void Its_facts_count_one_region_and_say_whether_the_regions_come()
    {
        ComponentCloneFacts every = ComponentClone.FactsOf(WithRegions(Target()), "U8", Regional(), "U7", "PAL")!;

        Assert.Equal("VIC PAL", every.Row.FriendlyName);
        Assert.Equal(2, every.Pins);
        Assert.True(every.EveryRegionComes);
        Assert.False(every.OtherRegionsLeftOut);

        ComponentCloneFacts one = ComponentClone.FactsOf(Target(), "U8", Regional(), "U7", "NTSC")!;

        Assert.Equal("NTSC", one.Region);
        Assert.Equal("VIC NTSC", one.Row.FriendlyName);
        Assert.False(one.EveryRegionComes);
        Assert.True(one.OtherRegionsLeftOut);
    }

    [Fact]
    public void Its_facts_tell_pictures_from_baselines_and_a_plain_component_says_nothing_of_regions()
    {
        ComponentCloneFacts facts = ComponentClone.FactsOf(Target(), "U8", Source(), "U7", "PAL")!;

        Assert.Equal((1, 1, 1, 1, 1, 1), (facts.Pins, facts.Replacements, facts.Pictures, facts.Baselines, facts.Files, facts.Links));
        Assert.False(facts.EveryRegionComes);
        Assert.False(facts.OtherRegionsLeftOut);
    }

    [Fact]
    public void Its_labels_are_each_component_once()
    {
        Assert.Equal(["U7", "U9"], ComponentClone.Labels(Source()));
        Assert.Equal(["U7"], ComponentClone.Labels(Regional()));
    }
}
