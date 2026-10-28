using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SchematicCadNameSuggestions - which KiCad view each picture seems to show (CRT 4.0 step 3, slice
// 3d - the plan: "Match: suggested pairs ("...Top.png" with "PCB Top") to confirm"). The live boards
// with KiCad data are the proof: their pictures as published, the CAD names taken away, and their
// projects' views as CRT's loader names them (read off the shipped KiCad data, 2026-10-28) - every
// suggestion must be the CAD name the board has, and where the words cannot tell, none at all.
// ###########################################################################################
public sealed class SchematicCadNameSuggestionsTests
{
    private static BoardSchematicEntry Picture(string name, string file, string cadName = "") =>
        new() { SchematicName = name, SchematicImageFile = file, CadName = cadName };

    private static KiCadViewToMatch Top(string pcb) => new($"{pcb} - PCB Top", KiCadViewSide.Top, pcb);

    private static KiCadViewToMatch Bottom(string pcb) => new($"{pcb} - PCB Bottom", KiCadViewSide.Bottom, pcb);

    private static KiCadViewToMatch Sheet(string name) => new(name, KiCadViewSide.Sheet);

    private static Dictionary<string, string> Suggested(IReadOnlyList<BoardSchematicEntry> pictures, IReadOnlyList<KiCadViewToMatch> views) =>
        SchematicCadNameSuggestions.For(pictures, views).ToDictionary(pair => pair.SchematicName, pair => pair.CadName);

    // ---- the live boards ------------------------------------------------------------------------------

    // The C128's Open128 replica: 24 pictures, named after its sheets - and plotted by KiCad as
    // "<project>-<parent>-<sheet>.png", so "I/O" is in most of the files' names and only the last
    // words say which sheet. Every one is matched as the board has it.
    [Fact]
    public void The_Open128_pictures_are_each_matched_to_the_view_the_board_gives_them()
    {
        (BoardSchematicEntry Picture, string CadName)[] live =
        [
            (Picture("PCB; Top", "C128 Neo.3.F.png"), "Open128 - PCB Top"),
            (Picture("PCB; Bottom", "C128 Neo.3.B.png"), "Open128 - PCB Bottom"),
            (Picture("Address Multiplexing", "Open128-Address Multiplexing.png"), "Address Multiplexing"),
            (Picture("Clock", "Open128-Clock.png"), "Clock"),
            (Picture("CPU 8502", "Open128-CPU 8502.png"), "CPU 8502"),
            (Picture("CPU Z80", "Open128-CPU Z80.png"), "CPU Z80"),
            (Picture("DRAM", "Open128-DRAM.png"), "DRAM"),
            (Picture("Expansion Port", "Open128-Expansion Port.png"), "Expansion Port"),
            (Picture("I/O", "Open128-I_O.png"), "I/O"),
            (Picture("I/O; Cassette Interface", "Open128-I_O-Casette Interface.png"), "Casette Interface"),
            (Picture("I/O; Control Ports", "Open128-I_O-Control Ports.png"), "Control Ports"),
            (Picture("I/O; Control Ports EMI", "Open128-I_O-Control Ports-Control Ports-EMI.png"), "Control Ports-EMI"),
            (Picture("I/O; Keyboard", "Open128-I_O-Keyboard.png"), "Keyboard"),
            (Picture("I/O; Keyboard EMI", "Open128-I_O-Keyboard-Key-EMI.png"), "Key-EMI"),
            (Picture("I/O; Serial Bus", "Open128-I_O-Serial Bus.png"), "Serial Bus"),
            (Picture("I/O; Userport", "Open128-I_O-Userport.png"), "Userport"),
            (Picture("Memory Management", "Open128-Memory Management.png"), "Memory Management"),
            (Picture("Power", "Open128-Power.png"), "Power"),
            (Picture("Reset", "Open128-Reset.png"), "Reset"),
            (Picture("ROMs", "Open128-ROMs.png"), "ROMs"),
            (Picture("Sound", "Open128-Sound.png"), "Sound"),
            (Picture("VDC", "Open128-VDC.png"), "VDC"),
            (Picture("VIC", "Open128-VIC.png"), "VIC"),
            (Picture("VIC; RF Modulator & AV connectors", "Open128-VIC-RF Modulator and AV connectors.png"), "RF Modulator and AV connectors"),
        ];

        KiCadViewToMatch[] views =
        [
            Top("Open128"), Bottom("Open128"),
            Sheet("Open128"), Sheet("Reset"), Sheet("Power"), Sheet("I/O"), Sheet("Keyboard"), Sheet("Key-EMI"),
            Sheet("Control Ports"), Sheet("Control Ports-EMI"), Sheet("Userport"), Sheet("Serial Bus"),
            Sheet("Casette Interface"), Sheet("Sound"), Sheet("Memory Management"), Sheet("ROMs"), Sheet("DRAM"),
            Sheet("VIC"), Sheet("RF Modulator and AV connectors"), Sheet("VDC"), Sheet("CPU 8502"), Sheet("CPU Z80"),
            Sheet("Expansion Port"), Sheet("Clock"), Sheet("Address Multiplexing"),
        ];

        Dictionary<string, string> suggested = Suggested(live.Select(row => row.Picture).ToList(), views);

        Assert.Equal(live.ToDictionary(row => row.Picture.SchematicName, row => row.CadName), suggested);
    }

    // The C64 250407: its two sides by their words; its two schematic sheets NOT - the first sheet is
    // "250407_", a word of three of the pictures' files, and the second "250407_sheet2", a word of
    // none. Where the words cannot tell, nothing is suggested.
    [Fact]
    public void The_250407_sides_are_matched_and_its_sheets_left_alone()
    {
        BoardSchematicEntry[] pictures =
        [
            Picture("Board layout", "Board Layout 250407 NTSC.png"),
            Picture("Board layout (CBM)", "c64-23.gif"),
            Picture("Top (replica)", "C64 250407 PCB Replica.1.1 Top.png"),
            Picture("Bottom (replica)", "C64 250407 PCB Replica.1.1 Bottom.png"),
            Picture("Schematics #1 of 2", "C64 250407 PCB Replica 1of2.png"),
            Picture("Schematics #2 of 2", "C64 250407 PCB Replica 2of2.png"),
        ];

        Dictionary<string, string> suggested = Suggested(pictures, [Top("250407_"), Bottom("250407_"), Sheet("250407_"), Sheet("250407_sheet2")]);

        Assert.Equal(
            new Dictionary<string, string>
            {
                ["Top (replica)"] = "250407_ - PCB Top",
                ["Bottom (replica)"] = "250407_ - PCB Bottom",
            },
            suggested);
    }

    // The C64 250469: its schematic by its file's name, which IS the sheet's; the layout drawing by nothing.
    [Fact]
    public void The_250469_pictures_are_each_matched_to_the_view_the_board_gives_them()
    {
        BoardSchematicEntry[] pictures =
        [
            Picture("Board layout", "Board-Layout-250469 (NTSC).jpg"),
            Picture("KiCAD replica; Top", "Replica of 250469 REV.4.1.1.Top.png"),
            Picture("KiCAD replica; Bottom", "Replica of 250469 REV.4.1.1.Bottom.png"),
            Picture("KiCAD schematics", "C64-250469-KiCad.png"),
        ];

        Dictionary<string, string> suggested = Suggested(pictures, [Top("C64-250469-KiCad"), Bottom("C64-250469-KiCad"), Sheet("C64-250469-KiCad")]);

        Assert.Equal(
            new Dictionary<string, string>
            {
                ["KiCAD replica; Top"] = "C64-250469-KiCad - PCB Top",
                ["KiCAD replica; Bottom"] = "C64-250469-KiCad - PCB Bottom",
                ["KiCAD schematics"] = "C64-250469-KiCad",
            },
            suggested);
    }

    // The C64 KU-14194HB: the schematic is the child sheet its file's name ENDS with ("...-Sheet1"),
    // not the first sheet, whose longer name stands earlier in it; and its "REV-B" is no bottom side.
    [Fact]
    public void The_KU_14194HB_schematic_is_the_sheet_its_file_name_ends_with()
    {
        BoardSchematicEntry[] pictures =
        [
            Picture("Top (replica)", "C64 Motherboard KU-14194HB.1.4.F.png"),
            Picture("Bottom (replica)", "C64 Motherboard KU-14194HB.1.4.B.png"),
            Picture("Schematics (replica)", "KU-14194HB-REV-B-KiCad-Sheet1.png"),
        ];

        const string pcb = "KU-14194HB-REV-B-KiCad";
        Dictionary<string, string> suggested = Suggested(pictures, [Top(pcb), Bottom(pcb), Sheet(pcb), Sheet("Sheet5D659BBD"), Sheet("Sheet1")]);

        Assert.Equal(
            new Dictionary<string, string>
            {
                ["Top (replica)"] = "KU-14194HB-REV-B-KiCad - PCB Top",
                ["Bottom (replica)"] = "KU-14194HB-REV-B-KiCad - PCB Bottom",
                ["Schematics (replica)"] = "Sheet1",
            },
            suggested);
    }

    // The ZX Spectrum Issue 4B: a PCB and no schematic - the scanned schematic is matched to nothing.
    [Fact]
    public void The_Spectrum_sides_are_matched_and_its_scanned_schematic_is_not()
    {
        BoardSchematicEntry[] pictures =
        [
            Picture("Top (replica)", "Issue 4.F.png"),
            Picture("Bottom (replica)", "Issue 4.B.png"),
            Picture("Schematics", "Issue 4.png"),
        ];

        Dictionary<string, string> suggested = Suggested(pictures, [Top("Issue 4"), Bottom("Issue 4")]);

        Assert.Equal(
            new Dictionary<string, string> { ["Top (replica)"] = "Issue 4 - PCB Top", ["Bottom (replica)"] = "Issue 4 - PCB Bottom" },
            suggested);
    }

    // ---- the rules ------------------------------------------------------------------------------------

    // A new board's pictures are named after their files: KiCad's plot letters say the side.
    [Fact]
    public void KiCads_plot_letters_say_the_side_when_no_word_does()
    {
        Dictionary<string, string> suggested = Suggested(
            [Picture("C128 Neo.3.F", "Images/C128 Neo.3.F.png"), Picture("C128 Neo.3.B", "Images/C128 Neo.3.B.png")],
            [Top("Open128"), Bottom("Open128")]);

        Assert.Equal("Open128 - PCB Top", suggested["C128 Neo.3.F"]);
        Assert.Equal("Open128 - PCB Bottom", suggested["C128 Neo.3.B"]);
    }

    // ###########################################################################################
    // KiCad's plot letters are the plot's SUFFIX (".F", ".B") or its layer ("F.Cu") - never a letter
    // standing anywhere in a name (code review, 2026-10-28): "C64 Rev B" and "BBC Model B" were
    // offered as the bottom of the board, and a top photo given the bottom's view by "Use all".
    // ###########################################################################################
    [Fact]
    public void A_revision_or_model_letter_is_no_side_of_the_board()
    {
        KiCadViewToMatch[] sides = [Top("Board"), Bottom("Board")];

        Assert.Empty(Suggested([Picture("C64 Rev B", "C64 Rev B.jpg")], sides));
        Assert.Empty(Suggested([Picture("Photo", "BBC Model B.jpg")], sides));
        Assert.Empty(Suggested([Picture("Rev F board", "IMG_2001.jpg")], sides));

        // A revision or model letter after a dot is no plot suffix either (second review, 2026-10-28).
        Assert.Empty(Suggested([Picture("Photo", "250407 Rev.B.jpg")], sides));
        Assert.Empty(Suggested([Picture("C64 Rev.F", "IMG_2002.jpg")], sides));
        Assert.Empty(Suggested([Picture("Photo", "BBC Model.B.jpg")], sides));
        Assert.Empty(Suggested([Picture("Photo", "Spectrum Issue.B.jpg")], sides));

        // The plot's own spellings still say it.
        Assert.Equal("Board - PCB Bottom", Suggested([Picture("Scan", "Board.B.png")], sides)["Scan"]);
        Assert.Equal("Board - PCB Top", Suggested([Picture("Board-F_Cu", "export.png")], sides)["Board-F_Cu"]);
    }

    // Two pictures that both say "top": which one shows the PCB's top cannot be told - neither is
    // suggested, and the bottom still is.
    [Fact]
    public void Two_pictures_equally_sure_of_one_view_get_no_suggestion()
    {
        Dictionary<string, string> suggested = Suggested(
            [Picture("Top photo", "IMG_2001.jpg"), Picture("Top (replica)", "top.png"), Picture("Bottom", "IMG_2002.jpg")],
            [Top("Board"), Bottom("Board")]);

        Assert.Equal(new Dictionary<string, string> { ["Bottom"] = "Board - PCB Bottom" }, suggested);
    }

    // Two PCB files: a side alone cannot say which board's - the one whose file the words name can,
    // and once that pair is taken, the other top is the only one left for the other picture.
    [Fact]
    public void Of_two_PCBs_the_side_goes_to_the_one_the_picture_names()
    {
        Dictionary<string, string> suggested = Suggested(
            [Picture("Keyboard PCB top", "kb.png"), Picture("Top", "main.png")],
            [Top("Main"), Top("Keyboard PCB")]);

        Assert.Equal(
            new Dictionary<string, string> { ["Keyboard PCB top"] = "Keyboard PCB - PCB Top", ["Top"] = "Main - PCB Top" },
            suggested);

        // With neither named, the two tops are equally sure of the picture: none.
        Assert.Empty(Suggested([Picture("Top", "IMG_2001.jpg")], [Top("Main"), Top("Keyboard PCB")]));
    }

    // A picture that says both sides - or is a schematic - is no side of the board.
    [Fact]
    public void A_picture_saying_both_sides_or_being_a_schematic_is_no_side()
    {
        Dictionary<string, string> suggested = Suggested(
            [Picture("Top and bottom", "both.png"), Picture("Schematic top half", "sch.png")],
            [Top("Board"), Bottom("Board")]);

        Assert.Empty(suggested);
    }

    // What a picture has, and what another picture has, is never suggested again - nor is a picture
    // without an image, or one sharing its name and file with another row.
    [Fact]
    public void Only_pictures_without_a_CAD_name_and_views_no_picture_has_are_matched()
    {
        BoardSchematicEntry[] pictures =
        [
            Picture("Clock", "clock.png", cadName: "clock"),
            Picture("Clock copy", "clock-copy.png"),
            Picture("Power", ""),
            Picture("Reset", "reset.png"),
            Picture("Reset", "reset.png"),
            Picture("Sound", "sound.png"),
        ];

        Dictionary<string, string> suggested = Suggested(pictures, [Sheet("Clock"), Sheet("Power"), Sheet("Reset"), Sheet("Sound")]);

        Assert.Equal(new Dictionary<string, string> { ["Sound"] = "Sound" }, suggested);
    }

    // Capitals and "&" are no difference.
    [Fact]
    public void Capitals_and_an_ampersand_are_no_difference()
    {
        Dictionary<string, string> suggested = Suggested(
            [Picture("rf modulator & av", "a.png")],
            [Sheet("RF Modulator and AV")]);

        Assert.Equal("RF Modulator and AV", suggested["rf modulator & av"]);
    }

    // ---- taking them ----------------------------------------------------------------------------------

    // Each suggestion taken is that row's CAD name, every other column as it was.
    [Fact]
    public void Taking_suggestions_writes_each_rows_CAD_name()
    {
        var board = new BoardData();
        board.Schematics.AddRange([Picture("Top", "a/top.png"), Picture("Bottom", "a/bottom.png")]);
        board.Schematics[0] = new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "a/top.png", SchematicHighlightColor = "Red" };

        BoardData result = SchematicCadNameSuggestions.With(
            board,
            [new SchematicCadNameSuggestion("Top", "a/top.png", "Board - PCB Top"), new SchematicCadNameSuggestion("Bottom", "a/bottom.png", "Board - PCB Bottom")],
            out IReadOnlyList<SchematicCadNameSuggestion> taken);

        Assert.Equal(2, taken.Count);
        Assert.Equal(["Board - PCB Top", "Board - PCB Bottom"], result.Schematics.Select(row => row.CadName));
        Assert.Equal("Red", result.Schematics[0].SchematicHighlightColor);
    }

    // The draft moved on since the suggestions were worked out: a picture given a CAD name meanwhile,
    // a view another picture took, a row gone - none of them is taken; the rest is.
    [Fact]
    public void A_suggestion_the_draft_has_moved_past_is_not_taken()
    {
        var board = new BoardData();
        board.Schematics.AddRange([Picture("Top", "top.png", cadName: "Something else"), Picture("Photo", "photo.png", cadName: "Board - PCB Bottom"), Picture("Bottom", "bottom.png"), Picture("Clock", "clock.png")]);

        SchematicCadNameSuggestion[] suggestions =
        [
            new("Top", "top.png", "Board - PCB Top"),
            new("Bottom", "bottom.png", "Board - PCB Bottom"),
            new("Gone", "gone.png", "Reset"),
            new("Clock", "clock.png", "Clock"),
        ];

        BoardData result = SchematicCadNameSuggestions.With(board, suggestions, out IReadOnlyList<SchematicCadNameSuggestion> taken);

        Assert.Equal([suggestions[3]], taken);
        Assert.Equal([suggestions[3]], SchematicCadNameSuggestions.StillOpen(board, suggestions));
        Assert.Equal(["Something else", "Board - PCB Bottom", "", "Clock"], result.Schematics.Select(row => row.CadName));
    }

    // Two suggestions for one view - the second finds it taken by the first.
    [Fact]
    public void Two_suggestions_of_one_view_take_it_once()
    {
        var board = new BoardData();
        board.Schematics.AddRange([Picture("A", "a.png"), Picture("B", "b.png")]);

        BoardData result = SchematicCadNameSuggestions.With(board, [new("A", "a.png", "Clock"), new("B", "b.png", "clock")], out IReadOnlyList<SchematicCadNameSuggestion> taken);

        Assert.Equal("A", Assert.Single(taken).SchematicName);
        Assert.Equal(["Clock", ""], result.Schematics.Select(row => row.CadName));
    }

    // Nothing to take is the very same board - no step for the draft.
    [Fact]
    public void Nothing_to_take_is_the_same_board()
    {
        var board = new BoardData();
        board.Schematics.Add(Picture("A", "a.png", cadName: "Clock"));

        Assert.Same(board, SchematicCadNameSuggestions.With(board, [new("A", "a.png", "Reset")], out IReadOnlyList<SchematicCadNameSuggestion> taken));
        Assert.Empty(taken);
    }
}
