using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Deleting a component from a draft's BOARD - what the component pane's "Delete component" does.
// ###########################################################################################
public sealed class ComponentBoardWriterTests
{
    [Fact]
    public void Deleting_a_component_removes_EVERY_component_scoped_row_for_it()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8" });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U8", Name = "Clock" });
        board.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet" });
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Ref" });

        BoardData result = ComponentBoardWriter.ApplyComponentDelete(board, "U8");

        Assert.Empty(result.Components);
        Assert.Empty(result.ComponentImages);
        Assert.Empty(result.ComponentLocalFiles);
        Assert.Empty(result.ComponentLinks);
    }

    [Fact]
    public void Deleting_a_component_leaves_OTHER_components_alone()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U9" });

        BoardData result = ComponentBoardWriter.ApplyComponentDelete(board, "U8");

        Assert.Equal("U9", Assert.Single(result.Components).BoardLabel);
    }

    // ###########################################################################################
    // *** A DELETE TAKES THE COMPONENT'S HIGHLIGHTS TOO, AND LEAVES THE BOARD'S OWN ROWS ALONE. ***
    //
    // Changed 2026-09-25 (owner: "if a component really is deleted, then it should remove
    // EVERYTHING related to this component"). Highlights used to be kept as "the label editor's
    // concern", which left rectangles naming a component the board no longer had. The board's own
    // files and links are not the component's, and stay.
    // ###########################################################################################
    [Fact]
    public void Deleting_a_component_removes_its_highlights_but_not_the_boards_own_rows()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8" });
        board.ComponentHighlights.Add(new ComponentHighlightEntry
        {
            SchematicName = "Sheet 1",
            BoardLabel = "U8",
        });
        board.BoardLinks.Add(new BoardLinkEntry { Category = "Docs", Name = "Manual" });

        board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 2", BoardLabel = "u8" });
        board.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U9" });

        BoardData result = ComponentBoardWriter.ApplyComponentDelete(board, "U8");

        Assert.Equal("U9", Assert.Single(result.ComponentHighlights).BoardLabel);
        Assert.Single(result.BoardLinks);
    }

    [Fact]
    public void A_board_label_is_matched_case_insensitively()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });

        Assert.Empty(ComponentBoardWriter.ApplyComponentDelete(board, "u8").Components);
    }

    // A blank label identifies no component; treating it as a wildcard would take every row whose own
    // label happened to be blank.
    [Fact]
    public void A_blank_label_deletes_nothing()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = string.Empty, FriendlyName = "Not labelled yet" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U8" });

        Assert.Equal(2, ComponentBoardWriter.ApplyComponentDelete(board, "  ").Components.Count);
    }

    [Fact]
    public void The_INPUT_board_is_never_mutated()
    {
        // DraftWorkbookStore.Edit hands over the board it read and writes whatever comes back, so a
        // failed write must leave the caller's copy untouched.
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });

        ComponentBoardWriter.ApplyComponentDelete(board, "U8");

        Assert.Equal("CPU", Assert.Single(board.Components).FriendlyName);
    }

    // ###########################################################################################
    // *** A DELETE KEEPS THE "# Hardware:" / "# Board:" CAPTION (owner report, 2026-09-28). *** The
    // writers built the next board by listing its sections by hand and left the caption out, so a
    // draft saved that way lost the first two lines of every sheet.
    // ###########################################################################################
    [Fact]
    public void Deleting_a_component_keeps_the_caption()
    {
        var board = new BoardData { HardwareName = "Commodore 128 and 128D", BoardName = "310378" };
        board.Components.Add(new ComponentEntry { BoardLabel = "U8" });

        BoardData deleted = ComponentBoardWriter.ApplyComponentDelete(board, "U8");

        Assert.Equal(("Commodore 128 and 128D", "310378"), (deleted.HardwareName, deleted.BoardName));
    }

    // Deleting the component takes its pins and replacement parts along - and no other component's.
    [Fact]
    public void Deleting_a_component_takes_its_pins_and_replacements_along()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", FriendlyName = "CIA", Region = "PAL", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U9", FriendlyName = "SID" },
            ],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" }, new ComponentPinEntry { BoardLabel = "U9", Pin = "1", Name = "CAP1A" }],
            ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "6526A", Kind = "Drop-in replacement" }],
        };

        BoardData deleted = ComponentBoardWriter.ApplyComponentDelete(board, "u8");

        Assert.Equal("U9", Assert.Single(deleted.ComponentPins).BoardLabel);
        Assert.Empty(deleted.ComponentReplacements);
    }

    // Every row of it - images, files, links and highlights - and every other component and the
    // board's own sections as they were.
    [Fact]
    public void Deleting_a_component_takes_every_row_of_it_and_leaves_the_boards_own_sections()
    {
        var board = new BoardData
        {
            RevisionDate = "2026-01-15",
            Components =
            {
                new ComponentEntry { BoardLabel = "U1", FriendlyName = "VIC-II", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            },
            ComponentImages = { new ComponentImageEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "Pin 1", File = "Scope baseline/u1-pin1.png" } },
            ComponentHighlights = { new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U1", X = "10", Y = "20", Width = "30", Height = "40" } },
            ComponentLocalFiles = { new ComponentLocalFileEntry { BoardLabel = "U1", Name = "Datasheet", File = "Component local files/vic.pdf" } },
            ComponentLinks = { new ComponentLinkEntry { BoardLabel = "U1", Name = "Datasheet", Url = "https://example.com/vic" } },
            BoardLocalFiles =
            {
                new BoardLocalFileEntry { Category = "Schematics", Name = "Sheet 1" },
                new BoardLocalFileEntry { Category = "Schematics", Name = "Sheet 2" },
            },
            BoardLinks = { new BoardLinkEntry { Category = "Repairs", Name = "Repair log", Url = "https://example.com/log" } },
        };

        BoardData result = ComponentBoardWriter.ApplyComponentDelete(board, "u1");

        Assert.Equal(["C1"], result.Components.Select(row => row.BoardLabel));
        Assert.Empty(result.ComponentImages);
        Assert.Empty(result.ComponentLocalFiles);
        Assert.Empty(result.ComponentLinks);
        Assert.Empty(result.ComponentHighlights);
        Assert.Equal(2, result.BoardLocalFiles.Count);
        Assert.Single(result.BoardLinks);
        Assert.Equal("2026-01-15", result.RevisionDate);
    }
}
