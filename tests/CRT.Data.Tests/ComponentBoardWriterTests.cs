using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Saving a component-contribution editing session into a draft's BOARD
// (NewContributeStrategy.md Phase 6 - maintainer request, 2026-09-23).
//
// Replaces ComponentDraftWriterTests, which pinned the same behaviours expressed as BoardDraft
// row deltas across six sections. What has gone is everything that existed only to describe an
// overlay: Added-vs-Modified per row, deletion tombstones, and the PURE official BoardData that
// had to be loaded to tell those two apart.
//
// The two rules worth reading before touching this file:
//   - the save is COMPONENT-SCOPED (this file's equivalent of the label editor's
//     schematic-scoped replace) - other components must survive untouched;
//   - BoardLocalFiles and BoardLinks are BOARD-scoped. The window shows all of them, but they are
//     MERGED against what it loaded rather than replaced, so a row added elsewhere (in Excel)
//     while it was open survives the save (code review, 2026-09-25).
// ###########################################################################################
public sealed class ComponentBoardWriterTests
{
    private static ComponentDraftWriter.ComponentDraftRow ComponentRow(
        string boardLabel,
        string friendlyName = "",
        string partNumber = "")
        => new()
        {
            BoardLabel = boardLabel,
            FriendlyName = friendlyName,
            PartNumber = partNumber,
            Category = "IC",
        };

    private static ComponentDraftWriter.LinkDraftRow LinkRow(string boardLabel, string name, string url) =>
        new() { BoardLabel = boardLabel, Name = name, Url = url };

    private static ComponentDraftWriter.LocalFileDraftRow FileRow(
        string boardLabel,
        string name,
        string fileLocation = "",
        string file = "")
        => new() { BoardLabel = boardLabel, Name = name, FileLocation = fileLocation, File = file };

    private static BoardData Save(
        BoardData current,
        string boardLabel,
        IReadOnlyList<ComponentDraftWriter.ComponentDraftRow>? components = null,
        IReadOnlyList<ComponentDraftWriter.ComponentImageDraftRow>? images = null,
        IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow>? componentFiles = null,
        IReadOnlyList<ComponentDraftWriter.LinkDraftRow>? componentLinks = null,
        IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow>? boardFiles = null,
        IReadOnlyList<ComponentDraftWriter.LinkDraftRow>? boardLinks = null,
        string editedRegion = "")
        => ComponentBoardWriter.ApplyComponentSave(
            current,
            boardLabel,
            components ?? [],
            images ?? [],
            componentFiles ?? [],
            componentLinks ?? [],
            boardFiles ?? current.BoardLocalFiles.Select(existing =>
                new ComponentDraftWriter.LocalFileDraftRow
                {
                    Category = existing.Category,
                    Name = existing.Name,
                    File = existing.File,
                }).ToList(),
            boardLinks ?? current.BoardLinks.Select(existing =>
                new ComponentDraftWriter.LinkDraftRow
                {
                    Category = existing.Category,
                    Name = existing.Name,
                    Url = existing.Url,
                }).ToList(),
            editedRegion);

    // ------------------------------------------------------------- Region-scoped image replace

    // ###########################################################################################
    // *** REGRESSION (2026-09-23): SAVING ONE REGION MUST NOT DELETE THE OTHER REGION'S IMAGES. ***
    //
    // The contribution window loads component images FILTERED BY REGION ("Component images
    // relevant for the PAL region"), while it loads every other section it edits unfiltered. The
    // writer replaced the label's image rows wholesale, so the rows the window never saw were
    // silently dropped.
    //
    // Reported by the maintainer: editing U1's short description while viewing PAL deleted all 40
    // of U1's NTSC scope baselines. The review app then correctly offered a submission that
    // removed 40 files, from a change that was meant to touch one line of text - and approving it
    // would have destroyed them on the server.
    // ###########################################################################################
    [Fact]
    public void Saving_while_viewing_one_region_KEEPS_the_other_regions_images()
    {
        var board = new BoardData
        {
            ComponentImages =
            {
                new ComponentImageEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", File = "U1_1_PAL.png" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", File = "U1_1_NTSC.png" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "NTSC", Pin = "2", File = "U1_2_NTSC.png" },
            },
        };

        // The window was opened on PAL, so it hands back only the PAL row.
        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U1",
            images:
            [
                new ComponentDraftWriter.ComponentImageDraftRow
                {
                    BoardLabel = "U1", Region = "PAL", Pin = "1", File = "U1_1_PAL.png",
                },
            ],
            editedRegion: "PAL");

        Assert.Equal(2, result.ComponentImages.Count(image =>
            string.Equals(image.Region, "NTSC", StringComparison.OrdinalIgnoreCase)));

        Assert.Single(result.ComponentImages, image =>
            string.Equals(image.Region, "PAL", StringComparison.OrdinalIgnoreCase));
    }

    // The other half: within the edited region the replace must still be a REPLACE. Without this,
    // the fix above could have been "never remove an image", which would make deletion impossible.
    [Fact]
    public void An_image_removed_in_the_edited_region_is_still_removed()
    {
        var board = new BoardData
        {
            ComponentImages =
            {
                new ComponentImageEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", File = "U1_1_PAL.png" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "PAL", Pin = "2", File = "U1_2_PAL.png" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", File = "U1_1_NTSC.png" },
            },
        };

        // The contributor deleted the pin-2 row while viewing PAL.
        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U1",
            images:
            [
                new ComponentDraftWriter.ComponentImageDraftRow
                {
                    BoardLabel = "U1", Region = "PAL", Pin = "1", File = "U1_1_PAL.png",
                },
            ],
            editedRegion: "PAL");

        Assert.DoesNotContain(result.ComponentImages, image => image.File == "U1_2_PAL.png");
        Assert.Contains(result.ComponentImages, image => image.File == "U1_1_NTSC.png");
    }

    // A BLANK-region row is shown in every region, so the window always loads it and it is inside
    // the replaced set. Preserving it instead would duplicate it on the next save.
    [Fact]
    public void A_blank_region_image_is_replaced_rather_than_preserved()
    {
        var board = new BoardData
        {
            ComponentImages =
            {
                new ComponentImageEntry { BoardLabel = "U1", Region = "", Pin = "", File = "pinout.png" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", File = "U1_1_NTSC.png" },
            },
        };

        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U1",
            images:
            [
                new ComponentDraftWriter.ComponentImageDraftRow
                {
                    BoardLabel = "U1", Region = "", Pin = "", File = "pinout.png",
                },
            ],
            editedRegion: "PAL");

        Assert.Single(result.ComponentImages, image => image.File == "pinout.png");
        Assert.Contains(result.ComponentImages, image => image.File == "U1_1_NTSC.png");
    }

    // A window opened with NO region can only have loaded the blank-region rows, so it must not be
    // treated as a wildcard that replaces every region - that is the original bug in another form.
    [Fact]
    public void An_edit_with_no_region_at_all_does_not_touch_region_scoped_images()
    {
        var board = new BoardData
        {
            ComponentImages =
            {
                new ComponentImageEntry { BoardLabel = "U1", Region = "", Pin = "", File = "pinout.png" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", File = "U1_1_PAL.png" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", File = "U1_1_NTSC.png" },
            },
        };

        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U1",
            images:
            [
                new ComponentDraftWriter.ComponentImageDraftRow
                {
                    BoardLabel = "U1", Region = "", Pin = "", File = "pinout.png",
                },
            ],
            editedRegion: "");

        Assert.Contains(result.ComponentImages, image => image.File == "U1_1_PAL.png");
        Assert.Contains(result.ComponentImages, image => image.File == "U1_1_NTSC.png");
        Assert.Single(result.ComponentImages, image => image.File == "pinout.png");
    }

    // ------------------------------------------------------------------ The basic save

    [Fact]
    public void A_new_component_is_added_to_the_board()
    {
        BoardData result = ComponentBoardWriterTests.Save(
            new BoardData(),
            "U8",
            components: [ComponentBoardWriterTests.ComponentRow("U8", "CPU", "6510")]);

        ComponentEntry component = Assert.Single(result.Components);

        Assert.Equal("U8", component.BoardLabel);
        Assert.Equal("CPU", component.FriendlyName);
        Assert.Equal("6510", component.PartNumber);
    }

    [Fact]
    public void An_EXISTING_components_row_is_replaced_not_duplicated()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });

        BoardData result = ComponentBoardWriterTests.Save(
            board, "U8", components: [ComponentBoardWriterTests.ComponentRow("U8", "MPU")]);

        Assert.Equal("MPU", Assert.Single(result.Components).FriendlyName);
    }

    // ###########################################################################################
    // *** THE COMPONENT-SCOPED REPLACE, and the rule most dangerous to get wrong. ***
    //
    // The window loads every row belonging to THIS component before editing starts, so the rows it
    // hands back are that component's complete set - anything missing was removed. But rows
    // belonging to every OTHER component must survive untouched, or a save from a window the
    // contributor believes edits one component would wipe the rest of the board.
    // ###########################################################################################
    [Fact]
    public void ANOTHER_components_rows_survive_the_save_untouched()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U9", FriendlyName = "VIC" });
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U9", Name = "Datasheet" });

        BoardData result = ComponentBoardWriterTests.Save(
            board, "U8", components: [ComponentBoardWriterTests.ComponentRow("U8", "MPU")]);

        Assert.Equal("VIC", result.Components.Single(c => c.BoardLabel == "U9").FriendlyName);
        Assert.Single(result.ComponentLinks);
    }

    // ###########################################################################################
    // A row the contributor removed from the window's collection is a row that is GONE.
    //
    // No tombstone is needed: the rows handed back are this component's complete set, so absence
    // is removal. That is the whole simplification of dropping the delta model.
    // ###########################################################################################
    [Fact]
    public void A_row_REMOVED_in_the_window_disappears_from_the_board()
    {
        var board = new BoardData();
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Keep" });
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Remove" });

        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U8",
            componentLinks: [ComponentBoardWriterTests.LinkRow("U8", "Keep", "https://example.com")]);

        Assert.Equal("Keep", Assert.Single(result.ComponentLinks).Name);
    }

    [Fact]
    public void A_row_with_a_BLANK_label_inherits_the_sessions_component()
    {
        // Matches the window's own ResolveEffectiveBoardLabel: a row added in the editor does not
        // repeat the label the whole session already belongs to.
        BoardData result = ComponentBoardWriterTests.Save(
            new BoardData(),
            "U8",
            componentLinks: [ComponentBoardWriterTests.LinkRow("", "Datasheet", "https://example.com")]);

        Assert.Equal("U8", Assert.Single(result.ComponentLinks).BoardLabel);
    }

    // ------------------------------------------------------------------ Scope boundaries

    // ###########################################################################################
    // *** HIGHLIGHTS ARE THE LABEL EDITOR'S BUSINESS, NEVER THIS WINDOW'S. ***
    //
    // A component save must not disturb a single rectangle - the two editors write different
    // sections of the same board, and this one touching highlights would silently undo label
    // editor work whenever both were used on the same component.
    // ###########################################################################################
    [Fact]
    public void Component_HIGHLIGHTS_are_never_touched_by_a_component_save()
    {
        var board = new BoardData();
        board.ComponentHighlights.Add(new ComponentHighlightEntry
        {
            SchematicName = "Sheet 1",
            BoardLabel = "U8",
            X = "1",
            Y = "2",
            Width = "3",
            Height = "4",
        });

        BoardData result = ComponentBoardWriterTests.Save(
            board, "U8", components: [ComponentBoardWriterTests.ComponentRow("U8", "CPU")]);

        Assert.Single(result.ComponentHighlights);
    }

    // ###########################################################################################
    // BOARD-scoped sections are replaced IN FULL, not per component, when the caller gives NO
    // baseline - the window's rows are then the complete set for the whole board. The app always
    // passes the rows it loaded, and the merge that enables is pinned under "Board-scoped merge".
    // ###########################################################################################
    [Fact]
    public void BOARD_scoped_sections_are_replaced_in_full()
    {
        var board = new BoardData();
        board.BoardLinks.Add(new BoardLinkEntry { Category = "Docs", Name = "Old" });

        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U8",
            boardLinks: [new ComponentDraftWriter.LinkDraftRow
            {
                Category = "Docs",
                Name = "New",
                Url = "https://example.com",
            }]);

        Assert.Equal("New", Assert.Single(result.BoardLinks).Name);
    }

    [Fact]
    public void Credits_and_schematics_are_carried_across_untouched()
    {
        // A section left out of the returned board is a section erased from the workbook, silently.
        var board = new BoardData
        {
            RevisionDate = "2026-09-01",
            Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1" }],
            Credits = [new CreditEntry { Category = "Data", NameOrHandle = "Dennis" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "PHI2" }],
        };

        BoardData result = ComponentBoardWriterTests.Save(
            board, "U8", components: [ComponentBoardWriterTests.ComponentRow("U8")]);

        Assert.Equal("2026-09-01", result.RevisionDate);
        Assert.Single(result.Schematics);
        Assert.Single(result.Credits);
        Assert.Single(result.KiCadImportantSignals);
    }

    // ------------------------------------------------------------------ File paths

    // ###########################################################################################
    // The stored File value is FileLocation + "/" + File, carried over VERBATIM from
    // ComponentDraftWriter.BuildStoredFile - a path stored before this change and one stored after
    // it must be byte-identical, or every attachment row would read as modified the first time a
    // component was re-saved.
    // ###########################################################################################
    [Fact]
    public void A_file_location_and_name_are_joined_with_a_forward_slash()
    {
        BoardData result = ComponentBoardWriterTests.Save(
            new BoardData(),
            "U8",
            componentFiles: [ComponentBoardWriterTests.FileRow("U8", "Datasheet", "Datasheets", "U8.pdf")]);

        Assert.Equal("Datasheets/U8.pdf", Assert.Single(result.ComponentLocalFiles).File);
    }

    [Fact]
    public void A_backslash_in_the_location_is_normalized()
    {
        // A contributor can type a Windows-style folder into the window, and the stored path is
        // read by the Linux server too.
        BoardData result = ComponentBoardWriterTests.Save(
            new BoardData(),
            "U8",
            componentFiles: [ComponentBoardWriterTests.FileRow("U8", "Datasheet", @"Docs\Datasheets", "U8.pdf")]);

        Assert.Equal("Docs/Datasheets/U8.pdf", Assert.Single(result.ComponentLocalFiles).File);
    }

    [Fact]
    public void A_row_with_NO_file_stores_an_empty_path_rather_than_a_bare_folder()
    {
        BoardData result = ComponentBoardWriterTests.Save(
            new BoardData(),
            "U8",
            componentFiles: [ComponentBoardWriterTests.FileRow("U8", "Datasheet", "Datasheets", "")]);

        Assert.Equal(string.Empty, Assert.Single(result.ComponentLocalFiles).File);
    }

    // ------------------------------------------------------------------ Delete

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
    // Changed 2026-09-25 (maintainer: "if a component really is deleted, then it should remove
    // EVERYTHING related to this component"). Highlights used to be kept as "the label editor's
    // concern", which left rectangles naming a component the board no longer had - while this
    // window's own notice said they would go. The board's own files and links are not the
    // component's, and stay.
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
    public void Board_labels_are_matched_case_insensitively_on_both_paths()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });

        Assert.Single(ComponentBoardWriterTests.Save(
            board, "u8", components: [ComponentBoardWriterTests.ComponentRow("u8", "MPU")]).Components);

        Assert.Empty(ComponentBoardWriter.ApplyComponentDelete(board, "u8").Components);
    }

    [Fact]
    public void The_INPUT_board_is_never_mutated()
    {
        // DraftWorkbookStore.Edit hands over the board it read and writes whatever comes back, so a
        // failed write must leave the caller's copy untouched.
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });

        ComponentBoardWriterTests.Save(
            board, "U8", components: [ComponentBoardWriterTests.ComponentRow("U8", "MPU")]);

        Assert.Equal("CPU", Assert.Single(board.Components).FriendlyName);
    }

    // ------------------------------------------------------------------ Row positions (2026-09-24)

    // ###########################################################################################
    // *** REPORTED BY THE MAINTAINER: C2 showed above C1, although C1 was added first. *** A save
    // removed the component's rows and appended them at the bottom, so re-saving C1 after C2 had
    // been added moved C1 below it. Row order is what the application's component list shows, so
    // the drift was visible. An edited component now stays where it was.
    // ###########################################################################################
    [Fact]
    public void Re_saving_an_EXISTING_component_keeps_its_position_in_the_list()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "C1", Category = "IC" },
                new ComponentEntry { BoardLabel = "C2", Category = "IC" },
            ],
        };

        BoardData result = ComponentBoardWriterTests.Save(
            board, "C1", components: [ComponentBoardWriterTests.ComponentRow("C1", "Edited")]);

        Assert.Equal(["C1", "C2"], result.Components.Select(component => component.BoardLabel));
    }

    [Fact]
    public void A_NEW_component_joins_its_category_in_label_order_not_the_bottom()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC" },
                new ComponentEntry { BoardLabel = "U3", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            ],
        };

        BoardData result = ComponentBoardWriterTests.Save(
            board, "U2", components: [ComponentBoardWriterTests.ComponentRow("U2")]);

        Assert.Equal(["U1", "U2", "U3", "C1"], result.Components.Select(component => component.BoardLabel));
    }

    [Fact]
    public void Re_saving_a_component_keeps_its_links_and_files_where_they_were()
    {
        var board = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U8", Category = "IC" }],
            ComponentLinks =
            [
                new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://a" },
                new ComponentLinkEntry { BoardLabel = "U1", Name = "Other", Url = "https://b" },
            ],
        };

        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U8",
            components: [ComponentBoardWriterTests.ComponentRow("U8")],
            componentLinks: [ComponentBoardWriterTests.LinkRow("U8", "Datasheet", "https://a2")]);

        Assert.Equal(["U8", "U1"], result.ComponentLinks.Select(link => link.BoardLabel));
        Assert.Equal("https://a2", result.ComponentLinks[0].Url);
    }

    [Fact]
    public void Re_saving_one_regions_images_keeps_them_where_that_region_was()
    {
        var board = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "pal" },
                new ComponentImageEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", Name = "ntsc" },
                new ComponentImageEntry { BoardLabel = "U2", Region = "PAL", Pin = "1", Name = "u2" },
            ],
        };

        BoardData result = ComponentBoardWriterTests.Save(
            board,
            "U1",
            components: [ComponentBoardWriterTests.ComponentRow("U1")],
            images: [new ComponentDraftWriter.ComponentImageDraftRow { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "pal edited" }],
            editedRegion: "PAL");

        Assert.Equal(["pal edited", "ntsc", "u2"], result.ComponentImages.Select(image => image.Name));
    }

    // ------------------------------------------------------------------ Board-scoped merge

    private static ComponentDraftWriter.LinkDraftRow BoardLink(string name, string url = "https://example.com") =>
        new() { Category = "Docs", Name = name, Url = url };

    private static BoardLinkEntry BoardLinkEntryOf(string name, string url = "https://example.com") =>
        new() { Category = "Docs", Name = name, Url = url };

    // Saves with an explicit baseline - the rows the window LOADED - which is what the app passes.
    private static BoardData SaveBoardLinks(
        BoardData current,
        IReadOnlyList<ComponentDraftWriter.LinkDraftRow> atOpen,
        IReadOnlyList<ComponentDraftWriter.LinkDraftRow> edited)
        => ComponentBoardWriter.ApplyComponentSave(
            current,
            "U8",
            [ComponentBoardWriterTests.ComponentRow("U8")],
            [],
            [],
            [],
            [],
            edited,
            string.Empty,
            boardLocalFileRowsAtOpen: [],
            boardLinkRowsAtOpen: atOpen);

    // ###########################################################################################
    // *** THE DATA-LOSS CASE (code review, 2026-09-25). ***
    //
    // The window loads every board link when it opens. Excel is the one writer the app cannot
    // hold back, so a link added there while the window sat open used to be DELETED by the next
    // "Save to draft" - the section was replaced wholesale with the window's stale copy, although
    // the save only meant to change one component. The window did not touch the section, so the
    // disk must win.
    // ###########################################################################################
    [Fact]
    public void A_board_link_added_ELSEWHERE_survives_a_save_that_did_not_touch_board_links()
    {
        var current = new BoardData();
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("Manual"));
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("Added in Excel"));

        BoardData result = ComponentBoardWriterTests.SaveBoardLinks(
            current,
            atOpen: [ComponentBoardWriterTests.BoardLink("Manual")],
            edited: [ComponentBoardWriterTests.BoardLink("Manual")]);

        Assert.Equal(["Manual", "Added in Excel"], result.BoardLinks.Select(link => link.Name));
    }

    [Fact]
    public void With_nothing_changed_elsewhere_the_windows_links_win_in_the_windows_order()
    {
        // The common case, and exactly the old behaviour: the window's rows are the whole set.
        var current = new BoardData();
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("A"));
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("B"));

        BoardData result = ComponentBoardWriterTests.SaveBoardLinks(
            current,
            atOpen: [ComponentBoardWriterTests.BoardLink("A"), ComponentBoardWriterTests.BoardLink("B")],
            edited: [ComponentBoardWriterTests.BoardLink("B"), ComponentBoardWriterTests.BoardLink("A"), ComponentBoardWriterTests.BoardLink("C")]);

        Assert.Equal(["B", "A", "C"], result.BoardLinks.Select(link => link.Name));
    }

    [Fact]
    public void When_BOTH_sides_changed_an_edited_link_replaces_its_original_and_the_outside_one_stays()
    {
        // The window edited B's URL; Excel added X. Both edits must land, and the edited row
        // must sit where the original was rather than jumping to the end.
        var current = new BoardData();
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("A"));
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("B", "https://old"));
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("X"));

        BoardData result = ComponentBoardWriterTests.SaveBoardLinks(
            current,
            atOpen: [ComponentBoardWriterTests.BoardLink("A"), ComponentBoardWriterTests.BoardLink("B", "https://old")],
            edited: [ComponentBoardWriterTests.BoardLink("A"), ComponentBoardWriterTests.BoardLink("B", "https://new")]);

        Assert.Equal(["A", "B", "X"], result.BoardLinks.Select(link => link.Name));
        Assert.Equal("https://new", result.BoardLinks[1].Url);
    }

    [Fact]
    public void When_BOTH_sides_changed_a_link_removed_in_the_window_goes_and_the_outside_one_stays()
    {
        var current = new BoardData();
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("A"));
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("B"));
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("X"));

        BoardData result = ComponentBoardWriterTests.SaveBoardLinks(
            current,
            atOpen: [ComponentBoardWriterTests.BoardLink("A"), ComponentBoardWriterTests.BoardLink("B")],
            edited: [ComponentBoardWriterTests.BoardLink("B")]);

        Assert.Equal(["B", "X"], result.BoardLinks.Select(link => link.Name));
    }

    // ###########################################################################################
    // A second save from the same window must never duplicate what the first one wrote. The
    // window moves its baseline forward after each save, but the merge is idempotent on its own
    // too, so a stale baseline cannot add a row that is already there.
    // ###########################################################################################
    [Fact]
    public void Merging_again_with_a_stale_baseline_does_not_duplicate_a_row()
    {
        var current = new BoardData();
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("A"));
        current.BoardLinks.Add(ComponentBoardWriterTests.BoardLinkEntryOf("C"));

        BoardData result = ComponentBoardWriterTests.SaveBoardLinks(
            current,
            atOpen: [ComponentBoardWriterTests.BoardLink("A")],
            edited: [ComponentBoardWriterTests.BoardLink("A"), ComponentBoardWriterTests.BoardLink("C")]);

        Assert.Equal(["A", "C"], result.BoardLinks.Select(link => link.Name));
    }

    [Fact]
    public void A_board_file_stored_with_backslashes_still_matches_the_windows_copy()
    {
        // The workbook may hold "Docs\manual.pdf" (typed by hand on Windows) while the window
        // rebuilds it as "Docs/manual.pdf". They are the same row, so an untouched section must
        // still read as untouched and keep the file added in Excel.
        var current = new BoardData();
        current.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Docs", Name = "Manual", File = "Docs\\manual.pdf" });
        current.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Docs", Name = "Excel", File = "Docs/excel.pdf" });

        ComponentDraftWriter.LocalFileDraftRow manual = new() { Category = "Docs", Name = "Manual", FileLocation = "Docs", File = "manual.pdf" };

        BoardData result = ComponentBoardWriter.ApplyComponentSave(
            current,
            "U8",
            [ComponentBoardWriterTests.ComponentRow("U8")],
            [],
            [],
            [],
            [manual],
            [],
            string.Empty,
            boardLocalFileRowsAtOpen: [manual],
            boardLinkRowsAtOpen: []);

        Assert.Equal(["Manual", "Excel"], result.BoardLocalFiles.Select(file => file.Name));
    }
}
