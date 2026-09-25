using System.Linq;
using System.Text.Json;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // The reviewer's table (2026-09-25): the rows<->board conversions, the one rule for what an
    // amendment may change (SubmissionRowsBoard.WithTableSections), and the wire shapes the server
    // writes and the review application and CRT read. The records are shared, so the wire tests
    // serialise the way ASP.NET does (web defaults, camelCase) and read back as the clients do - a
    // renamed property fails here instead of arriving as an empty table or a missing flag.
    // ###########################################################################################
    public sealed class ReviewTableContractTests
    {
        private static SubmissionRows Rows(string partNumber = "") => new()
        {
            RevisionDate = "2026-September-25",
            Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "a/sheet1.png" }],
            Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", PartNumber = partNumber }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Name = "Pin 1", File = "a/u8.png" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "a/u8.pdf" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Ref", Url = "https://example.org" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "Service", Name = "Manual", File = "a/manual.pdf" }],
            BoardLinks = [new BoardLinkEntry { Category = "Web", Name = "Site", Url = "https://example.org" }],
            Credits = [new CreditEntry { Category = "Board", NameOrHandle = "Dennis" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "CLK", KiCadNetName = "CLK" }],
            KiCadCalibrations = [new KiCadCalibrationEntry()]
        };

        [Fact]
        public void Rows_to_a_board_and_back_keep_every_section_but_the_calibrations()
        {
            SubmissionRows original = ReviewTableContractTests.Rows("251715-01");

            SubmissionRows back = SubmissionRowsBoard.FromBoard(SubmissionRowsBoard.ToBoard(original));

            Assert.Equal(original.RevisionDate, back.RevisionDate);
            Assert.Equal("251715-01", Assert.Single(back.Components).PartNumber);
            Assert.Single(back.Schematics);
            Assert.Single(back.ComponentImages);
            Assert.Single(back.ComponentHighlights);
            Assert.Single(back.ComponentLocalFiles);
            Assert.Single(back.ComponentLinks);
            Assert.Single(back.BoardLocalFiles);
            Assert.Single(back.BoardLinks);
            Assert.Single(back.Credits);
            Assert.Single(back.KiCadImportantSignals);

            // A BoardData carries no calibrations - they travel beside it.
            Assert.Empty(back.KiCadCalibrations);
        }

        [Fact]
        public void The_board_is_a_copy_so_editing_it_leaves_the_rows_alone()
        {
            SubmissionRows rows = ReviewTableContractTests.Rows();

            BoardData board = SubmissionRowsBoard.ToBoard(rows);
            board.Components.Clear();

            Assert.Single(rows.Components);
        }

        // ###########################################################################################
        // WHAT AN AMENDMENT MAY CHANGE. The table's nine sheets come from the edit; the highlights,
        // the calibrations and the revision date - none of which the table shows - always stay as
        // the submission has them, whatever the request carries.
        // ###########################################################################################
        [Fact]
        public void An_amendment_takes_the_tables_sheets_and_keeps_everything_the_table_does_not_show()
        {
            SubmissionRows current = ReviewTableContractTests.Rows();
            SubmissionRows edited = ReviewTableContractTests.Rows("251715-01");
            edited.RevisionDate = "1999-January-01";
            edited.ComponentHighlights.Clear();
            edited.KiCadCalibrations.Clear();
            edited.Credits.Add(new CreditEntry { Category = "Board", NameOrHandle = "Anna" });

            SubmissionRows amended = SubmissionRowsBoard.WithTableSections(current, edited);

            Assert.Equal("251715-01", Assert.Single(amended.Components).PartNumber);
            Assert.Equal(2, amended.Credits.Count);

            Assert.Equal("2026-September-25", amended.RevisionDate);
            Assert.Single(amended.ComponentHighlights);
            Assert.Single(amended.KiCadCalibrations);
        }

        // ###########################################################################################
        // *** HIGHLIGHTS AND CALIBRATIONS FOLLOW THE SCHEMATIC THEY ARE DRAWN ON (code review,
        // 2026-09-25). *** Kept as submitted, the ones on a schematic the reviewer deleted named a
        // schematic that no longer existed, the validator refused the save, and nothing in the table
        // could fix it.
        // ###########################################################################################
        private static SubmissionRows TwoSheets()
        {
            SubmissionRows rows = ReviewTableContractTests.Rows();
            rows.Schematics.Add(new BoardSchematicEntry { SchematicName = "Sheet 2", SchematicImageFile = "a/sheet2.png" });
            rows.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 2", BoardLabel = "U8", X = "5", Y = "6", Width = "7", Height = "8" });
            rows.KiCadCalibrations.Clear();
            rows.KiCadCalibrations.Add(new KiCadCalibrationEntry { SchematicName = "Sheet 1", OffsetX = 1 });
            rows.KiCadCalibrations.Add(new KiCadCalibrationEntry { SchematicName = "Sheet 2", OffsetX = 2 });
            return rows;
        }

        [Fact]
        public void Deleting_a_schematic_drops_the_highlights_and_calibrations_drawn_on_it()
        {
            SubmissionRows current = ReviewTableContractTests.TwoSheets();
            SubmissionRows edited = ReviewTableContractTests.TwoSheets();
            edited.Schematics.RemoveAll(schematic => schematic.SchematicName == "Sheet 2");

            SubmissionRows amended = SubmissionRowsBoard.WithTableSections(current, edited);

            Assert.Equal("Sheet 1", Assert.Single(amended.ComponentHighlights).SchematicName);
            Assert.Equal("Sheet 1", Assert.Single(amended.KiCadCalibrations).SchematicName);
        }

        // A rename in the table is a deleted row plus an added one; the same image says it is one
        // schematic, and what is drawn on it moves with it.
        [Fact]
        public void Renaming_a_schematic_moves_its_highlights_and_calibrations_to_the_new_name()
        {
            SubmissionRows current = ReviewTableContractTests.TwoSheets();
            SubmissionRows edited = ReviewTableContractTests.TwoSheets();
            edited.Schematics[1] = new BoardSchematicEntry { SchematicName = "Sheet 2 (PAL)", SchematicImageFile = "a/sheet2.png" };

            SubmissionRows amended = SubmissionRowsBoard.WithTableSections(current, edited);

            ComponentHighlightEntry moved = Assert.Single(amended.ComponentHighlights, highlight => highlight.X == "5");
            Assert.Equal("Sheet 2 (PAL)", moved.SchematicName);
            Assert.Equal("8", moved.Height);
            Assert.Equal("U8", moved.BoardLabel);
            Assert.Equal("Sheet 2 (PAL)", Assert.Single(amended.KiCadCalibrations, calibration => calibration.OffsetX == 2).SchematicName);

            // The untouched schematic's are left exactly as they were.
            Assert.Equal("Sheet 1", Assert.Single(amended.ComponentHighlights, highlight => highlight.X == "1").SchematicName);
        }

        // Only an unambiguous pair is a rename: two new rows on the removed one's image could each
        // be it, and a highlight on the wrong one is worse than none.
        [Fact]
        public void An_ambiguous_rename_drops_rather_than_guesses()
        {
            SubmissionRows current = ReviewTableContractTests.TwoSheets();
            SubmissionRows edited = ReviewTableContractTests.TwoSheets();
            edited.Schematics[1] = new BoardSchematicEntry { SchematicName = "Sheet 2a", SchematicImageFile = "a/sheet2.png" };
            edited.Schematics.Add(new BoardSchematicEntry { SchematicName = "Sheet 2b", SchematicImageFile = "a/sheet2.png" });

            SubmissionRows amended = SubmissionRowsBoard.WithTableSections(current, edited);

            Assert.DoesNotContain(amended.ComponentHighlights, highlight => highlight.X == "5");
        }

        // A name is matched ignoring case, as the validator matches it.
        [Fact]
        public void A_schematic_kept_in_another_case_keeps_what_is_drawn_on_it()
        {
            SubmissionRows current = ReviewTableContractTests.TwoSheets();
            SubmissionRows edited = ReviewTableContractTests.TwoSheets();
            edited.Schematics[1] = new BoardSchematicEntry { SchematicName = "SHEET 2", SchematicImageFile = "a/sheet2.png" };

            SubmissionRows amended = SubmissionRowsBoard.WithTableSections(current, edited);

            Assert.Equal(2, amended.ComponentHighlights.Count);
        }

        // ###########################################################################################
        // A DELETED COMPONENT'S HIGHLIGHTS (2026-09-25): an amendment may drop a highlight only when
        // no component in the edit has its label - the table sends them without it when a component
        // is deleted. Anything else a request leaves out is kept, so the table's route still cannot
        // remove or change a highlight of a component the board has.
        // ###########################################################################################
        private static SubmissionRows TwoComponents()
        {
            SubmissionRows rows = ReviewTableContractTests.Rows();
            rows.Components.Add(new ComponentEntry { BoardLabel = "U9", FriendlyName = "CIA" });
            rows.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U9", X = "5", Y = "6", Width = "7", Height = "8" });
            return rows;
        }

        [Fact]
        public void The_highlights_of_a_component_the_edit_deleted_and_dropped_are_removed()
        {
            SubmissionRows current = ReviewTableContractTests.TwoComponents();
            SubmissionRows edited = ReviewTableContractTests.TwoComponents();
            edited.Components.RemoveAll(component => component.BoardLabel == "U8");
            edited.ComponentHighlights.RemoveAll(highlight => highlight.BoardLabel == "U8");

            SubmissionRows amended = SubmissionRowsBoard.WithTableSections(current, edited);

            Assert.Equal("U9", Assert.Single(amended.ComponentHighlights).BoardLabel);
        }

        [Fact]
        public void A_request_cannot_drop_the_highlights_of_a_component_still_on_the_board()
        {
            SubmissionRows current = ReviewTableContractTests.TwoComponents();
            SubmissionRows edited = ReviewTableContractTests.TwoComponents();
            edited.ComponentHighlights.Clear();

            Assert.Equal(2, SubmissionRowsBoard.WithTableSections(current, edited).ComponentHighlights.Count);
        }

        // A component renamed in the table keeps its highlights: the table sends them, so they stay.
        [Fact]
        public void A_renamed_components_highlights_are_kept_when_the_table_sends_them()
        {
            SubmissionRows current = ReviewTableContractTests.TwoComponents();
            SubmissionRows edited = ReviewTableContractTests.TwoComponents();
            edited.Components[0] = new ComponentEntry { BoardLabel = "U10", FriendlyName = "PLA" };

            Assert.Equal(2, SubmissionRowsBoard.WithTableSections(current, edited).ComponentHighlights.Count);
        }

        // Every sheet the table shows is one an amendment takes - pinned against the schema, so a
        // sheet added to the table and forgotten here fails rather than silently not saving.
        [Fact]
        public void Every_sheet_the_table_shows_is_taken_from_the_edit()
        {
            SubmissionRows current = new();
            SubmissionRows edited = ReviewTableContractTests.Rows();

            BoardData amended = SubmissionRowsBoard.ToBoard(SubmissionRowsBoard.WithTableSections(current, edited));

            foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
            {
                Assert.True(
                    BoardWorkbookSchema.BuildRows(sheet, amended).Count > 0,
                    $"The table's [{sheet.SheetName}] sheet is not taken from an amendment.");
            }
        }

        // ---- on the wire ---------------------------------------------------------------------

        [Fact]
        public void The_table_travels_as_the_same_record_both_ends_use()
        {
            var sent = new ReviewTableData(2, ReviewTableContractTests.Rows(), ReviewTableContractTests.Rows("251715-01"));

            ReviewTableData? read = JsonSerializer.Deserialize<ReviewTableData>(
                JsonSerializer.Serialize(sent, JsonSerializerOptions.Web),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            Assert.Equal(2, read!.Version);
            Assert.Equal("251715-01", read.Submitted.Components.Single().PartNumber);
            Assert.Single(read.Published!.Schematics);
        }

        // The server answers { "amendedByReviewer": true }; CRT reads the status with web defaults.
        [Fact]
        public void The_amended_flag_reaches_CRT_under_the_name_the_server_writes()
        {
            string server = JsonSerializer.Serialize(new { state = "merged", amendedByReviewer = true }, JsonSerializerOptions.Web);

            SubmissionStatus? status = JsonSerializer.Deserialize<SubmissionStatus>(server, JsonSerializerOptions.Web);

            Assert.True(status!.AmendedByReviewer);
        }
    }
}
