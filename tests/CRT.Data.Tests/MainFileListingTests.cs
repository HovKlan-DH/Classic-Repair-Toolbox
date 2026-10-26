using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// MainFileListing - the drop-down lists in the CRT 4.0 main file ("Classic-Repair-Toolbox.v4.0.0.json"),
// edited through MasterListing's own calls (2026-10-25): a new board's publish and promotion insert
// its row, a push-back or a deletion removes it, the Boards screen reorders. The rules are the Excel
// file's - a row is a board, names another board is listed under are refused, nothing written when
// nothing changes - and the oscilloscopes and every field this CRT does not know come through.
// ###########################################################################################
public sealed class MainFileListingTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private static readonly DateTimeOffset When = new(2026, 10, 25, 9, 0, 0, TimeSpan.Zero);

    private string MainPath => Path.Combine(this.thisWorkspace.Root, DataTreeLayout.MainFileName);

    private void WriteMain(params (string BoardId, string Hardware, string Board)[] boards)
    {
        var document = new MainDocument
        {
            BoardsRevisionDate = "2026-October-1",
            OscilloscopesRevisionDate = "2026-October-2",
            Boards = boards.Select(board => new MainDocument.BoardListingRow { BoardId = board.BoardId, HardwareName = board.Hardware, BoardName = board.Board }).ToList(),
            Oscilloscopes = [new MainDocument.OscilloscopeRow { Brand = "Rigol", SeriesOrModel = "DS1000Z", Port = "5555" }],
            Unknown = new Dictionary<string, JsonElement> { ["laterField"] = JsonDocument.Parse("\"kept\"").RootElement.Clone() }
        };

        File.WriteAllText(this.MainPath, DataTreeJson.Write(document));
    }

    private MainDocument ReadMain() => DataTreeJson.ReadMain(File.ReadAllText(this.MainPath));

    private static MasterListingRow Row(string boardId, string hardware, string board, string notes = "") =>
        new(hardware, board, DataTreeLayout.BoardFilePath(boardId), notes);

    [Fact]
    public void The_main_file_is_the_newest_master_of_a_4_0_tree()
    {
        this.WriteMain();

        Assert.Equal(this.MainPath, MasterListing.NewestMasterPath(this.thisWorkspace.Root));
    }

    [Fact]
    public void Its_rows_are_read_as_boards_naming_their_board_files()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"), ("Commodore/C128/310378", "", "310378"));

        Assert.True(MasterListing.TryRead(this.MainPath, out IReadOnlyList<MasterListingRow> rows, out string why), why);

        Assert.Equal("Commodore/C64/250407/board.json", rows[0].ExcelDataFile);
        Assert.Equal("Commodore/C64/250407", rows[0].BoardId);

        // A blank hardware name takes the one above, as the Excel sheet's did.
        Assert.Equal("Commodore 64", rows[1].HardwareName);
    }

    [Fact]
    public void A_main_file_that_cannot_be_read_says_why_and_lists_nothing()
    {
        File.WriteAllText(this.MainPath, "{ not json");

        Assert.False(MasterListing.TryRead(this.MainPath, out IReadOnlyList<MasterListingRow> rows, out string why));
        Assert.Empty(rows);
        Assert.NotEmpty(why);
    }

    [Fact]
    public void A_new_board_goes_after_the_row_it_was_placed_after()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"), ("Commodore/C128/310378", "Commodore 128", "310378"));

        MasterListingEdit edit = MasterListing.Insert(
            this.MainPath,
            MainFileListingTests.Row("Commodore/C64/250425", "Commodore 64", "250425", "Short board."),
            "Commodore/C64/250407/board.json",
            MainFileListingTests.When);

        Assert.True(edit.IsDone, edit.Failure);
        Assert.True(edit.Changed);
        Assert.Equal(
            ["Commodore/C64/250407", "Commodore/C64/250425", "Commodore/C128/310378"],
            this.ReadMain().Boards.Select(board => board.BoardId));
        Assert.Equal("Short board.", this.ReadMain().Boards[1].HardwareNotes);
    }

    // A placement saved before the 4.0 tree names the board it goes after by its WORKBOOK - the
    // board it names is found all the same.
    [Fact]
    public void A_placement_naming_the_board_before_by_its_old_workbook_still_finds_it()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"));

        MasterListingEdit edit = MasterListing.Insert(
            this.MainPath,
            MainFileListingTests.Row("Commodore/C64/250425", "Commodore 64", "250425"),
            "Commodore/C64/250407/Data C64 250407 v3.0.0.xlsx",
            MainFileListingTests.When);

        Assert.True(edit.IsDone, edit.Failure);
        Assert.Equal("Commodore/C64/250425", this.ReadMain().Boards[1].BoardId);
    }

    [Fact]
    public void A_board_placed_with_nothing_before_it_goes_first()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"));

        MasterListing.Insert(this.MainPath, MainFileListingTests.Row("Amstrad/CPC 464/Z70200", "Amstrad CPC 464", "Z70200"), null, MainFileListingTests.When);

        Assert.Equal("Amstrad/CPC 464/Z70200", this.ReadMain().Boards[0].BoardId);
    }

    [Fact]
    public void A_row_placed_after_a_board_no_longer_listed_is_refused_and_nothing_is_written()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"));
        string before = File.ReadAllText(this.MainPath);

        MasterListingEdit edit = MasterListing.Insert(
            this.MainPath,
            MainFileListingTests.Row("Commodore/C64/250425", "Commodore 64", "250425"),
            "Commodore/VIC-20/325078/board.json",
            MainFileListingTests.When);

        Assert.False(edit.IsDone);
        Assert.Equal(before, File.ReadAllText(this.MainPath));
    }

    [Fact]
    public void Names_another_board_is_listed_under_are_refused()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"));

        MasterListingEdit edit = MasterListing.Insert(
            this.MainPath,
            MainFileListingTests.Row("Commodore/C64/250407b", "commodore 64", "250407"),
            null,
            MainFileListingTests.When);

        Assert.False(edit.IsDone);
        Assert.Contains("Commodore/C64/250407", edit.Failure, StringComparison.Ordinal);
    }

    // Re-running an interrupted publish is safe: the row is the board's, updated where it is.
    [Fact]
    public void Inserting_a_listed_board_updates_its_row_in_place_and_an_identical_one_writes_nothing()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"), ("Commodore/C128/310378", "Commodore 128", "310378"));

        MasterListingEdit same = MasterListing.Insert(
            this.MainPath, MainFileListingTests.Row("Commodore/C64/250407", "Commodore 64", "250407"), null, MainFileListingTests.When);

        Assert.True(same.IsDone);
        Assert.False(same.Changed);

        MasterListingEdit renamed = MasterListing.Insert(
            this.MainPath, MainFileListingTests.Row("Commodore/C64/250407", "Commodore 64", "250407 (long board)"), null, MainFileListingTests.When);

        Assert.True(renamed.Changed);
        Assert.Equal(2, this.ReadMain().Boards.Count);
        Assert.Equal("250407 (long board)", this.ReadMain().Boards[0].BoardName);
    }

    [Fact]
    public void Removing_a_board_takes_its_row_out_and_an_unlisted_one_changes_nothing()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"), ("Commodore/C128/310378", "Commodore 128", "310378"));

        Assert.True(MasterListing.Remove(this.MainPath, "Commodore/C64/250407", MainFileListingTests.When).Changed);
        Assert.Equal(["Commodore/C128/310378"], this.ReadMain().Boards.Select(board => board.BoardId));

        MasterListingEdit again = MasterListing.Remove(this.MainPath, "Commodore/C64/250407", MainFileListingTests.When);

        Assert.True(again.IsDone);
        Assert.False(again.Changed);
    }

    [Fact]
    public void Reordering_moves_whole_rows_and_an_unchanged_order_writes_nothing()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"), ("Commodore/C128/310378", "Commodore 128", "310378"));

        Assert.False(MasterListing.Reorder(this.MainPath, ["Commodore/C64/250407", "Commodore/C128/310378"], MainFileListingTests.When).Changed);

        MasterListingEdit edit = MasterListing.Reorder(this.MainPath, ["Commodore/C128/310378", "Commodore/C64/250407"], MainFileListingTests.When);

        Assert.True(edit.Changed);
        Assert.Equal(["Commodore/C128/310378", "Commodore/C64/250407"], this.ReadMain().Boards.Select(board => board.BoardId));
        Assert.Equal("Commodore 128", this.ReadMain().Boards[0].HardwareName);
    }

    // ###########################################################################################
    // *** EVERY WRITE DATES BOTH LISTS, AND KEEPS WHAT IT DOES NOT EDIT (owner request, 2026-10-04,
    // for the Excel file's two sheets). *** The oscilloscopes and a field a later CRT added come
    // through every edit untouched.
    // ###########################################################################################
    [Fact]
    public void A_write_dates_both_lists_and_keeps_the_oscilloscopes_and_unknown_fields()
    {
        this.WriteMain(("Commodore/C64/250407", "Commodore 64", "250407"));

        MasterListing.Remove(this.MainPath, "Commodore/C64/250407", MainFileListingTests.When);

        MainDocument main = this.ReadMain();

        Assert.Equal("2026-October-25", main.BoardsRevisionDate);
        Assert.Equal("2026-October-25", main.OscilloscopesRevisionDate);
        Assert.Equal("DS1000Z", Assert.Single(main.Oscilloscopes).SeriesOrModel);
        Assert.Equal("kept", main.Unknown!["laterField"].GetString());
    }

    [Fact]
    public void A_row_with_a_blank_name_is_refused_before_the_file_is_opened()
    {
        this.WriteMain();

        Assert.False(MasterListing.Insert(this.MainPath, MainFileListingTests.Row("Commodore/C64/250407", " ", "250407"), null).IsDone);
    }
}
