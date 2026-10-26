using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Where things live in the CRT 4.0 data tree - above all, where a file's bytes are stored.
public sealed class DataTreeLayoutTests
{
    private const string Sha = "3f7a0c19b2d44e8f9a6b5c4d3e2f1a0b9c8d7e6f5a4b3c2d1e0f9a8b7c6d5e4f";

    [Fact]
    public void A_file_is_stored_under_its_checksum_in_a_folder_of_its_first_two_characters()
    {
        Assert.Equal($"Files/3f/{Sha}.pdf", DataTreeLayout.StoredFilePath(Sha, "Component local files/7406.pdf"));
    }

    // Two boards naming the same bytes ".PDF" and ".pdf" share ONE stored file.
    [Fact]
    public void The_extension_is_lower_cased_so_one_file_is_stored_once()
    {
        Assert.Equal(
            DataTreeLayout.StoredFilePath(Sha, "Datasheets/7406.PDF"),
            DataTreeLayout.StoredFilePath(Sha, "Component local files/7406.pdf"));
    }

    // KiCad's own extensions carry an underscore and must keep it - the web server and a person
    // looking at the folder both go by it.
    [Fact]
    public void A_KiCad_extension_is_kept()
    {
        Assert.Equal(".kicad_pcb", DataTreeLayout.StoredExtension("KiCad data/250407_.kicad_pcb"));
    }

    [Theory]
    [InlineData("README")]
    [InlineData("odd.p df")]
    [InlineData("odd.pdf!")]
    [InlineData("long.abcdefghijklmnopq")]
    [InlineData("trailing.")]
    public void An_extension_that_is_not_plain_letters_and_digits_is_left_off(string name)
    {
        Assert.Equal($"Files/3f/{Sha}", DataTreeLayout.StoredFilePath(Sha, name));
    }

    [Theory]
    [InlineData("3F7A0C19B2D44E8F9A6B5C4D3E2F1A0B9C8D7E6F5A4B3C2D1E0F9A8B7C6D5E4F")]
    [InlineData("3f7a")]
    [InlineData("")]
    [InlineData("../../etc/passwd/../../../../../../../../../../../../../../../..")]
    public void Anything_but_a_lower_case_SHA_256_is_refused_as_a_stored_name(string sha)
    {
        Assert.Throws<ArgumentException>(() => DataTreeLayout.StoredFilePath(sha, "x.png"));
    }

    [Fact]
    public void A_board_file_is_found_by_the_board_id()
    {
        Assert.Equal("Commodore/C64/250407/board.json", DataTreeLayout.BoardFilePath("Commodore/C64/250407"));
        Assert.Equal("ZX Spectrum/Spectrum 16K-48K/Issue 4B/board.json", DataTreeLayout.BoardFilePath(" ZX Spectrum/Spectrum 16K-48K/Issue 4B/ "));
    }

    // The identity CRT has always carried for a board - Excel-era or 4.0 - finds its board file.
    [Theory]
    [InlineData("Commodore/C64/250407/Data C64 250407 v3.0.0.xlsx", "Commodore/C64/250407/board.json")]
    [InlineData("Commodore/C64/250407/board.json", "Commodore/C64/250407/board.json")]
    [InlineData(@"\Commodore\C64\250407\Data C64 250407.xlsx", "Commodore/C64/250407/board.json")]
    [InlineData("board.json", "")]
    [InlineData("", "")]
    public void A_board_identity_finds_the_board_file_in_its_folder(string identity, string boardFile)
    {
        Assert.Equal(boardFile, DataTreeLayout.BoardFileFor(identity));
    }

    [Theory]
    [InlineData("Commodore/C64/250407/board.json", true)]
    [InlineData("BOARD.JSON", true)]
    [InlineData("Commodore/C64/250407/Data C64 250407.xlsx", false)]
    [InlineData("Commodore/C64/250407/Data C64 250407.json", false)]
    [InlineData(".crt-writing-board.json", false)]
    public void Only_a_board_json_is_read_as_a_4_0_board(string path, bool isBoardFile)
    {
        Assert.Equal(isBoardFile, DataTreeLayout.IsBoardFile(path));
    }
}
