using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardPartRequired - a file added to the board comes with every field it needs filled in (owner request,
// 2026-10-30). What each row of the board's own lists needs is CRT.Data's BoardRowRequired, tested there and
// through BoardDataChecks; the headings this class once named them by went with their last use (2026-11-15).
// ###########################################################################################
public sealed class BoardPartRequiredTests
{
    // "Add file" makes a row with every field filled in - its file is copied into the draft at once, and
    // a row held back would leave it cited by nothing: the last row's category, else "Resources".
    [Fact]
    public void A_file_added_has_every_field_filled_in()
    {
        string source = Path.Combine(Path.GetTempPath(), "Picked", "C64 service manual.pdf");

        BoardLocalFileEntry first = BoardPartRequired.NewFile(source, "Board local files/C64 service manual.pdf", lastCategory: null);
        Assert.Equal(("Resources", "C64 service manual", "Board local files/C64 service manual.pdf"), (first.Category, first.Name, first.File));
        Assert.Empty(BoardRowRequired.EmptyColumns(first));

        Assert.Equal("Troubleshooting", BoardPartRequired.NewFile(source, "x.pdf", " Troubleshooting ").Category);
        Assert.Equal("Resources", BoardPartRequired.NewFile(source, "x.pdf", "  ").Category);

        // A file named only by its extension keeps that as its name.
        Assert.Equal(".pdf", BoardPartRequired.NewFile(Path.Combine(Path.GetTempPath(), ".pdf"), "x.pdf", null).Name);
    }
}
