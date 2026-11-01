using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardPartRequired - what a row of the board's own lists must have filled in (owner request,
// 2026-10-30): a file and a link all three fields, a credit "What" and "Name or handle", an important
// signal its net and its name - CRT.Data's BoardRowRequired, named by the component pane's headings.
// ###########################################################################################
public sealed class BoardPartRequiredTests
{
    [Fact]
    public void A_file_needs_its_category_its_name_and_its_file()
    {
        Assert.Empty(BoardPartRequired.EmptyOf(new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf" }));
        Assert.Equal(["Category", "Name", "File"], BoardPartRequired.EmptyOf(new BoardLocalFileEntry()));

        // Blanks are empty too.
        Assert.Equal(["Name"], BoardPartRequired.EmptyOf(new BoardLocalFileEntry { Category = "Resources", Name = "  ", File = "x.pdf" }));
    }

    [Fact]
    public void A_link_needs_its_category_its_name_and_its_address()
    {
        Assert.Empty(BoardPartRequired.EmptyOf(new BoardLinkEntry { Category = "Forums", Name = "Lemon64", Url = "https://www.lemon64.com" }));
        Assert.Equal(["Name", "Address"], BoardPartRequired.EmptyOf(new BoardLinkEntry { Category = "Forums" }));
        Assert.Equal(["Category"], BoardPartRequired.EmptyOf(new BoardLinkEntry { Name = "Lemon64", Url = "https://www.lemon64.com" }));
    }

    // Which part of it and the contact may stay empty - the live boards leave them so on many credits.
    [Fact]
    public void A_credit_needs_what_it_is_for_and_a_name_or_handle_only()
    {
        Assert.Empty(BoardPartRequired.EmptyOf(new CreditEntry { Category = "Schematic image", NameOrHandle = "Anna" }));
        Assert.Equal(["What", "Name or handle"], BoardPartRequired.EmptyOf(new CreditEntry { SubCategory = "Top", Contact = "anna@example.com" }));
        Assert.Equal(["Name or handle"], BoardPartRequired.EmptyOf(new CreditEntry { Category = "Board data" }));
    }

    [Fact]
    public void An_important_signal_needs_its_net_and_its_name()
    {
        Assert.Empty(BoardPartRequired.EmptyOf(new KiCadImportantSignalEntry { DisplayName = "Reset", KiCadNetName = "/RESET" }));
        Assert.Equal(["Important signal name"], BoardPartRequired.EmptyOf(new KiCadImportantSignalEntry { KiCadNetName = "/RESET" }));
        Assert.Equal(["KiCad net", "Important signal name"], BoardPartRequired.EmptyOf(new KiCadImportantSignalEntry()));
    }

    // "Add file" makes a row with every field filled in - its file is copied into the draft at once, and
    // a row held back would leave it cited by nothing: the last row's category, else "Resources".
    [Fact]
    public void A_file_added_has_every_field_filled_in()
    {
        string source = Path.Combine(Path.GetTempPath(), "Picked", "C64 service manual.pdf");

        BoardLocalFileEntry first = BoardPartRequired.NewFile(source, "Board local files/C64 service manual.pdf", lastCategory: null);
        Assert.Equal(("Resources", "C64 service manual", "Board local files/C64 service manual.pdf"), (first.Category, first.Name, first.File));
        Assert.Empty(BoardPartRequired.EmptyOf(first));

        Assert.Equal("Troubleshooting", BoardPartRequired.NewFile(source, "x.pdf", " Troubleshooting ").Category);
        Assert.Equal("Resources", BoardPartRequired.NewFile(source, "x.pdf", "  ").Category);

        // A file named only by its extension keeps that as its name.
        Assert.Equal(".pdf", BoardPartRequired.NewFile(Path.Combine(Path.GetTempPath(), ".pdf"), "x.pdf", null).Name);
    }
}
