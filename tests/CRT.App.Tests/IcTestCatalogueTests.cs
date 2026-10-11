using Handlers.IcTesting;

namespace ClassicRepairToolbox.Tests;

// IcTestCatalogue's reading and matching, without its cache: ReadEntries, IndexByMatch and Find
// are what Load and Lookup are made of. The shipped catalogue itself is checked in
// IcTestCatalogueShippedDataTests.
public sealed class IcTestCatalogueTests
{
    private static string Entry(string id, params string[] match) =>
        $$"""{ "schema": 1, "id": "{{id}}", "match": [{{string.Join(", ", match.Select(m => $"\"{m}\""))}}], "support": "testable" }""";

    [Fact]
    public void Every_json_file_but_the_index_is_an_entry()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile("74LS00.json", Entry("74ls00", "74LS00"));
        workspace.WriteFile("74LS04.json", Entry("74ls04", "74LS04"));
        workspace.WriteFile("index.json", """{ "schema": 1, "entries": [] }""");
        workspace.WriteFile("readme.txt", "not an entry");

        List<IcTestEntry> entries = IcTestCatalogue.ReadEntries(workspace.Root);

        Assert.Equal(new[] { "74ls00", "74ls04" }, entries.Select(e => e.Id));
    }

    [Fact]
    public void A_malformed_file_is_skipped_and_the_rest_still_load()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile("74LS00.json", Entry("74ls00", "74LS00"));
        workspace.WriteFile("74LS02.json", "{ \"id\": \"74ls02\", ");
        workspace.WriteFile("74LS04.json", Entry("74ls04", "74LS04"));

        List<IcTestEntry> entries = IcTestCatalogue.ReadEntries(workspace.Root);

        Assert.Equal(new[] { "74ls00", "74ls04" }, entries.Select(e => e.Id));
    }

    [Fact]
    public void An_entry_without_an_id_is_skipped()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile("74LS00.json", Entry(" ", "74LS00"));

        Assert.Empty(IcTestCatalogue.ReadEntries(workspace.Root));
    }

    [Fact]
    public void A_missing_folder_is_an_empty_catalogue()
    {
        using var workspace = new TempWorkspace();

        Assert.Empty(IcTestCatalogue.ReadEntries(Path.Combine(workspace.Root, "Catalogue")));
    }

    [Fact]
    public void A_technical_name_is_found_whatever_its_capitals_and_surrounding_spaces()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile("74LS06.json", Entry("74ls06", "74LS06", "7406"));
        var index = IcTestCatalogue.IndexByMatch(IcTestCatalogue.ReadEntries(workspace.Root));

        Assert.Equal("74ls06", IcTestCatalogue.Find(index, "  74ls06 ")?.Id);
        Assert.Equal("74ls06", IcTestCatalogue.Find(index, "7406")?.Id);
        Assert.Null(IcTestCatalogue.Find(index, "74LS07"));
        Assert.Null(IcTestCatalogue.Find(index, "   "));
        Assert.Null(IcTestCatalogue.Find(index, null));
    }

    [Fact]
    public void Two_entries_naming_one_part_leave_it_with_the_first_in_file_name_order()
    {
        using var workspace = new TempWorkspace();
        workspace.WriteFile("B.json", Entry("second", "7400"));
        workspace.WriteFile("A.json", Entry("first", "7400"));

        var index = IcTestCatalogue.IndexByMatch(IcTestCatalogue.ReadEntries(workspace.Root));

        Assert.Equal("first", IcTestCatalogue.Find(index, "7400")?.Id);
    }
}
