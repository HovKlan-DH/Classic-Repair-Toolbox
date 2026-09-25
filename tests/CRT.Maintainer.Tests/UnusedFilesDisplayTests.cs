using System.Text.Json;
using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer.Tests;

// Covers UnusedFilesDisplay and the parsing behind it - the administrator's "Unused files" window
// (owner decision, 2026-09-25: "there must be no orphan files").
//
// The one rule here that is a decision rather than words is CanRemove: a complete list, with
// something on it, AND the administrator's own tick. The first clean-up of a tree is looked at
// before anything goes.
public sealed class UnusedFilesDisplayTests
{
    private static UnusedFileListing Listing(bool complete = true, params UnusedFileEntry[] files) =>
        new("beta", complete, complete ? [] : ["The master workbook [x] lists [y], which is not in the data."], 2, 22, 10921, files);

    [Fact]
    public void The_summary_names_the_count_the_size_and_what_was_checked()
    {
        UnusedFileListing listing = UnusedFilesDisplayTests.Listing(true,
            new UnusedFileEntry("Generic shared files/Component images/7408.jpg", 3 * 1024 * 1024),
            new UnusedFileEntry("Commodore/Shared files/readme.txt", 512 * 1024));

        Assert.Equal(
            "2 files in the BETA data that nothing uses, 3.5 MB (10,921 files checked against 2 master workbooks and 22 board workbooks).",
            UnusedFilesDisplay.Summary(listing));
    }

    [Fact]
    public void A_clean_tree_says_so()
    {
        Assert.StartsWith("Nothing in the BETA data is unused (", UnusedFilesDisplay.Summary(UnusedFilesDisplayTests.Listing()));
    }

    // "Could not check" must never read as "nothing unused".
    [Fact]
    public void An_incomplete_list_says_nothing_can_be_removed_and_why()
    {
        string summary = UnusedFilesDisplay.Summary(UnusedFilesDisplayTests.Listing(complete: false));

        Assert.Contains("nothing can be removed", summary, StringComparison.Ordinal);
        Assert.Contains("which is not in the data", summary, StringComparison.Ordinal);
    }

    [Fact]
    public void Remove_needs_a_complete_list_with_something_on_it_AND_the_tick()
    {
        UnusedFileListing some = UnusedFilesDisplayTests.Listing(true, new UnusedFileEntry("a.png", 10));

        Assert.True(UnusedFilesDisplay.CanRemove(some, lookedThrough: true));
        Assert.False(UnusedFilesDisplay.CanRemove(some, lookedThrough: false));
        Assert.False(UnusedFilesDisplay.CanRemove(UnusedFilesDisplayTests.Listing(), lookedThrough: true));
        Assert.False(UnusedFilesDisplay.CanRemove(UnusedFilesDisplayTests.Listing(false, new UnusedFileEntry("a.png", 10)), lookedThrough: true));
        Assert.False(UnusedFilesDisplay.CanRemove(null, lookedThrough: true));
    }

    [Fact]
    public void The_button_says_how_many_files_and_how_much_goes()
    {
        Assert.Equal("Remove 1 file (812 bytes)", UnusedFilesDisplay.RemoveButton(UnusedFilesDisplayTests.Listing(true, new UnusedFileEntry("a.png", 812))));
        Assert.Equal("Remove unused files", UnusedFilesDisplay.RemoveButton(null));
    }

    [Theory]
    [InlineData(0, "0 bytes")]
    [InlineData(1, "1 byte")]
    [InlineData(1023, "1023 bytes")]
    [InlineData(1536, "1.5 KB")]
    [InlineData(9122611, "8.7 MB")]
    public void Sizes_read_the_same_in_every_locale(long bytes, string expected)
    {
        Assert.Equal(expected, UnusedFilesDisplay.Size(bytes));
    }

    [Fact]
    public void Production_is_named_as_production()
    {
        Assert.Equal("production", UnusedFilesDisplay.TreeName("production"));
        Assert.Equal("the BETA data", UnusedFilesDisplay.TreeName("beta"));
    }

    [Fact]
    public void The_result_says_what_went_and_what_was_kept()
    {
        Assert.Equal(
            "2 files removed from the BETA data. 1 file kept - something uses it again, or it was already gone.",
            UnusedFilesDisplay.Result(new UnusedFileRemovalResult("beta", ["a", "b"], ["c"], null)));

        Assert.Equal(
            "Nothing was removed: Only an administrator can remove unused files.",
            UnusedFilesDisplay.Result(new UnusedFileRemovalResult("beta", [], ["a"], "Only an administrator can remove unused files.")));
    }

    // CRT.Data's record, serialised the way the server does and read back through the parser.
    [Fact]
    public void The_list_is_read_as_the_servers_own_record()
    {
        string json = JsonSerializer.Serialize(
            UnusedFilesDisplayTests.Listing(true, new UnusedFileEntry("Generic shared files/Component images/7408.jpg", 42)),
            JsonSerializerOptions.Web);

        UnusedFileListing? listing = ReviewApiParser.ParseUnusedFiles(json);

        Assert.Equal("beta", listing!.Tree);
        Assert.Equal(10921, listing.FileCount);
        Assert.Equal(42, Assert.Single(listing.Files).SizeBytes);
    }

    [Fact]
    public void The_removal_answer_is_read()
    {
        UnusedFileRemovalResult? result = ReviewApiParser.ParseUnusedFileRemoval(
            """{"tree":"production","removed":["a.png"],"kept":[],"notDoneBecause":null}""");

        Assert.Equal("production", result!.Tree);
        Assert.Equal(["a.png"], result.Removed);
        Assert.Null(result.NotDoneBecause);
    }
}
