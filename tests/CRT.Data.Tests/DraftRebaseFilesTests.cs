using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The file lists' three-way merge when a 4.0 draft is brought up to date (DraftRebaseFiles):
// BASE as copied, DRAFT now, PUBLISHED now, by name. A file the contributor left as copied
// follows the update; one they replaced is theirs, and so are the rows citing it.
// ###########################################################################################
public sealed class DraftRebaseFilesTests
{
    private static BoardDocument.StoredFileRow File(string name, string content) => StoredFiles.EntryFor(name, content);

    private static List<BoardDocument.StoredFileRow> Merge(
        BoardDocument.StoredFileRow[] based,
        BoardDocument.StoredFileRow[] draft,
        BoardDocument.StoredFileRow[] published,
        out int followed,
        params string[] arrived) =>
        DraftRebaseFiles.Merge(based, draft, published, DraftRebaseFiles.Replaced(based, draft, name => arrived.Contains(name)), out followed);

    private static string ContentOf(List<BoardDocument.StoredFileRow> files, string name) =>
        files.Single(file => file.Name == name).Sha256;

    [Fact]
    public void A_file_left_as_copied_follows_the_update()
    {
        var merged = Merge([File("A.png", "v1")], [File("A.png", "v1")], [File("A.png", "v2")], out int followed);

        Assert.Equal(File("A.png", "v2").Sha256, ContentOf(merged, "A.png"));
        Assert.Equal(1, followed);
    }

    [Fact]
    public void A_file_the_contributor_replaced_stays_theirs()
    {
        var merged = Merge([File("A.png", "v1")], [File("A.png", "mine")], [File("A.png", "v2")], out int followed);

        Assert.Equal(File("A.png", "mine").Sha256, ContentOf(merged, "A.png"));
        Assert.Equal(0, followed);
    }

    // A new file lying in the draft folder under the name is the contributor's replacement, not
    // yet taken in - the list keeps the draft's entry, and the draft's write takes the new file.
    [Fact]
    public void A_file_replaced_by_a_new_file_not_yet_taken_in_is_not_followed()
    {
        var merged = Merge([File("A.png", "v1")], [File("A.png", "v1")], [File("A.png", "v2")], out int followed, "A.png");

        Assert.Equal(File("A.png", "v1").Sha256, ContentOf(merged, "A.png"));
        Assert.Equal(0, followed);
    }

    [Fact]
    public void A_file_the_update_added_comes_in_and_one_the_contributor_removed_stays_removed()
    {
        var merged = Merge(
            [File("A.png", "a")],
            [],
            [File("A.png", "a"), File("B.png", "new")],
            out int followed);

        Assert.Equal(["B.png"], merged.Select(file => file.Name));
        Assert.Equal(1, followed);
    }

    [Fact]
    public void A_file_left_as_copied_that_the_update_removed_goes_and_a_replaced_one_stays()
    {
        var merged = Merge(
            [File("A.png", "a"), File("B.png", "b")],
            [File("A.png", "a"), File("B.png", "mine")],
            [],
            out int followed);

        Assert.Equal(["B.png"], merged.Select(file => file.Name));
        Assert.Equal(1, followed);
    }

    // Both sides added a file of one name: the contributor's, as for a row both changed.
    [Fact]
    public void A_file_both_sides_added_under_one_name_is_the_contributors()
    {
        var merged = Merge([], [File("A.png", "mine")], [File("A.png", "theirs")], out int followed);

        Assert.Equal(File("A.png", "mine").Sha256, ContentOf(merged, "A.png"));
        Assert.Equal(0, followed);
    }

    [Fact]
    public void Names_are_matched_without_regard_to_capitals()
    {
        var merged = Merge([File("Sheet.PNG", "v1")], [File("sheet.png", "v1")], [File("SHEET.png", "v2")], out _);

        Assert.Equal(File("SHEET.png", "v2").Sha256, Assert.Single(merged).Sha256);
    }

    // ###########################################################################################
    // A replaced file is a change to the rows citing it: the base's citing rows are told apart, so
    // an update deleting the row cannot take the contributor's picture with it - while every other
    // row is left exactly as it was, for the merge to compare as before.
    // ###########################################################################################
    [Fact]
    public void The_bases_rows_citing_a_replaced_file_are_told_apart_and_no_others()
    {
        var based = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "Main.png" }],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "U8.png" },
                new ComponentImageEntry { BoardLabel = "U9", Name = "Top", File = "U9.png" },
            ],
            BoardLocalFiles = [new BoardLocalFileEntry { Name = "Manual", File = "u8.PNG" }],
        };

        BoardData marked = DraftRebaseFiles.MarkReplaced(based, new HashSet<string>(["U8.png"], StringComparer.OrdinalIgnoreCase));

        Assert.NotEqual("U8.png", marked.ComponentImages[0].File);
        Assert.Equal("Top", marked.ComponentImages[0].Name);
        Assert.Same(based.ComponentImages[1], marked.ComponentImages[1]);
        Assert.NotEqual("u8.PNG", marked.BoardLocalFiles[0].File);
        Assert.Same(based.Schematics[0], marked.Schematics[0]);

        // The base itself is not changed.
        Assert.Equal("U8.png", based.ComponentImages[0].File);
    }
}
