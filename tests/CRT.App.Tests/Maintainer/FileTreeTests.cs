using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers FileTree and FileTreeWording - a system's files as a folder tree, with what the next step
// does to each (owner request, 2026-09-28: "a file-structure for all existing files ... a
// highlighting of files changed ... 'show only changed files' ... including their parent
// folder").
// ###########################################################################################
public sealed class FileTreeTests
{
    private static SystemFileEntry E(string path, SystemFileChange change = SystemFileChange.Unchanged) =>
        new(path, change, SystemFileSource.Beta);

    private static FileTreeNode Sample() => FileTree.Build(
    [
        FileTreeTests.E("Commodore/C128/310378/Images/C10.png"),
        FileTreeTests.E("Commodore/C128/310378/Images/C2.png", SystemFileChange.Changed),
        FileTreeTests.E("Commodore/C128/310378/Data.xlsx", SystemFileChange.Changed),
        FileTreeTests.E("Commodore/C128/310378/Manuals/a.pdf"),
        FileTreeTests.E("Commodore/Shared files/manual.pdf", SystemFileChange.Added),
        FileTreeTests.E("Generic shared files/7400.pdf")
    ]);

    // Folders first, then files, each in the order a person reads them - C2 before C10.
    [Fact]
    public void Folders_come_first_and_names_sort_naturally()
    {
        FileTreeNode root = FileTreeTests.Sample();

        Assert.Equal(["Commodore", "Generic shared files"], root.Children.Select(node => node.Name));

        FileTreeNode board = root.Children[0].Children.Single(node => node.Name == "C128").Children.Single();
        Assert.Equal(["Images", "Manuals", "Data.xlsx"], board.Children.Select(node => node.Name));
        Assert.Equal(["C2.png", "C10.png"], board.Children[0].Children.Select(node => node.Name));
    }

    // Every folder knows what changes under it, so a closed one can still say so.
    [Fact]
    public void A_folder_counts_the_changes_under_it()
    {
        FileTreeNode root = FileTreeTests.Sample();
        FileTreeNode commodore = root.Children[0];

        Assert.Equal(3, commodore.ChangeCount);
        Assert.Equal(1, commodore.Added);
        Assert.Equal(2, commodore.Changed);
        Assert.Equal(2, commodore.Unchanged);
        Assert.False(root.Children[1].HasChanges);
        Assert.Equal("3 changing", FileTreeWording.FolderNote(commodore));
        Assert.Equal(string.Empty, FileTreeWording.FolderNote(root.Children[1]));
    }

    // ###########################################################################################
    // *** "SHOW ONLY CHANGED FILES ... INCLUDING THEIR PARENT FOLDER." *** Every changed file under
    // ALL its folders - opened as the view starts (DefaultExpanded) - and nothing else.
    // ###########################################################################################
    [Fact]
    public void Only_changed_shows_the_changed_files_under_all_their_folders_open()
    {
        FileTreeNode root = FileTreeTests.Sample();
        IReadOnlyList<FileTreeRow> rows = FileTree.Rows(root, FileTree.DefaultExpanded(root), onlyChanged: true);

        Assert.Equal(
            ["Commodore", "C128", "310378", "Images", "C2.png", "Data.xlsx", "Shared files", "manual.pdf"],
            rows.Select(row => row.Node.Name));
        Assert.Equal([0, 1, 2, 3, 4, 3, 1, 2], rows.Select(row => row.Depth));
        Assert.All(rows.Where(row => row.Node.IsFolder), row => Assert.True(row.IsExpanded));
    }

    // ###########################################################################################
    // Folders close in that view too - "Collapse all" there leaves the top folders, each still
    // saying how many changes are in it (owner request, 2026-09-28).
    // ###########################################################################################
    [Fact]
    public void Only_changed_keeps_a_closed_folder_closed()
    {
        FileTreeNode root = FileTreeTests.Sample();

        FileTreeRow top = Assert.Single(FileTree.Rows(root, new HashSet<string>(), onlyChanged: true));
        Assert.Equal("Commodore", top.Node.Name);
        Assert.False(top.IsExpanded);

        var shut = new HashSet<string>(FileTree.DefaultExpanded(root));
        shut.Remove("Commodore/C128/310378/Images");

        Assert.Equal(
            ["Commodore", "C128", "310378", "Images", "Data.xlsx", "Shared files", "manual.pdf"],
            FileTree.Rows(root, shut, onlyChanged: true).Select(row => row.Node.Name));
    }

    // "Expand all" opens every folder, unchanged ones included.
    [Fact]
    public void Every_folder_is_found_for_expand_all()
    {
        FileTreeNode root = FileTreeTests.Sample();

        Assert.Equal(
            new HashSet<string>
            {
                "Commodore", "Commodore/C128", "Commodore/C128/310378", "Commodore/C128/310378/Images",
                "Commodore/C128/310378/Manuals", "Commodore/Shared files", "Generic shared files"
            },
            FileTree.AllFolders(root).ToHashSet());

        // Seven folders and six files: every row of the tree.
        Assert.Equal(13, FileTree.Rows(root, FileTree.AllFolders(root), onlyChanged: false).Count);
    }

    // Everything, a folder's contents only when it is open.
    [Fact]
    public void The_whole_tree_opens_only_the_folders_asked_for()
    {
        FileTreeNode root = FileTreeTests.Sample();

        Assert.Equal(["Commodore", "Generic shared files"], FileTree.Rows(root, new HashSet<string>(), onlyChanged: false).Select(row => row.Node.Name));

        IReadOnlyList<FileTreeRow> rows = FileTree.Rows(root, new HashSet<string> { "Generic shared files" }, onlyChanged: false);
        Assert.Equal(["Commodore", "Generic shared files", "7400.pdf"], rows.Select(row => row.Node.Name));
    }

    // ###########################################################################################
    // The folders open at first are those on the way to a change: the changes are in view, and a
    // folder of untouched images is not spread across the screen.
    // ###########################################################################################
    [Fact]
    public void At_first_only_the_folders_on_the_way_to_a_change_are_open()
    {
        IReadOnlySet<string> expanded = FileTree.DefaultExpanded(FileTreeTests.Sample());

        Assert.Equal(
            new HashSet<string> { "Commodore", "Commodore/C128", "Commodore/C128/310378", "Commodore/C128/310378/Images", "Commodore/Shared files" },
            expanded.ToHashSet());
    }

    [Fact]
    public void The_summary_counts_what_changes_and_what_stays()
    {
        Assert.Equal(
            "3 files change: 1 new, 2 changed, 0 removed. 3 files already the same.",
            FileTreeWording.Summary(FileTreeTests.Sample()));

        Assert.Equal(
            "Nothing changes - 1 file already the same.",
            FileTreeWording.Summary(FileTree.Build([FileTreeTests.E("a/b.png")])));

        Assert.Equal("No files.", FileTreeWording.Summary(FileTree.Build([])));
    }

    // ###########################################################################################
    // *** THE CARD SAYS NOTHING ABOUT THE CHANGE - only what opens, when it is not the file as it
    // will be (owner request, 2026-09-28). *** The workbook and the highlight file are written
    // from the table on approval, so before then only BETA's current copy opens - or nothing, for a
    // new system. Everything else says nothing beside its path.
    // ###########################################################################################
    [Fact]
    public void Only_a_file_the_approval_writes_says_what_opens()
    {
        Assert.Contains("BETA's copy as it is now",
            FileTreeWording.Note(new SystemFileEntry("a/Data.xlsx", SystemFileChange.Changed, SystemFileSource.Beta, WrittenOnApproval: true)),
            StringComparison.Ordinal);

        Assert.Contains("nothing to open yet",
            FileTreeWording.Note(new SystemFileEntry("a/Data.xlsx", SystemFileChange.Added, SystemFileSource.NotWrittenYet, WrittenOnApproval: true)),
            StringComparison.Ordinal);

        // Kept as it is, it IS the file: nothing to say.
        Assert.Null(FileTreeWording.Note(new SystemFileEntry("a/Data.json", SystemFileChange.Unchanged, SystemFileSource.Beta, WrittenOnApproval: true)));
        Assert.Null(FileTreeWording.Note(new SystemFileEntry("a/b.png", SystemFileChange.Added, SystemFileSource.Submission, "h")));
        Assert.Null(FileTreeWording.Note(new SystemFileEntry("a/b.png", SystemFileChange.Removed, SystemFileSource.Production)));
        Assert.Null(FileTreeWording.Note(new SystemFileEntry("a/b.png", SystemFileChange.Changed, SystemFileSource.Beta)));
    }

    [Fact]
    public void A_file_row_says_new_changed_or_removed_and_nothing_when_it_stays()
    {
        Assert.Equal("new", FileTreeWording.ChangeWord(SystemFileChange.Added));
        Assert.Equal("changed", FileTreeWording.ChangeWord(SystemFileChange.Changed));
        Assert.Equal("removed", FileTreeWording.ChangeWord(SystemFileChange.Removed));
        Assert.Equal(string.Empty, FileTreeWording.ChangeWord(SystemFileChange.Unchanged));
    }
}
