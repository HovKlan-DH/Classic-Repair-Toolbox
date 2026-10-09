using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers ReviewTableFiles - the table's file hover card in the Maintainer tab (owner
// request, 2026-09-26): which submitted file a path means, and what a file is saved as before the
// operating system opens it. The second is a security decision - these are a contributor's bytes.
// ###########################################################################################
public sealed class ReviewTableFilesTests
{
    private static SubmittedFileFact Fact(string path, string hash) =>
        new(path, hash, 10, SubmissionFileScope.Own, IsReferenced: true, PublishedSha256: null);

    // ###########################################################################################
    // A file opened from the table is written into the temp folder and handed to the operating
    // board, so it cannot be deleted while a viewer may still hold it - but nothing deleted them
    // at all, and months of reviewing piled up contributor files there (code review, 2026-09-26).
    // Swept on the next open, a day later.
    // ###########################################################################################
    [Fact]
    public void An_opened_files_folder_is_swept_only_once_it_is_a_day_old()
    {
        var now = new DateTimeOffset(2026, 9, 26, 12, 0, 0, TimeSpan.Zero);

        Assert.False(ReviewTableFiles.IsStaleOpenedFolder(now, now));
        Assert.False(ReviewTableFiles.IsStaleOpenedFolder(now.AddHours(-23), now));
        Assert.True(ReviewTableFiles.IsStaleOpenedFolder(now.AddHours(-25), now));
        Assert.Equal(TimeSpan.FromDays(1), ReviewTableFiles.OpenedFileLifetime);
    }

    [Fact]
    public void A_path_the_submission_carries_gives_its_submitted_hash()
    {
        IReadOnlyList<SubmittedFileFact> files =
        [
            ReviewTableFilesTests.Fact("Commodore/C64/250407/a.png", "hash-a"),
            ReviewTableFilesTests.Fact("Commodore/C64/250407/b.png", "hash-b"),
        ];

        Assert.Equal("hash-b", ReviewTableFiles.SubmittedHashFor(files, "Commodore/C64/250407/b.png"));
    }

    // Case-sensitive, like the server's tree: a path differing only by case is another file, and
    // falls back to the published side rather than showing the wrong picture.
    [Fact]
    public void A_path_the_submission_does_not_carry_has_no_hash()
    {
        IReadOnlyList<SubmittedFileFact> files = [ReviewTableFilesTests.Fact("Commodore/C64/250407/a.png", "hash-a")];

        Assert.Null(ReviewTableFiles.SubmittedHashFor(files, "Commodore/C64/250407/A.png"));
        Assert.Null(ReviewTableFiles.SubmittedHashFor(files, "Commodore/C64/250407/other.png"));
        Assert.Null(ReviewTableFiles.SubmittedHashFor(files, string.Empty));
        Assert.Null(ReviewTableFiles.SubmittedHashFor(null, "Commodore/C64/250407/a.png"));
    }

    // A published board: the current side is the submission's file, the published side the
    // published tree's (null = read the tree by path).
    [Fact]
    public void For_a_published_board_only_the_current_side_reads_the_submission()
    {
        IReadOnlyList<SubmittedFileFact> files = [ReviewTableFilesTests.Fact("Commodore/C64/250407/a.png", "hash-a")];

        Assert.Equal("hash-a", ReviewTableFiles.HashToRead(files, "Commodore/C64/250407/a.png", CRT.BoardTableFileSide.Current, nothingPublished: false));
        Assert.Null(ReviewTableFiles.HashToRead(files, "Commodore/C64/250407/a.png", CRT.BoardTableFileSide.Published, nothingPublished: false));
    }

    // ###########################################################################################
    // *** A NEW BOARD'S "PUBLISHED" SIDE IS THE SUBMISSION TOO. *** Its table is compared with the
    // submission as opened, and the published tree holds nothing of it - read from there, every
    // picture showed beside "There is no file at this path" (reported, 2026-09-26). Both sides
    // reading the same file is what lets the card show it once.
    // ###########################################################################################
    [Fact]
    public void For_a_new_board_both_sides_read_the_submission()
    {
        IReadOnlyList<SubmittedFileFact> files = [ReviewTableFilesTests.Fact("Manu1/Hardware1/Board1/1N4148.png", "hash-d")];

        Assert.Equal("hash-d", ReviewTableFiles.HashToRead(files, "Manu1/Hardware1/Board1/1N4148.png", CRT.BoardTableFileSide.Published, nothingPublished: true));
        Assert.Equal("hash-d", ReviewTableFiles.HashToRead(files, "Manu1/Hardware1/Board1/1N4148.png", CRT.BoardTableFileSide.Current, nothingPublished: true));

        // A file the maintainer has named but not saved is in neither.
        Assert.Null(ReviewTableFiles.HashToRead(files, "Manu1/Hardware1/Board1/other.png", CRT.BoardTableFileSide.Current, nothingPublished: true));
    }

    // Two sides are shown only when they differ - for a new board, when the maintainer has named
    // another file - and are then named for what they are. Nothing of it is "published". A
    // published board's older side is the BETA source, as its text cells' tooltips say.
    [Fact]
    public void A_new_boards_two_sides_are_the_submission_and_the_maintainers_change()
    {
        Assert.Equal(("As submitted", "Your change"), ReviewTableFiles.SideLabels(nothingPublished: true));
        Assert.Equal(("Before (BETA source)", "After (submitted)"), ReviewTableFiles.SideLabels(nothingPublished: false));
    }

    [Theory]
    [InlineData("Commodore/C64/250407/Service manual.pdf", "Service manual.pdf")]
    [InlineData("Commodore/Shared files/Component images/6526.PNG", "6526.PNG")]
    [InlineData("Generic shared files/Component local files/notes.txt", "notes.txt")]
    public void A_file_of_a_type_a_submission_may_carry_keeps_its_own_name(string path, string expected)
    {
        Assert.True(ReviewTableFiles.TryGetOpenName(path, out string name));
        Assert.Equal(expected, name);
    }

    // ###########################################################################################
    // *** A CONTRIBUTOR'S WEB PAGE OPENS AS TEXT. *** Opened as .html it would run whatever script
    // it carries in the maintainer's browser; as .txt it is read, which is what a review needs.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore/Shared files/Board local files/CBM Approved Cross Reference.html", "CBM Approved Cross Reference.html.txt")]
    [InlineData("x/y/page.HTM", "page.HTM.txt")]
    public void A_web_page_is_saved_as_text(string path, string expected)
    {
        Assert.True(ReviewTableFiles.TryGetOpenName(path, out string name));
        Assert.Equal(expected, name);
    }

    // Nothing a submission may not carry is ever handed to the operating system.
    [Theory]
    [InlineData("Commodore/C64/250407/run.exe")]
    [InlineData("Commodore/C64/250407/drawing.svg")]
    [InlineData("Commodore/C64/250407/shortcut.lnk")]
    [InlineData("Commodore/C64/250407/Makefile")]
    [InlineData("")]
    [InlineData(null)]
    public void Any_other_type_is_not_opened(string? path)
    {
        Assert.False(ReviewTableFiles.TryGetOpenName(path, out _));
    }

    // Only the file's own name is kept, so the saved file cannot land anywhere but its own folder.
    [Theory]
    [InlineData("../../Windows/evil.pdf", "evil.pdf")]
    [InlineData("a\\b\\manual.pdf", "manual.pdf")]
    public void Folders_in_the_path_are_dropped(string path, string expected)
    {
        Assert.True(ReviewTableFiles.TryGetOpenName(path, out string name));
        Assert.Equal(expected, name);
    }
}
