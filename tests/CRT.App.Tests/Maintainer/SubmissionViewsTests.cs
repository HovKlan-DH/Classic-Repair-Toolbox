using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers SubmissionViews - a submission's three views (owner request, 2026-09-30: "Maybe one row
// with "Contributor info", "Board data", and then "Files"?") and the count on the Files button.
//
// *** THE COUNT CARRIES WHAT THE RETIRED FILE LINES GUARANTEED. *** "Files: [3] included ([1]
// replaced under the same name)" and "KiCad data included: ..." stood above the table because a
// file replaced under its own path colours no cell - without them such a file was approved unseen.
// Those lines went (ReviewNotInTable), so every case their tests covered is covered here, against
// the button that now says it.
// ###########################################################################################
public sealed class SubmissionViewsTests
{
    private const string Board = "Commodore/C128/310378";

    private static SubmittedFileFact File(string name, string hash = "aa", string? published = null) =>
        new($"{Board}/{name}", hash, 10, SubmissionFileScope.Own, IsReferenced: true, PublishedSha256: published);

    private static SubmittedFileFact KiCad(string name, string hash = "aa", string? published = null) =>
        new($"{Board}/KiCad data/{name}", hash, 10, SubmissionFileScope.Own, IsReferenced: false, PublishedSha256: published);

    private static FileRemovalPreview Removing(params string[] names) =>
        new(names.Select(name => $"{Board}/{name}").ToList(), null);

    // The one the whole count exists for: same path, different bytes, and no cell in the table moves.
    [Fact]
    public void A_file_replaced_under_its_own_name_is_counted_although_no_row_changed()
    {
        Assert.Equal(1, SubmissionViews.ChangingFiles([File("manual.pdf", "new", published: "old")], null));
    }

    [Fact]
    public void A_new_file_is_counted_and_one_published_as_it_is_is_not()
    {
        Assert.Equal(1, SubmissionViews.ChangingFiles(
            [
                File("extra.png", "bb", published: null),
                File("Sheet1.png", "aa", published: "aa")
            ],
            null));
    }

    // KiCad data is cited by no row, so no sheet ever shows it - counted like any other file.
    [Fact]
    public void KiCad_data_is_counted_when_new_or_changed_and_not_when_it_travels_unchanged()
    {
        Assert.Equal(2, SubmissionViews.ChangingFiles(
            [
                KiCad("board.kicad_pcb", "aa", published: "aa"),
                KiCad("board.kicad_sch", "bb", published: "old"),
                KiCad("Pages/vic.kicad_sch", "cc", published: null)
            ],
            null));
    }

    // A file the approval removes is a change too - and a submission carrying it unchanged keeps it,
    // exactly as the tree decides (the submitted file wins over the removal).
    [Fact]
    public void A_removed_file_is_counted_unless_the_submission_carries_it_as_it_is()
    {
        Assert.Equal(1, SubmissionViews.ChangingFiles([], Removing("old.png")));

        Assert.Equal(0, SubmissionViews.ChangingFiles(
            [File("old.png", "aa", published: "aa")],
            Removing("old.png")));
    }

    // A new system: nothing is published, so every file it carries is new.
    [Fact]
    public void A_new_systems_files_are_all_counted_as_new()
    {
        Assert.Equal(3, SubmissionViews.ChangingFiles(
            [File("Sheet1.png"), File("manual.pdf", "bb"), KiCad("board.kicad_pcb", "cc")],
            null));
    }

    // Nothing changes: no badge - a "0" would read as something to look at.
    [Fact]
    public void A_submission_changing_no_file_has_no_badge()
    {
        int changing = SubmissionViews.ChangingFiles([File("Sheet1.png", "aa", published: "aa")], FileRemovalPreview.Nothing);

        Assert.Equal(0, changing);
        Assert.Null(SubmissionViews.FilesBadge(changing));
        Assert.Equal("4", SubmissionViews.FilesBadge(4));
        Assert.Equal(0, SubmissionViews.ChangingFiles(null, null));
    }

    // ###########################################################################################
    // *** THE BUTTON AND THE TREE SAY THE SAME NUMBER. *** The button counts from the submission
    // detail; the tree is the server's (SubmissionFileTreeFlow) and carries the workbook and the
    // highlight file too, which the approval writes from the table. Built here exactly as the server
    // builds it - the same facts, plus the board's other files and both generated ones - the tree's
    // headline must open with the button's number. Fails if either side counts the generated files
    // and the other does not.
    // ###########################################################################################
    [Fact]
    public void The_files_count_is_the_number_the_trees_headline_gives()
    {
        SubmittedFileFact[] submitted =
        [
            File("manual.pdf", "new", published: "old"),
            File("Sheet1.png", "aa", published: "aa"),
            File("extra.png", "bb", published: null),
            KiCad("board.kicad_sch", "cc", published: "old")
        ];

        FileRemovalPreview removals = Removing("gone.png");

        IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForApproval(
            submitted,
            removals.Files,
            ownInBeta: [$"{Board}/Sheet1.png", $"{Board}/manual.pdf", $"{Board}/gone.png", $"{Board}/Untouched.png"],
            workbook: new GeneratedFile($"{Board}/Board.xlsx", ExistsInBeta: true, Changes: true),
            sidecar: new GeneratedFile($"{Board}/Board.json", ExistsInBeta: true, Changes: true));

        int changing = SubmissionViews.ChangingFiles(submitted, removals);

        Assert.Equal(4, changing);
        Assert.StartsWith($"{changing} files change: ", FileTreeWording.Summary(FileTree.Build(entries)), StringComparison.Ordinal);
    }

    // ###########################################################################################
    // Board data first and opened on (owner request, 2026-09-26: "make the table the default first
    // view"); then Files, then Contributor. A row of buttons whose chosen one sits in the middle
    // reads as if something had already been clicked.
    // ###########################################################################################
    [Fact]
    public void Board_data_comes_first_and_is_what_a_submission_opens_on()
    {
        Assert.Equal([SubmissionView.BoardData, SubmissionView.Files, SubmissionView.Contributor], SubmissionViews.Order);
        Assert.Equal(SubmissionView.BoardData, SubmissionViews.Opening);
        Assert.Equal(["Board data", "Files", "Contributor"], SubmissionViews.Order.Select(SubmissionViews.Label));
    }
}
