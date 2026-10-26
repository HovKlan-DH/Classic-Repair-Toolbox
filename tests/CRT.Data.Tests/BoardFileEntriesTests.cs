using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers BoardFileEntries - what each file of a board becomes in the maintainer's file trees
    // (owner request, 2026-09-28: "a file-structure for all existing files in BETA or PROD, and then
    // a highlighting of files changed (added, removed, changed)").
    // ###########################################################################################
    public sealed class BoardFileEntriesTests
    {
        private const string Board = "Commodore/C128/310378";

        private static string P(string rest) => $"{BoardFileEntriesTests.Board}/{rest}";

        private static BoardFileEntry Of(IReadOnlyList<BoardFileEntry> entries, string path) =>
            Assert.Single(entries, entry => entry.Path == path);

        private static SubmittedFileFact Fact(string path, string sha, string? published) =>
            new(path, sha, 10, SubmissionFileScope.Own, IsReferenced: true, PublishedSha256: published);

        // -----------------------------------------------------------------------------------
        // Production after publishing (Beta > Prod)
        // -----------------------------------------------------------------------------------

        // Everything after the promotion is BETA's bytes; a removed file exists only in production.
        [Fact]
        public void A_promotion_names_copies_removals_and_the_rest_and_where_each_opens_from()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForPromotion(
                [
                    new PromotionFile(P("new.png"), "n", PromotionChange.Added, PromotionStage.Content, IsShared: false),
                    new PromotionFile(P("Data.xlsx"), "w", PromotionChange.Replaced, PromotionStage.Board, IsShared: false)
                ],
                [P("old.png")],
                [P("same.png"), "Commodore/Shared files/manual.pdf"]);

            Assert.Equal(
                ["Commodore/C128/310378/Data.xlsx", "Commodore/C128/310378/new.png", "Commodore/C128/310378/old.png",
                 "Commodore/C128/310378/same.png", "Commodore/Shared files/manual.pdf"],
                entries.Select(entry => entry.Path));

            Assert.Equal(new BoardFileEntry(P("new.png"), BoardFileChange.Added, BoardFileSource.Beta), Of(entries, P("new.png")));
            Assert.Equal(BoardFileChange.Changed, Of(entries, P("Data.xlsx")).Change);
            Assert.Equal(new BoardFileEntry(P("old.png"), BoardFileChange.Removed, BoardFileSource.Production), Of(entries, P("old.png")));
            Assert.Equal(BoardFileChange.Unchanged, Of(entries, "Commodore/Shared files/manual.pdf").Change);
        }

        [Fact]
        public void An_older_server_with_no_unchanged_list_still_gives_the_changes()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForPromotion(
                [new PromotionFile(P("a.png"), "a", PromotionChange.Added, PromotionStage.Content, IsShared: false)], null, null);

            Assert.Equal([P("a.png")], entries.Select(entry => entry.Path));
        }

        // -----------------------------------------------------------------------------------
        // The BETA data after approving (Contributor Submissions)
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The reported case, as it will now look: only "Short one-liner" changed. The workbook is
        // written again (it carries the publish date), the highlight file is not (its content is
        // the same), every other file stays - and the whole folder is there around them.
        // ###########################################################################################
        [Fact]
        public void A_rows_only_submission_changes_the_workbook_and_nothing_else()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForApproval(
                [BoardFileEntriesTests.Fact(P("Images/a.png"), "same", "same")],
                [],
                [P("Images/a.png"), P("Data C128 310378.xlsx"), P("Data C128 310378 v2.0.0.xlsx"), P("Data C128 310378 v2.0.0.json")],
                new GeneratedFile(P("Data C128 310378 v2.0.0.xlsx"), ExistsInBeta: true, Changes: true),
                new GeneratedFile(P("Data C128 310378 v2.0.0.json"), ExistsInBeta: true, Changes: false));

            Assert.Equal(
                [P("Data C128 310378 v2.0.0.xlsx")],
                entries.Where(entry => entry.Change != BoardFileChange.Unchanged).Select(entry => entry.Path));

            BoardFileEntry workbook = Of(entries, P("Data C128 310378 v2.0.0.xlsx"));
            Assert.True(workbook.WrittenOnApproval);
            Assert.Equal(BoardFileSource.Beta, workbook.OpenFrom);

            // The older generation's workbook is there, and stays.
            Assert.Equal(BoardFileChange.Unchanged, Of(entries, P("Data C128 310378.xlsx")).Change);
        }

        // A file the submission carries opens as the contributor sent it - by its hash - whether it
        // is new or replaces one; one it carries unchanged opens from BETA.
        [Fact]
        public void Submitted_files_are_new_replaced_or_the_same_by_BETAs_bytes()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForApproval(
                [
                    BoardFileEntriesTests.Fact(P("new.png"), "n", null),
                    BoardFileEntriesTests.Fact(P("changed.png"), "c2", "c1"),
                    BoardFileEntriesTests.Fact("Commodore/Shared files/manual.pdf", "m", "m")
                ],
                [],
                [P("changed.png")],
                null,
                null);

            Assert.Equal(new BoardFileEntry(P("new.png"), BoardFileChange.Added, BoardFileSource.Submission, "n"), Of(entries, P("new.png")));
            Assert.Equal(new BoardFileEntry(P("changed.png"), BoardFileChange.Changed, BoardFileSource.Submission, "c2"), Of(entries, P("changed.png")));
            Assert.Equal(new BoardFileEntry("Commodore/Shared files/manual.pdf", BoardFileChange.Unchanged, BoardFileSource.Beta), Of(entries, "Commodore/Shared files/manual.pdf"));
        }

        [Fact]
        public void A_file_the_approval_removes_is_marked_removed()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForApproval([], [P("old.png")], [P("old.png"), P("kept.png")], null, null);

            Assert.Equal(BoardFileChange.Removed, Of(entries, P("old.png")).Change);
            Assert.Equal(BoardFileChange.Unchanged, Of(entries, P("kept.png")).Change);
        }

        // ###########################################################################################
        // A NEW BOARD "should just show how the new file-system will look like": everything added,
        // and the two files the approval generates marked as not written yet - there is nothing to
        // open until it is approved.
        // ###########################################################################################
        [Fact]
        public void A_new_board_is_all_new_and_its_generated_files_are_not_written_yet()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForApproval(
                [BoardFileEntriesTests.Fact(P("Images/a.png"), "a", null)],
                [],
                [],
                new GeneratedFile(P("Data.xlsx"), ExistsInBeta: false, Changes: true),
                new GeneratedFile(P("Data.json"), ExistsInBeta: false, Changes: true));

            Assert.All(entries, entry => Assert.Equal(BoardFileChange.Added, entry.Change));
            Assert.Equal(BoardFileSource.NotWrittenYet, Of(entries, P("Data.xlsx")).OpenFrom);
            Assert.Equal(BoardFileSource.NotWrittenYet, Of(entries, P("Data.json")).OpenFrom);
        }

        // A path is one entry however many sources name it, spelled with forward slashes.
        [Fact]
        public void Each_path_is_one_entry_with_forward_slashes()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForApproval(
                [BoardFileEntriesTests.Fact(P("a.png"), "a2", "a1")],
                [],
                // windows-path-literal: a backslash spelling compared as text, normalised to slashes on every OS.
                ["Commodore\\C128\\310378\\a.png", P("a.png")],
                null,
                null);

            Assert.Equal(P("a.png"), Assert.Single(entries).Path);
        }

        // -----------------------------------------------------------------------------------
        // A board as BETA holds it (the Boards screen's Files view, 2026-10-03)
        // -----------------------------------------------------------------------------------

        // "Just list all files": nothing changes, everything opens from BETA, a file both in the
        // folder and cited by a row is listed once, and the list is in path order.
        [Fact]
        public void A_boards_files_are_all_listed_once_unchanged_and_from_BETA()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForBoard(
                [P("manual.pdf"), P("Data.xlsx")],
                [P("manual.pdf"), "Commodore/Shared files/74LS08.pdf"]);

            Assert.Equal(
                new[] { P("Data.xlsx"), P("manual.pdf"), "Commodore/Shared files/74LS08.pdf" }.Order(StringComparer.Ordinal),
                entries.Select(entry => entry.Path));

            Assert.All(entries, entry =>
            {
                Assert.Equal(BoardFileChange.Unchanged, entry.Change);
                Assert.Equal(BoardFileSource.Beta, entry.OpenFrom);
                Assert.False(entry.WrittenOnApproval);
            });
        }

        [Fact]
        public void A_board_with_no_files_lists_none()
        {
            Assert.Empty(BoardFileEntries.ForBoard(null, null));
        }

        // Sizes in every file tree (owner request, 2026-10-04): an entry gets the size answered for
        // it, one with none answered keeps none, and a negative answer is no size at all.
        [Fact]
        public void WithSizes_gives_each_entry_the_size_answered_for_it_and_leaves_the_rest()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForBoard(["a/x.png", "a/y.png", "a/z.png"], null);

            IReadOnlyList<BoardFileEntry> sized = BoardFileEntries.WithSizes(entries, entry => entry.Path switch
            {
                "a/x.png" => 1234,
                "a/z.png" => -1,
                _ => null
            });

            Assert.Equal([1234L, null, null], sized.Select(entry => entry.SizeBytes));
            Assert.Equal(entries.Select(entry => entry.Path), sized.Select(entry => entry.Path));
            Assert.Empty(BoardFileEntries.WithSizes(null, _ => 1));
        }

        // ---- the CRT 4.0 tree (2026-10-25) --------------------------------------------------------

        // A file shown by its board name is fetched from where its bytes lie; one with no other
        // place lies at its name.
        [Fact]
        public void An_entry_is_fetched_from_its_tree_path_when_it_has_one()
        {
            Assert.Equal("Files/3f/3f7a.png", new BoardFileEntry("Images/u8.png", BoardFileChange.Unchanged, BoardFileSource.Beta, TreePath: "Files/3f/3f7a.png").PlaceInTree);
            Assert.Equal("Images/u8.png", new BoardFileEntry("Images/u8.png", BoardFileChange.Unchanged, BoardFileSource.Beta).PlaceInTree);
        }

        // The place is worked out by each end from Path and TreePath, so it is no field on the wire
        // - the API carries the two it is made from.
        [Fact]
        public void An_entrys_place_in_its_tree_is_not_sent_on_the_wire()
        {
            string json = System.Text.Json.JsonSerializer.Serialize(
                new BoardFileEntry("Images/u8.png", BoardFileChange.Unchanged, BoardFileSource.Beta, TreePath: "Files/3f/3f7a.png"),
                ReviewApiContract.WireSettings);

            Assert.Contains("\"treePath\":\"Files/3f/3f7a.png\"", json, StringComparison.Ordinal);
            Assert.DoesNotContain("placeInTree", json, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WithTreePaths_gives_each_entry_its_place_and_leaves_the_rest()
        {
            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForBoard(["board.json", "u8.png", "u9.png"], null);

            IReadOnlyList<BoardFileEntry> placed = BoardFileEntries.WithTreePaths(entries, entry => entry.Path switch
            {
                "u8.png" => "Files/aa/aa.png",
                "board.json" => "board.json",
                _ => null
            });

            Assert.Equal([null, "Files/aa/aa.png", null], placed.Select(entry => entry.TreePath));
        }

        // ###########################################################################################
        // *** A PROMOTION OF A 4.0 BOARD IS DRAWN BY ITS NAMES. *** The plan copies stored files,
        // whose paths are checksums; the tree shows each under the board's name, opened from its
        // place - and a name whose old bytes leave production while its new ones arrive is ONE file
        // changed, not one removed beside one added.
        // ###########################################################################################
        [Fact]
        public void A_promotion_of_a_4_0_board_is_shown_by_name_with_a_replaced_file_changed()
        {
            var names = new Dictionary<string, string>
            {
                ["Files/aa/new.png"] = "Images/u8.png",
                ["Files/bb/old.png"] = "Images/u8.png",
                ["Files/cc/same.pdf"] = "manual.pdf",
                ["Commodore/C64/250407/board.json"] = "board.json"
            };

            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForPromotion(
                [
                    new PromotionFile("Files/aa/new.png", "aa", PromotionChange.Added, PromotionStage.Content, false, "Images/u8.png"),
                    new PromotionFile("Commodore/C64/250407/board.json", "11", PromotionChange.Replaced, PromotionStage.Board, false, "board.json")
                ],
                removals: ["Files/bb/old.png"],
                unchanged: ["Files/cc/same.pdf"],
                names);

            Assert.Equal(["Images/u8.png", "board.json", "manual.pdf"], entries.Select(entry => entry.Path));

            BoardFileEntry u8 = entries[0];
            Assert.Equal(BoardFileChange.Changed, u8.Change);
            Assert.Equal("Files/aa/new.png", u8.PlaceInTree);
            Assert.Equal(BoardFileSource.Beta, u8.OpenFrom);

            Assert.Equal("Files/cc/same.pdf", entries[2].PlaceInTree);
        }

        // The board file holds everything; it is not "the highlight file" an Excel board had.
        [Theory]
        [InlineData("Commodore/C64/250407/Data C64 250407 v2.0.0.json", true)]
        [InlineData("Commodore/C64/250407/board.json", false)]
        [InlineData("board.json", false)]
        [InlineData("x.png", false)]
        public void Only_an_Excel_boards_sidecar_is_a_highlight_file(string path, bool expected)
        {
            Assert.Equal(expected, BoardFileEntries.IsHighlightFile(path));
        }
    }
}
