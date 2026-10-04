using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SystemFileEntries - what each file of a system becomes in the maintainer's file trees
    // (owner request, 2026-09-28: "a file-structure for all existing files in BETA or PROD, and then
    // a highlighting of files changed (added, removed, changed)").
    // ###########################################################################################
    public sealed class SystemFileEntriesTests
    {
        private const string Board = "Commodore/C128/310378";

        private static string P(string rest) => $"{SystemFileEntriesTests.Board}/{rest}";

        private static SystemFileEntry Of(IReadOnlyList<SystemFileEntry> entries, string path) =>
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
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForPromotion(
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

            Assert.Equal(new SystemFileEntry(P("new.png"), SystemFileChange.Added, SystemFileSource.Beta), Of(entries, P("new.png")));
            Assert.Equal(SystemFileChange.Changed, Of(entries, P("Data.xlsx")).Change);
            Assert.Equal(new SystemFileEntry(P("old.png"), SystemFileChange.Removed, SystemFileSource.Production), Of(entries, P("old.png")));
            Assert.Equal(SystemFileChange.Unchanged, Of(entries, "Commodore/Shared files/manual.pdf").Change);
        }

        [Fact]
        public void An_older_server_with_no_unchanged_list_still_gives_the_changes()
        {
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForPromotion(
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
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForApproval(
                [SystemFileEntriesTests.Fact(P("Images/a.png"), "same", "same")],
                [],
                [P("Images/a.png"), P("Data C128 310378.xlsx"), P("Data C128 310378 v2.0.0.xlsx"), P("Data C128 310378 v2.0.0.json")],
                new GeneratedFile(P("Data C128 310378 v2.0.0.xlsx"), ExistsInBeta: true, Changes: true),
                new GeneratedFile(P("Data C128 310378 v2.0.0.json"), ExistsInBeta: true, Changes: false));

            Assert.Equal(
                [P("Data C128 310378 v2.0.0.xlsx")],
                entries.Where(entry => entry.Change != SystemFileChange.Unchanged).Select(entry => entry.Path));

            SystemFileEntry workbook = Of(entries, P("Data C128 310378 v2.0.0.xlsx"));
            Assert.True(workbook.WrittenOnApproval);
            Assert.Equal(SystemFileSource.Beta, workbook.OpenFrom);

            // The older generation's workbook is there, and stays.
            Assert.Equal(SystemFileChange.Unchanged, Of(entries, P("Data C128 310378.xlsx")).Change);
        }

        // A file the submission carries opens as the contributor sent it - by its hash - whether it
        // is new or replaces one; one it carries unchanged opens from BETA.
        [Fact]
        public void Submitted_files_are_new_replaced_or_the_same_by_BETAs_bytes()
        {
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForApproval(
                [
                    SystemFileEntriesTests.Fact(P("new.png"), "n", null),
                    SystemFileEntriesTests.Fact(P("changed.png"), "c2", "c1"),
                    SystemFileEntriesTests.Fact("Commodore/Shared files/manual.pdf", "m", "m")
                ],
                [],
                [P("changed.png")],
                null,
                null);

            Assert.Equal(new SystemFileEntry(P("new.png"), SystemFileChange.Added, SystemFileSource.Submission, "n"), Of(entries, P("new.png")));
            Assert.Equal(new SystemFileEntry(P("changed.png"), SystemFileChange.Changed, SystemFileSource.Submission, "c2"), Of(entries, P("changed.png")));
            Assert.Equal(new SystemFileEntry("Commodore/Shared files/manual.pdf", SystemFileChange.Unchanged, SystemFileSource.Beta), Of(entries, "Commodore/Shared files/manual.pdf"));
        }

        [Fact]
        public void A_file_the_approval_removes_is_marked_removed()
        {
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForApproval([], [P("old.png")], [P("old.png"), P("kept.png")], null, null);

            Assert.Equal(SystemFileChange.Removed, Of(entries, P("old.png")).Change);
            Assert.Equal(SystemFileChange.Unchanged, Of(entries, P("kept.png")).Change);
        }

        // ###########################################################################################
        // A NEW SYSTEM "should just show how the new file-system will look like": everything added,
        // and the two files the approval generates marked as not written yet - there is nothing to
        // open until it is approved.
        // ###########################################################################################
        [Fact]
        public void A_new_system_is_all_new_and_its_generated_files_are_not_written_yet()
        {
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForApproval(
                [SystemFileEntriesTests.Fact(P("Images/a.png"), "a", null)],
                [],
                [],
                new GeneratedFile(P("Data.xlsx"), ExistsInBeta: false, Changes: true),
                new GeneratedFile(P("Data.json"), ExistsInBeta: false, Changes: true));

            Assert.All(entries, entry => Assert.Equal(SystemFileChange.Added, entry.Change));
            Assert.Equal(SystemFileSource.NotWrittenYet, Of(entries, P("Data.xlsx")).OpenFrom);
            Assert.Equal(SystemFileSource.NotWrittenYet, Of(entries, P("Data.json")).OpenFrom);
        }

        // A path is one entry however many sources name it, spelled with forward slashes.
        [Fact]
        public void Each_path_is_one_entry_with_forward_slashes()
        {
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForApproval(
                [SystemFileEntriesTests.Fact(P("a.png"), "a2", "a1")],
                [],
                // windows-path-literal: a backslash spelling compared as text, normalised to slashes on every OS.
                ["Commodore\\C128\\310378\\a.png", P("a.png")],
                null,
                null);

            Assert.Equal(P("a.png"), Assert.Single(entries).Path);
        }

        // -----------------------------------------------------------------------------------
        // A system as BETA holds it (the Systems screen's Files view, 2026-10-03)
        // -----------------------------------------------------------------------------------

        // "Just list all files": nothing changes, everything opens from BETA, a file both in the
        // folder and cited by a row is listed once, and the list is in path order.
        [Fact]
        public void A_systems_files_are_all_listed_once_unchanged_and_from_BETA()
        {
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForSystem(
                [P("manual.pdf"), P("Data.xlsx")],
                [P("manual.pdf"), "Commodore/Shared files/74LS08.pdf"]);

            Assert.Equal(
                new[] { P("Data.xlsx"), P("manual.pdf"), "Commodore/Shared files/74LS08.pdf" }.Order(StringComparer.Ordinal),
                entries.Select(entry => entry.Path));

            Assert.All(entries, entry =>
            {
                Assert.Equal(SystemFileChange.Unchanged, entry.Change);
                Assert.Equal(SystemFileSource.Beta, entry.OpenFrom);
                Assert.False(entry.WrittenOnApproval);
            });
        }

        [Fact]
        public void A_system_with_no_files_lists_none()
        {
            Assert.Empty(SystemFileEntries.ForSystem(null, null));
        }

        // Sizes in every file tree (owner request, 2026-10-04): an entry gets the size answered for
        // it, one with none answered keeps none, and a negative answer is no size at all.
        [Fact]
        public void WithSizes_gives_each_entry_the_size_answered_for_it_and_leaves_the_rest()
        {
            IReadOnlyList<SystemFileEntry> entries = SystemFileEntries.ForSystem(["a/x.png", "a/y.png", "a/z.png"], null);

            IReadOnlyList<SystemFileEntry> sized = SystemFileEntries.WithSizes(entries, entry => entry.Path switch
            {
                "a/x.png" => 1234,
                "a/z.png" => -1,
                _ => null
            });

            Assert.Equal([1234L, null, null], sized.Select(entry => entry.SizeBytes));
            Assert.Equal(entries.Select(entry => entry.Path), sized.Select(entry => entry.Path));
            Assert.Empty(SystemFileEntries.WithSizes(null, _ => 1));
        }
    }
}
