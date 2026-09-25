using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers PublishedTreeView.FindCaseVariant - the check that stops a submitted path landing
    // beside a published one that differs only by capitalisation (security review, 2026-09-25).
    //
    // Two such paths are two files on the Linux server and ONE on every Windows and macOS client,
    // so whichever syncs last silently replaces the other there.
    // ###########################################################################################
    public sealed class PublishedTreeViewTests
    {
        // A tree built from relative file paths; folders are implied by them.
        internal static PublishedTreeView TreeOf(params string[] files)
        {
            return new PublishedTreeView(
                path => files.Contains(path, StringComparer.Ordinal) ? new string('a', 64) : null,
                folder =>
                {
                    string prefix = folder.Length == 0 ? string.Empty : folder + "/";

                    List<string> entries = files
                        .Where(path => path.StartsWith(prefix, StringComparison.Ordinal))
                        .Select(path => path[prefix.Length..].Split('/')[0])
                        .Distinct(StringComparer.Ordinal)
                        .ToList();

                    return entries.Count == 0 && folder.Length > 0 ? null : entries;
                });
        }

        [Fact]
        public void A_case_variant_FILE_is_found_and_named_as_it_exists()
        {
            PublishedTreeView tree = PublishedTreeViewTests.TreeOf("Commodore/C64/250407/Sheet1.png");

            Assert.Equal(
                "Commodore/C64/250407/Sheet1.png",
                tree.FindCaseVariant("Commodore/C64/250407/sheet1.png"));
        }

        // A variant FOLDER anywhere along the path is the same problem - the files beneath would be
        // merged on a client.
        [Fact]
        public void A_case_variant_FOLDER_part_way_down_is_found()
        {
            PublishedTreeView tree = PublishedTreeViewTests.TreeOf("Commodore/C64/250407/Scope baseline/x.png");

            Assert.Equal(
                "Commodore/C64/250407/Scope baseline",
                tree.FindCaseVariant("Commodore/C64/250407/scope baseline/new.png"));

            Assert.Equal("Commodore", tree.FindCaseVariant("commodore/C64/250407"));
        }

        [Fact]
        public void The_exact_path_is_not_a_variant_of_itself()
        {
            PublishedTreeView tree = PublishedTreeViewTests.TreeOf("Commodore/C64/250407/Sheet1.png");

            Assert.Null(tree.FindCaseVariant("Commodore/C64/250407/Sheet1.png"));
        }

        // A path that is new from some level down cannot collide below that level.
        [Fact]
        public void A_genuinely_new_path_has_no_variant()
        {
            PublishedTreeView tree = PublishedTreeViewTests.TreeOf("Commodore/C64/250407/Sheet1.png");

            Assert.Null(tree.FindCaseVariant("Commodore/C64/250407/Sheet2.png"));
            Assert.Null(tree.FindCaseVariant("Amstrad/CPC 664/MC0005A/board.png"));
            Assert.Null(PublishedTreeView.Empty.FindCaseVariant("Commodore/C64/250407/Sheet1.png"));
        }

        // Each folder is listed once however many paths walk through it - a 1,200-file submission
        // must not mean 1,200 directory reads of the same folder on the server.
        [Fact]
        public void Each_folder_is_listed_only_once()
        {
            var listed = new List<string>();

            var tree = new PublishedTreeView(
                _ => null,
                folder =>
                {
                    listed.Add(folder);
                    return folder switch
                    {
                        "" => ["Commodore"],
                        "Commodore" => ["C64"],
                        "Commodore/C64" => ["250407"],
                        _ => []
                    };
                });

            tree.FindCaseVariant("Commodore/C64/250407/a.png");
            tree.FindCaseVariant("Commodore/C64/250407/b.png");

            Assert.Equal(listed.Distinct().Count(), listed.Count);
        }

        [Fact]
        public void HashOf_answers_through_the_delegate_and_blank_is_nothing()
        {
            PublishedTreeView tree = PublishedTreeViewTests.TreeOf("a/b/c.png");

            Assert.Equal(new string('a', 64), tree.HashOf("a/b/c.png"));
            Assert.Null(tree.HashOf("a/b/d.png"));
            Assert.Null(tree.HashOf(""));
        }
    }
}
