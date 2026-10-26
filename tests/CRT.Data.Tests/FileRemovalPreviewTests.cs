using System;
using System.Collections.Generic;
using System.IO;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // FileRemovalPreview - the list of files a publish will remove, shown to the maintainer before
    // they approve and checked again when they do (owner, 2026-09-25: "It must be, so it is
    // clear what will happen").
    // ###########################################################################################
    public sealed class FileRemovalPreviewTests
    {
        // The common case - a rows-only edit drops no file - must not read the tree at all. A root
        // that does not exist proves it: reading it would come back blocked.
        [Fact]
        public void With_no_candidates_nothing_is_removed_and_the_tree_is_not_read()
        {
            string nowhere = Path.Combine(Path.GetTempPath(), "crt-no-such-tree-" + Guid.NewGuid().ToString("N"));

            FileRemovalPreview preview = FileRemovalPreview.Compute(nowhere, [], null);

            Assert.Same(FileRemovalPreview.Nothing, preview);
            Assert.False(preview.IsBlocked);
        }

        [Fact]
        public void An_unreadable_tree_removes_nothing_and_says_why()
        {
            string nowhere = Path.Combine(Path.GetTempPath(), "crt-no-such-tree-" + Guid.NewGuid().ToString("N"));

            FileRemovalPreview preview = FileRemovalPreview.Compute(nowhere, ["a/old.png"], null);

            Assert.True(preview.IsBlocked);
            Assert.Empty(preview.Files);
            Assert.StartsWith("Nothing is removed", preview.BlockedBecause, StringComparison.Ordinal);
        }

        [Fact]
        public void The_list_shown_matches_in_any_order_but_not_with_a_file_more_or_less()
        {
            var preview = new FileRemovalPreview(["a/one.png", "a/two.png"], null);

            Assert.True(preview.Matches(["a/two.png", "a/one.png"]));
            Assert.False(preview.Matches(["a/one.png"]));
            Assert.False(preview.Matches(["a/one.png", "a/two.png", "a/three.png"]));

            // Exact spelling: the maintainer was shown one path, and that path is what goes.
            Assert.False(preview.Matches(["a/ONE.png", "a/two.png"]));
        }

        // A client that sends no list was shown none, which is right only when nothing goes.
        [Fact]
        public void Nothing_shown_matches_only_nothing_to_remove()
        {
            Assert.True(FileRemovalPreview.Nothing.Matches(null));
            Assert.True(FileRemovalPreview.Nothing.Matches([]));
            Assert.False(new FileRemovalPreview(["a/one.png"], null).Matches(null));
        }

        // ###########################################################################################
        // *** A 4.0 TREE'S REMOVAL IS A STORED FILE, SHOWN BY ITS BOARD NAME (2026-10-25). ***
        // "Files/3f/3f7a...png" says nothing to a maintainer; the name the board knew it by does.
        // Matching stays on the place - what is removed is what was shown.
        // ###########################################################################################
        [Fact]
        public void A_removal_is_shown_by_its_board_name_and_matched_by_its_place()
        {
            var preview = new FileRemovalPreview(["Files/3f/3f7a.png", "Files/aa/aa01.pdf"], null)
                .Named(new Dictionary<string, string> { ["Files/3f/3f7a.png"] = "Images/u8.png", ["Files/zz/elsewhere.png"] = "x.png" });

            Assert.Equal("Images/u8.png", preview.NameOf("Files/3f/3f7a.png"));
            Assert.Equal("Files/aa/aa01.pdf", preview.NameOf("Files/aa/aa01.pdf"));
            Assert.Equal(["Files/3f/3f7a.png"], preview.Names!.Keys);
            Assert.True(preview.Matches(["Files/aa/aa01.pdf", "Files/3f/3f7a.png"]));
        }

        [Fact]
        public void Naming_nothing_leaves_the_preview_as_it_is()
        {
            var preview = new FileRemovalPreview(["a/one.png"], null);

            Assert.Same(preview, preview.Named(null));
            Assert.Same(FileRemovalPreview.Nothing, FileRemovalPreview.Nothing.Named(new Dictionary<string, string> { ["a"] = "b" }));
            Assert.Null(preview.Named(new Dictionary<string, string>()).Names);
        }
    }
}
