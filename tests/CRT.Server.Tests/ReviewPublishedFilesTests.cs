using System.Reflection;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewEndpoints.PublishedFilePaths - the list of files the PUBLISHED board
    // references, which the maintainer app pairs a submission's own files against (Phase 5, task 4).
    //
    // *** THIS EXISTS BECAUSE THE FIRST VERSION DERIVED IT ON THE CLIENT AND WAS WRONG. *** The
    // review window built the published side from the change summary's row keys. A summary's keys
    // are NATURAL keys - for a component image, BoardLabel|Region|Pin|Name - not file paths. So
    // nothing ever matched: every submitted image would have been reported as an addition, and a
    // deleted image would never have appeared at all.
    //
    // Both failures are silent. The panel draws perfectly; it simply describes something that did
    // not happen. That is the worst shape a defect can take on a screen whose entire purpose is
    // telling a maintainer what a submission does, so the server states the list outright instead.
    //
    // Reached by reflection rather than made public: it is an implementation detail of one
    // endpoint, and widening it to public for a test would put it on the server's API surface for
    // no caller. Same approach and same reasoning as ExternalTargetLauncherTests in CRT.App.
    // ###########################################################################################
    public sealed class ReviewPublishedFilesTests
    {
        private static IReadOnlyList<string> Invoke(BoardData? published)
        {
            MethodInfo method = typeof(ReviewEndpoints).GetMethod(
                "PublishedFilePaths",
                BindingFlags.NonPublic | BindingFlags.Static)!;

            Assert.NotNull(method);

            return (IReadOnlyList<string>)method.Invoke(null, [published])!;
        }

        private static BoardData Board(params string[] files) => new()
        {
            ComponentImages =
            [
                .. files.Select((file, index) => new ComponentImageEntry
                {
                    BoardLabel = $"U{index}",
                    Name = $"Capture {index}",
                    File = file
                })
            ]
        };

        [Fact]
        public void Every_referenced_image_FILE_is_listed()
        {
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(
                ReviewPublishedFilesTests.Board("Images/U8.png", "Images/U9.png"));

            Assert.Equal(["Images/U8.png", "Images/U9.png"], files);
        }

        [Fact]
        public void The_list_is_FILE_PATHS_and_not_natural_keys()
        {
            // *** THE REGRESSION TEST FOR THE BUG THIS CLASS EXISTS FOR. *** A natural key for
            // this row would read "U0|<region>|<pin>|Capture 0". Asserting the value is a path
            // means a future refactor that reaches for the key builder here fails loudly rather
            // than producing a comparison that matches nothing.
            string file = Assert.Single(ReviewPublishedFilesTests.Invoke(
                ReviewPublishedFilesTests.Board("Images/U8.png")));

            Assert.Equal("Images/U8.png", file);
            Assert.DoesNotContain("|", file);
        }

        [Fact]
        public void One_file_cited_by_SEVERAL_rows_is_listed_ONCE()
        {
            // Legitimate and common: the same scope capture cited for two pins. The maintainer app
            // wants the set of files, and a duplicate would draw the same comparison twice.
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(
                ReviewPublishedFilesTests.Board("Images/shared.png", "Images/shared.png"));

            Assert.Single(files);
        }

        [Fact]
        public void A_row_with_NO_file_contributes_nothing()
        {
            // A component image row can carry an expected reading and no picture. A blank entry
            // here would build a fetch URL for "" and report a failure that is really an empty
            // cell.
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(
                ReviewPublishedFilesTests.Board("Images/real.png", "", "   "));

            Assert.Equal(["Images/real.png"], files);
        }

        [Fact]
        public void Files_differing_only_in_CASE_are_kept_apart()
        {
            // The server's filesystem is Linux and the data tree is case-sensitive from Phase 3
            // onward. Folding them together here would hide one of the two from the maintainer.
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(
                ReviewPublishedFilesTests.Board("Images/U8.png", "Images/u8.png"));

            Assert.Equal(2, files.Count);
        }

        [Fact]
        public void The_order_is_STABLE_so_two_maintainers_see_the_same_list()
        {
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(
                ReviewPublishedFilesTests.Board("z.png", "a.png", "m.png"));

            Assert.Equal(["a.png", "m.png", "z.png"], files);
        }

        [Fact]
        public void A_NEW_SYSTEM_has_no_published_files_at_all()
        {
            // Null is what PublishedBoardReader returns for a system nobody has published, and it
            // correctly makes every image in the submission an addition.
            Assert.Empty(ReviewPublishedFilesTests.Invoke(null));
        }

        // -------------------------------------------------------------------------------------
        // *** ALL FOUR FILE SOURCES, NOT JUST COMPONENT IMAGES (regression, 2026-09-23). ***
        //
        // This method used to read ComponentImages alone, while the SUBMISSION side
        // (SubmissionManifestBuilder.CollectReferencedFiles) collects four sources: schematic
        // images, component images, component local files and board local files. The maintainer app
        // compares the two lists, so every file from the three missing sources was present on the
        // submitted side, absent on the published side, and reported to the maintainer as ADDED.
        //
        // Reported by the project owner: a submission changing one component's short description
        // listed every schematic image as "Added / Not in the published board".
        //
        // Every test ABOVE builds a board out of ComponentImages only, which is exactly why none
        // of them caught this - the gap was in a source they never populated.
        // -------------------------------------------------------------------------------------

        [Fact]
        public void A_SCHEMATIC_image_is_listed_as_a_published_file()
        {
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(new BoardData
            {
                Schematics =
                [
                    new BoardSchematicEntry
                    {
                        SchematicName = "Sheet 1",
                        SchematicImageFile = "Commodore/C64/250407/Board Layout 250407 NTSC.png"
                    }
                ]
            });

            Assert.Equal(["Commodore/C64/250407/Board Layout 250407 NTSC.png"], files);
        }

        [Fact]
        public void A_COMPONENT_LOCAL_FILE_is_listed_as_a_published_file()
        {
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(new BoardData
            {
                ComponentLocalFiles =
                [
                    new ComponentLocalFileEntry
                    {
                        BoardLabel = "U8", Name = "Datasheet", File = "Datasheets/6581.pdf"
                    }
                ]
            });

            Assert.Equal(["Datasheets/6581.pdf"], files);
        }

        [Fact]
        public void A_BOARD_LOCAL_FILE_is_listed_as_a_published_file()
        {
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(new BoardData
            {
                BoardLocalFiles =
                [
                    new BoardLocalFileEntry
                    {
                        Category = "Docs", Name = "Service manual", File = "Docs/service.pdf"
                    }
                ]
            });

            Assert.Equal(["Docs/service.pdf"], files);
        }

        // ###########################################################################################
        // *** THE TEST THAT WOULD FAIL IF ONLY ONE SIDE MOVED. ***
        //
        // CLAUDE.md's rule: a shared change needs a test that fails when the two halves disagree.
        // The published list and the submitted list are compared against each other, so what
        // matters is not what either contains in isolation but that they agree - which is why this
        // asserts against CollectReferencedFiles rather than against a hand-written expectation.
        //
        // Add a fifth file source to one side only and this test fails.
        // ###########################################################################################
        [Fact]
        public void The_published_list_names_exactly_what_the_SUBMISSION_side_would_collect()
        {
            var board = new BoardData
            {
                Schematics =
                [
                    new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "sheet1.png" }
                ],
                ComponentImages =
                [
                    new ComponentImageEntry { BoardLabel = "U8", Name = "Pin 1", File = "Scope/U8_1.png" }
                ],
                ComponentLocalFiles =
                [
                    new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Docs/6581.pdf" }
                ],
                BoardLocalFiles =
                [
                    new BoardLocalFileEntry { Category = "Docs", Name = "Manual", File = "Docs/manual.pdf" }
                ]
            };

            Assert.Equal(
                SubmissionManifestBuilder.CollectReferencedFiles(board),
                ReviewPublishedFilesTests.Invoke(board));
        }

        // The concrete shape of the report: a board whose schematic image is already published
        // must NOT have that image described as something the submission adds.
        [Fact]
        public void A_schematic_image_already_published_is_not_reported_as_new()
        {
            IReadOnlyList<string> files = ReviewPublishedFilesTests.Invoke(new BoardData
            {
                Schematics =
                [
                    new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "sheet1.png" }
                ]
            });

            Assert.Contains("sheet1.png", files);
        }

        // -------------------------------------------------------------------------------------
        // SchematicImageFiles - which picture each schematic is drawn from, so the maintainer app can
        // put a moved highlight back on its own board.
        // -------------------------------------------------------------------------------------

        private static IReadOnlyDictionary<string, string> InvokeImages(
            BoardData submitted,
            BoardData? published)
        {
            MethodInfo method = typeof(ReviewEndpoints).GetMethod(
                "SchematicImageFiles",
                BindingFlags.NonPublic | BindingFlags.Static)!;

            Assert.NotNull(method);

            return (IReadOnlyDictionary<string, string>)method.Invoke(null, [submitted, published])!;
        }

        private static BoardData WithSchematics(params (string Name, string File)[] schematics) => new()
        {
            Schematics =
            [
                .. schematics.Select(entry => new BoardSchematicEntry
                {
                    SchematicName = entry.Name,
                    SchematicImageFile = entry.File
                })
            ]
        };

        [Fact]
        public void Each_schematic_is_mapped_to_its_image()
        {
            IReadOnlyDictionary<string, string> images = ReviewPublishedFilesTests.InvokeImages(
                ReviewPublishedFilesTests.WithSchematics(("Sheet 1", "Images/s1.png")),
                published: null);

            Assert.Equal("Images/s1.png", images["Sheet 1"]);
        }

        [Fact]
        public void The_SUBMITTED_board_WINS_when_both_name_a_schematic()
        {
            // *** THE ORDERING IS THE POINT. *** A submission can repoint a schematic at a new
            // image, and the maintainer must see the board as the submission PROPOSES it - drawing
            // the moved highlight on the old picture would be judging the change against the
            // wrong board.
            IReadOnlyDictionary<string, string> images = ReviewPublishedFilesTests.InvokeImages(
                ReviewPublishedFilesTests.WithSchematics(("Sheet 1", "Images/new.png")),
                ReviewPublishedFilesTests.WithSchematics(("Sheet 1", "Images/old.png")));

            Assert.Equal("Images/new.png", images["Sheet 1"]);
        }

        [Fact]
        public void A_PUBLISHED_schematic_the_submission_does_not_mention_is_still_mapped()
        {
            // The anti-vacuity half of the rule above. A submission that moves a highlight without
            // touching the schematic row carries no entry for it, and the highlight must still
            // draw - so the published board fills the gap rather than the submission replacing the
            // whole map.
            IReadOnlyDictionary<string, string> images = ReviewPublishedFilesTests.InvokeImages(
                ReviewPublishedFilesTests.WithSchematics(("Sheet 2", "Images/s2.png")),
                ReviewPublishedFilesTests.WithSchematics(("Sheet 1", "Images/s1.png")));

            Assert.Equal("Images/s1.png", images["Sheet 1"]);
            Assert.Equal("Images/s2.png", images["Sheet 2"]);
        }

        [Fact]
        public void A_schematic_with_NO_image_file_is_skipped()
        {
            // A blank would build a fetch URL for "" and report a failure that is really an empty
            // cell in the workbook.
            IReadOnlyDictionary<string, string> images = ReviewPublishedFilesTests.InvokeImages(
                ReviewPublishedFilesTests.WithSchematics(("Sheet 1", ""), ("Sheet 2", "Images/s2.png")),
                published: null);

            Assert.Single(images);
            Assert.True(images.ContainsKey("Sheet 2"));
        }

        [Fact]
        public void Schematic_names_differing_only_in_CASE_are_kept_apart()
        {
            // Matched against a row key built from the same strings, on a case-sensitive tree.
            IReadOnlyDictionary<string, string> images = ReviewPublishedFilesTests.InvokeImages(
                ReviewPublishedFilesTests.WithSchematics(("Sheet 1", "a.png"), ("sheet 1", "b.png")),
                published: null);

            Assert.Equal(2, images.Count);
        }
    }
}
