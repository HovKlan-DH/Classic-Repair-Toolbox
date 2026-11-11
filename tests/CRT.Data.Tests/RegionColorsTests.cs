using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// A BOARD'S REGIONS IN COLORS OF ITS OWN (owner request, 2026-11-11: "In "Board: Regions" I would like to be able to set the
// background color, and then you should calculate the complementary text color, that gives the highest visiblity of the
// text. However, I would also like to be able to set the text color"; asked, kept "In the board data") - the rule, and every
// side it travels: the board file, a submission and its publish, an amendment, the review and the board's history, "Review
// changes", the draft's own bookkeeping, and the checks both sides hold a submission to. A change that moves one side only
// fails here.
// ###########################################################################################
public sealed class RegionColorsTests
{
    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private static ComponentEntry Component(string label, string region = "") =>
        new() { BoardLabel = label, Category = "IC", Region = region };

    private static BoardData Regional() => BoardRegions.WithRegions(Board(Component("U1")), true);

    private static BoardData Colored(string background = "#1F3A93", string text = "") =>
        RegionColors.WithColors(Regional(), "PAL", background, text);

    // ---- the rule ----------------------------------------------------------------------------------

    [Theory]
    [InlineData("#1f3a93", "#1F3A93")]
    [InlineData(" #FFFFFF ", "#FFFFFF")]
    [InlineData("#FFF", null)]
    [InlineData("2E8B57", null)]
    [InlineData("#2E8B5G", null)]
    [InlineData("SeaGreen", null)]
    [InlineData("#802E8B57", null)]
    [InlineData("", null)]
    [InlineData(null, null)]
    public void Only_RRGGBB_is_a_color(string? text, string? expected)
    {
        Assert.Equal(expected, RegionColors.Normalised(text));
    }

    // "The complementary text color, that gives the highest visibility": black or white, by WCAG 2's contrast ratio.
    [Theory]
    [InlineData("#000000", "#FFFFFF")]
    [InlineData("#FFFFFF", "#000000")]
    [InlineData("#1F3A93", "#FFFFFF")]
    [InlineData("#2E8B57", "#000000")]
    [InlineData("#E0912F", "#000000")]
    [InlineData("#FFFF00", "#000000")]
    [InlineData("#CD5C5C", "#000000")]
    [InlineData("#3A7CA5", "#000000")]
    [InlineData("not a color", "#000000")]
    public void The_text_worked_out_is_black_or_white_whichever_reads_better(string background, string expected)
    {
        Assert.Equal(expected, RegionColors.ContrastText(background));
    }

    [Fact]
    public void The_contrast_ratio_is_WCAGs()
    {
        Assert.Equal(21.0, RegionColors.ContrastRatio("#000000", "#FFFFFF"), precision: 6);
        Assert.Equal(1.0, RegionColors.ContrastRatio("#777777", "#777777"), precision: 6);
        Assert.Equal(4.48, RegionColors.ContrastRatio("#777777", "#FFFFFF"), precision: 2);
        Assert.Equal(1.0, RegionColors.ContrastRatio("nope", "#FFFFFF"));
    }

    [Fact]
    public void A_region_with_a_background_alone_takes_the_text_that_reads_best_and_a_text_set_is_kept()
    {
        Assert.Equal(("#1F3A93", "#FFFFFF"), RegionColors.Of(Colored(), "pal"));
        Assert.Equal(("#1F3A93", "#FFFF00"), RegionColors.Of(Colored(text: "#ffff00"), "PAL"));

        // A region with none, one the board does not have, no board: CRT's own colors.
        Assert.Null(RegionColors.Of(Colored(), "NTSC"));
        Assert.Null(RegionColors.Of(Colored(), "Japan"));
        Assert.Null(RegionColors.Of(null, "PAL"));
    }

    [Fact]
    public void Setting_what_it_holds_already_is_the_same_board_and_no_background_takes_the_colors_away()
    {
        BoardData colored = Colored();

        Assert.Same(colored, RegionColors.WithColors(colored, "PAL", "#1f3a93", ""));

        BoardData plain = Regional();
        Assert.Same(plain, RegionColors.WithColors(plain, "PAL", "", "#FFFFFF"));

        BoardData reset = RegionColors.WithColors(colored, "PAL", null, null);
        Assert.Empty(reset.RegionColors);
        Assert.Null(RegionColors.Of(reset, "PAL"));

        // A region the board does not have takes none.
        Assert.Same(colored, RegionColors.WithColors(colored, "Japan", "#000000", ""));
    }

    // A board whose rows decide its regions gets them written down: colors are for a region the board SAYS.
    [Fact]
    public void Colors_for_a_region_its_rows_decide_write_the_regions_down()
    {
        BoardData fromRows = Board(Component("U1", "PAL"));

        Assert.Empty(fromRows.Regions);

        BoardData colored = RegionColors.WithColors(fromRows, "PAL", "#000080", "");

        Assert.Equal(["PAL", "NTSC"], colored.Regions);
        Assert.Empty(RegionColors.Problems(colored.RegionColors, colored.Regions));
    }

    [Fact]
    public void A_region_deleted_or_every_region_taken_away_takes_its_colors_and_a_reorder_keeps_them_in_order()
    {
        BoardData three = BoardRegions.WithRegion(Regional(), "Japan");
        three = RegionColors.WithColors(three, "Japan", "#FF0000", "");
        three = RegionColors.WithColors(three, "NTSC", "#0000FF", "");

        Assert.Equal(["NTSC", "Japan"], three.RegionColors.Select(row => row.Region));

        BoardData reordered = BoardRegions.InOrder(three, ["Japan", "PAL", "NTSC"]);
        Assert.Equal(["Japan", "NTSC"], reordered.RegionColors.Select(row => row.Region));

        BoardData deleted = BoardRegions.WithoutRegion(three, "Japan");
        Assert.Equal(["NTSC"], deleted.RegionColors.Select(row => row.Region));

        Assert.Empty(BoardRegions.WithRegions(deleted, false).RegionColors);
    }

    // ---- the board file ---------------------------------------------------------------------------

    // A board that gives none is written byte for byte as before: nothing published needs converting.
    [Fact]
    public void The_board_file_says_colors_only_for_a_board_that_gives_some_and_reads_them_back()
    {
        string without = DataTreeJson.Write(BoardDocumentMapping.ToDocument(new BoardContent(Regional(), [], []), name => name, []));
        Assert.DoesNotContain("\"regionColors\"", without);

        string written = DataTreeJson.Write(BoardDocumentMapping.ToDocument(new BoardContent(Colored(text: "#FFFF00"), [], []), name => name, []));
        Assert.Contains("\"regionColors\"", written);

        BoardContent read = BoardDocumentMapping.ToContent(DataTreeJson.ReadBoard(written));
        Assert.Equal(("#1F3A93", "#FFFF00"), RegionColors.Of(read.Board, "PAL"));
        Assert.Equal(written, DataTreeJson.Write(BoardDocumentMapping.ToDocument(read, name => name, [])));
    }

    // ---- a submission, its publish and an amendment ------------------------------------------------

    [Fact]
    public void A_submission_carries_them_to_the_publish_and_not_sending_them_keeps_what_is_published()
    {
        SubmissionRows rows = SubmissionRowsBoard.FromBoard(Colored());

        Assert.Equal("#1F3A93", Assert.Single(rows.RegionColors!).Background);
        Assert.Equal(("#1F3A93", "#FFFFFF"), RegionColors.Of(PublishMerge.Build(new SubmissionManifest { Rows = rows }, Regional()), "PAL"));

        BoardData published = Colored();
        Assert.Equal(("#1F3A93", "#FFFFFF"), RegionColors.Of(PublishMerge.Build(new SubmissionManifest { Rows = new SubmissionRows { RegionColors = null } }, published), "PAL"));
        Assert.Null(RegionColors.Of(PublishMerge.Build(new SubmissionManifest { Rows = new SubmissionRows { RegionColors = [] } }, published), "PAL"));
    }

    // *** Security (code review, 2026-11-11): colors sent without regions keep the published regions - a color of a region
    // the board does not have was published beside them, never drawn. It is let go of at the publish. ***
    [Fact]
    public void A_publish_lets_go_of_colors_of_regions_the_merged_board_does_not_have()
    {
        var rows = new SubmissionRows
        {
            RegionColors =
            [
                new RegionColorEntry { Region = "PAL", Background = "#1F3A93" },
                new RegionColorEntry { Region = "Japan", Background = "#000000" },
            ],
        };

        BoardData merged = PublishMerge.Build(new SubmissionManifest { Rows = rows }, Regional());

        Assert.Equal("PAL", Assert.Single(merged.RegionColors).Region);
    }

    // A board file a hand edited to hold a color with no region is still rewritten - never a crash at every write.
    [Fact]
    public void A_color_row_with_no_region_in_a_board_file_does_not_stop_its_rewrite()
    {
        var previous = new BoardDocument { RegionColors = [new BoardDocument.RegionColorRow { Region = null!, Background = "#000000" }] };
        var next = new BoardDocument { RegionColors = [new BoardDocument.RegionColorRow { Region = null!, Background = "#000000" }] };

        BoardDocumentUnknowns.CarryOver(previous, next);
    }

    [Fact]
    public void An_amendment_keeps_the_submissions_whatever_a_request_says()
    {
        var current = new SubmissionRows { Regions = ["PAL", "NTSC"], RegionColors = [new RegionColorEntry { Region = "PAL", Background = "#1F3A93" }] };
        var edited = new SubmissionRows { Regions = [], RegionColors = [] };

        Assert.Single(SubmissionRowsBoard.WithTableSections(current, edited).RegionColors!);
    }

    // ---- the checks, the server's too -------------------------------------------------------------

    // *** Security: what is published is drawn on every CRT - "#RRGGBB" of a region the board says, each once. ***
    [Fact]
    public void The_checks_refuse_a_color_that_is_not_RRGGBB_or_not_of_a_region_the_board_says()
    {
        static List<BoardDataProblem> Found(params RegionColorEntry[] colors) =>
            BoardDataChecks.Check(BoardCheckRows.From(new SubmissionRows { Regions = ["PAL", "NTSC"], RegionColors = [.. colors] }), files: null, BoardCheckScope.Submission)
                .Where(problem => problem.Code == "regions.color")
                .ToList();

        Assert.Empty(Found(new RegionColorEntry { Region = "PAL", Background = "#1F3A93" }, new RegionColorEntry { Region = "NTSC", Background = "#000000", Text = "#FFFF00" }));

        Assert.Single(Found(new RegionColorEntry { Region = "PAL", Background = "red" }));
        Assert.Single(Found(new RegionColorEntry { Region = "PAL", Background = "#1F3A93", Text = "url(x)" }));
        Assert.Single(Found(new RegionColorEntry { Region = "Japan", Background = "#1F3A93" }));
        Assert.Single(Found(new RegionColorEntry { Region = "PAL", Background = "#1F3A93" }, new RegionColorEntry { Region = "pal", Background = "#000000" }));
        Assert.All(Found(new RegionColorEntry { Region = "PAL", Background = "red" }), problem => Assert.Equal(BoardProblemLevel.Error, problem.Level));

        // Not sent, nothing to check.
        Assert.DoesNotContain(
            BoardDataChecks.Check(BoardCheckRows.From(new SubmissionRows { Regions = ["PAL", "NTSC"] }), files: null, BoardCheckScope.Submission),
            problem => problem.Code == "regions.color");
    }

    // *** Security (code review, 2026-11-11): with no regions sent, the color rows name their regions themselves, and those
    // names reached the published board file whatever they said. They are held to every region name's rule and number. ***
    [Fact]
    public void Colors_sent_without_regions_name_regions_by_the_rule_every_region_name_follows()
    {
        static List<BoardDataProblem> Found(params RegionColorEntry[] colors) =>
            BoardDataChecks.Check(BoardCheckRows.From(new SubmissionRows { RegionColors = [.. colors] }), files: null, BoardCheckScope.Submission)
                .Where(problem => problem.Code == "regions.color")
                .ToList();

        Assert.Empty(Found(new RegionColorEntry { Region = "PAL", Background = "#1F3A93" }, new RegionColorEntry { Region = "Japan", Background = "#000000" }));

        Assert.Single(Found(new RegionColorEntry { Region = new string('x', BoardRegions.LongestName + 1), Background = "#1F3A93" }));
        Assert.Single(Found(new RegionColorEntry { Region = "<script>", Background = "#1F3A93" }));
        Assert.Single(Found(new RegionColorEntry { Region = BoardRegions.EveryRegion, Background = "#1F3A93" }));
        Assert.Single(Found([.. Enumerable.Range(1, BoardRegions.MostRegions + 1).Select(n => new RegionColorEntry { Region = $"R{n}", Background = "#1F3A93" })]));

        // A name twice is said once, as with regions sent.
        Assert.Single(Found(new RegionColorEntry { Region = "PAL", Background = "#1F3A93" }, new RegionColorEntry { Region = "pal", Background = "#000000" }));
    }

    // ---- the review, the history, the draft --------------------------------------------------------

    [Fact]
    public void The_review_and_the_boards_history_are_told_and_nothing_else_needs_to_have_changed()
    {
        ReviewChangeSummary changed = ReviewSummary.Compare(Regional(), Colored());

        Assert.True(changed.RegionColorsChanged);
        Assert.Null(changed.RegionsNow);
        Assert.True(changed.HasChanges);
        Assert.Equal("region colors changed", changed.Describe());

        SubmissionChanges facts = SubmissionChangeFacts.Build(changed);
        Assert.True(facts.RegionColorsChanged);
        Assert.True(SubmissionChangeFacts.HasChanges(facts));

        Assert.False(ReviewSummary.Compare(Colored(), Colored()).RegionColorsChanged);

        // A text set that is the one worked out anyway is no change.
        Assert.False(ReviewSummary.Compare(Colored(), Colored(text: "#FFFFFF")).RegionColorsChanged);
    }

    [Fact]
    public void Review_changes_says_it_and_bringing_a_draft_up_to_date_keeps_the_contributors()
    {
        Assert.Contains(
            DraftChangeSummary.Build(new BoardContent(Regional(), [], []), new BoardContent(Colored(), [], [])).AsLines(),
            line => line.EndsWith(RegionColors.ChangedLine, StringComparison.Ordinal));

        BoardData plain = Regional();
        BoardData colored = Colored();

        Assert.Equal(("#1F3A93", "#FFFFFF"), RegionColors.Of(DraftRebase.Merge(plain, colored, plain, null, null, null).Board, "PAL"));
        Assert.Equal(("#1F3A93", "#FFFFFF"), RegionColors.Of(DraftRebase.Merge(plain, plain, colored, null, null, null).Board, "PAL"));
    }

    // ###########################################################################################
    // A board whose rows decide its regions (no list written - most published boards): the contributor's colors wrote the
    // list down in the draft, which BoardRegions.Differ does not count as a change of the regions. Brought up to date, the
    // update's unwritten list came back beside the draft's colors - colors of no region the board SAYS, refused by the
    // checks ("regions.color") and so by Submit (code review, 2026-11-11).
    // ###########################################################################################
    [Fact]
    public void Bringing_a_draft_up_to_date_keeps_the_regions_its_colors_are_for_written_down()
    {
        BoardData fromRows = Board(Component("U1", "PAL"));
        BoardData colored = RegionColors.WithColors(fromRows, "PAL", "#000080", "");
        BoardData updated = Board(Component("U1", "PAL"), Component("U2", "NTSC"));

        BoardData merged = DraftRebase.Merge(fromRows, colored, updated, null, null, null).Board;

        Assert.Equal(["PAL", "NTSC"], merged.Regions);
        Assert.Equal(("#000080", "#FFFFFF"), RegionColors.Of(merged, "PAL"));
        Assert.DoesNotContain(
            BoardDataChecks.Check(BoardCheckRows.From(merged), files: null, BoardCheckScope.Submission),
            problem => problem.Code == "regions.color");
    }

    // Colors the board holds already, set again, are no step - never one that only writes the region list down.
    [Fact]
    public void Setting_the_colors_a_board_with_no_list_holds_already_is_the_same_board()
    {
        BoardData fromRows = Board(Component("U1", "PAL"));
        BoardData unwritten = RegionColors.WithColors(fromRows, "PAL", "#000080", "").WithRegions([]);

        Assert.Same(unwritten, RegionColors.WithColors(unwritten, "PAL", "#000080", ""));
    }

    // ---- what a region is applied on (owner request, 2026-11-11) -------------------------------------

    // "it should be possible to see which component has the region applied, so I can find these and change them" - every
    // component with a row saying it, each once in the board's order, and the board's own files and links. A component told
    // apart in every region is among them (owner report, 2026-11-11: on a PAL and NTSC board that is nearly every one, and
    // "Applied on:" showed nowhere), though its rows go WITH the region rather than hold it.
    [Fact]
    public void A_region_says_which_components_and_board_rows_it_is_applied_on()
    {
        var board = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components =
            [
                Component("U9"),
                Component("U7", "PAL"),
                Component("U7", "NTSC"),
                Component("C2", "PAL"),
                Component("U1"),
                Component("U3"),
            ],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U3", Name = "Top", File = "a.png", Region = "PAL" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Datasheet", Url = "https://example.org", Region = "PAL" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Region = "PAL" }],
            BoardLinks = [new BoardLinkEntry { Category = "Web", Name = "Forum", Url = "https://example.org", Region = "PAL" }],
        };

        BoardRegions.RegionHolding pal = BoardRegions.HoldingOf(board).Single(region => region.Region == "PAL");

        Assert.Equal(["U7", "C2", "U1", "U3"], pal.Components);
        Assert.Equal(0, pal.BoardFiles);
        Assert.Equal(1, pal.BoardLinks);

        // The rows counted, which keep the region: C2's own row, U3's picture, U1's link and the board link. U7's row and U1's
        // pin are of components in every region, told apart per region - they go WITH it (RowsHolding).
        Assert.Equal(4, pal.Rows);

        Assert.Equal(["U7"], BoardRegions.HoldingOf(board).Single(region => region.Region == "NTSC").Components);
    }
}
