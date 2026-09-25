using System.Globalization;
using CRT.Maintainer.Handlers;

namespace CRT.Maintainer.Tests;

// Covers ReviewHighlightGeometry - turning a moved highlight into something drawable (Phase 5,
// task 4, "a moved highlight drawn on the schematic, before and after").
//
// *** THE PLACEMENT IS PROPORTIONAL, AND THAT IS THE WHOLE DESIGN. *** A highlight is stored in
// the schematic's own PIXEL coordinates, while the review panel draws that image at whatever
// width the layout leaves it - a size decided at layout time and never reported back. So the
// geometry converts pixels to FRACTIONS of the image, and the panel multiplies by its own size.
//
// The PDF export already made the opposite mistake once (ExportOverlayGeometry's header): it
// computed fractions correctly and then passed them to a points-based API believing they were
// percentages, so a highlight covering a tenth of the board was drawn covering most of it.
// Nothing threw. These tests exist so that class of error is caught by the suite rather than by
// holding the screen next to the board.
public sealed class ReviewHighlightGeometryTests
{
    [Fact]
    public void A_highlight_becomes_a_FRACTION_of_the_image()
    {
        // A 100x50 rect at (200,100) on a 1000x500 board is a fifth across, a fifth down, and a
        // tenth of the board in each direction.
        Assert.True(ReviewHighlightGeometry.TryBuild(
            "200", "100", "100", "50", imageWidth: 1000, imageHeight: 500, out ReviewHighlightBox box));

        Assert.Equal(0.2, box.Left, 6);
        Assert.Equal(0.2, box.Top, 6);
        Assert.Equal(0.1, box.Width, 6);
        Assert.Equal(0.1, box.Height, 6);
    }

    [Fact]
    public void Every_fraction_is_between_ZERO_and_ONE()
    {
        // The contract the drawing code relies on. A value outside it multiplied by a panel width
        // puts the rectangle off the image, where it is invisible - so the maintainer is shown no
        // highlight at all rather than a wrong one, which reads as "nothing moved".
        Assert.True(ReviewHighlightGeometry.TryBuild(
            "0", "0", "1000", "500", 1000, 500, out ReviewHighlightBox box));

        Assert.InRange(box.Left, 0, 1);
        Assert.InRange(box.Top, 0, 1);
        Assert.InRange(box.Width, 0, 1);
        Assert.InRange(box.Height, 0, 1);
    }

    [Fact]
    public void Coordinates_parse_INVARIANT_so_a_Danish_machine_agrees_with_a_British_one()
    {
        // *** THE BUG THIS PROJECT HAS ALREADY HIT THREE TIMES. *** These are STRINGS in BoardData,
        // and "0.5" read under da-DK with culture-sensitive parsing becomes 5 - a highlight ten
        // times too far across, on a maintainer's machine only. Verified by actually switching the
        // culture rather than by reading the parse call.
        CultureInfo original = CultureInfo.CurrentCulture;

        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("da-DK");

            Assert.True(ReviewHighlightGeometry.TryBuild(
                "100.5", "50.25", "10.5", "20.75", 1000, 500, out ReviewHighlightBox box));

            Assert.Equal(0.1005, box.Left, 6);
            Assert.Equal(0.1005, box.Top, 6);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void A_highlight_running_off_the_EDGE_is_CLIPPED_not_slid_inward()
    {
        // *** CLIPPED, NOT CLAMPED - the same rule the PDF export settled on. *** Clamping only the
        // origin keeps the full width and slides the rectangle off the thing it marks: a highlight
        // at x=-50 w=100 on a 1000px board would be drawn 100 wide starting at 0, marking the
        // wrong half. Intersecting keeps the part that is genuinely on the board.
        Assert.True(ReviewHighlightGeometry.TryBuild(
            "-50", "0", "100", "50", 1000, 500, out ReviewHighlightBox box));

        Assert.Equal(0, box.Left, 6);

        // Half the rect was off the left edge, so half of it survives.
        Assert.Equal(0.05, box.Width, 6);
    }

    [Fact]
    public void A_highlight_running_off_the_RIGHT_edge_is_clipped_too()
    {
        Assert.True(ReviewHighlightGeometry.TryBuild(
            "950", "0", "100", "50", 1000, 500, out ReviewHighlightBox box));

        Assert.Equal(0.95, box.Left, 6);
        Assert.Equal(0.05, box.Width, 6);
    }

    [Fact]
    public void A_highlight_ENTIRELY_off_the_image_is_refused_rather_than_drawn_as_a_sliver()
    {
        // Nothing of it is on the board, so there is nothing to show. Returning a zero-width box
        // would draw a hairline at the edge, which a maintainer would read as a real highlight in
        // the wrong place.
        Assert.False(ReviewHighlightGeometry.TryBuild(
            "2000", "0", "100", "50", 1000, 500, out _));
    }

    [Theory]
    [InlineData("", "0", "10", "10")]
    [InlineData("abc", "0", "10", "10")]
    [InlineData("0", "0", "0", "10")]
    [InlineData("0", "0", "10", "0")]
    [InlineData("0", "0", "-10", "10")]
    public void A_MALFORMED_or_EMPTY_highlight_is_refused_rather_than_throwing(
        string x, string y, string width, string height)
    {
        // These values are contributed data and reach this from a manifest a stranger built. A
        // zero or negative size draws as nothing or inverts the rectangle; either way it is not
        // something to put in front of a maintainer as a comparison.
        Assert.False(ReviewHighlightGeometry.TryBuild(x, y, width, height, 1000, 500, out _));
    }

    [Theory]
    [InlineData(0, 500)]
    [InlineData(1000, 0)]
    [InlineData(-1, 500)]
    public void An_UNKNOWN_image_size_is_refused_rather_than_dividing_by_zero(
        double imageWidth, double imageHeight)
    {
        // The image size comes from decoding the picture, which can fail. Dividing by it without
        // checking yields infinity or NaN, and a NaN in a layout silently collapses the control -
        // so the panel would show the picture with no highlight and look perfectly fine.
        Assert.False(ReviewHighlightGeometry.TryBuild(
            "10", "10", "10", "10", imageWidth, imageHeight, out _));
    }

    // -----------------------------------------------------------------------------------------
    // Reading the BEFORE and AFTER coordinates out of a field diff.
    // -----------------------------------------------------------------------------------------

    // The real separator, taken from CRT.Data rather than typed out. It is U+241F (the SYMBOL FOR
    // UNIT SEPARATOR), chosen precisely because it cannot occur in a schematic name or a board
    // label - and it is invisible in most editors, so a hand-typed "|" here would build a key that
    // looks right in the source and matches nothing at runtime.
    private static string Key(string schematicName, string boardLabel) =>
        global::Handlers.DataHandling.BoardDraftNaturalKeys.ForComponentHighlight(schematicName, boardLabel);

    private static ReviewSectionView Section(params ReviewFieldChangeView[] fields) =>
        new(
            ReviewHighlightGeometry.SectionName,
            Added: [],
            Removed: [],
            Changed: [ReviewHighlightGeometryTests.Key("Sheet 1", "U8")],
            Renamed: [],
            FieldChanges: new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>
            {
                [ReviewHighlightGeometryTests.Key("Sheet 1", "U8")] = fields
            });

    [Fact]
    public void A_MOVED_highlight_yields_both_its_old_and_new_position()
    {
        // The whole point: a maintainer sees where it WAS and where it is being PUT, on the board
        // itself. A row diff saying "X: 100 -> 400" is technically the same information and tells
        // nobody whether the new position is right.
        ReviewSectionView section = ReviewHighlightGeometryTests.Section(
            new ReviewFieldChangeView("X", "100", "400"),
            new ReviewFieldChangeView("Y", "50", "250"));

        Assert.True(ReviewHighlightGeometry.TryReadMove(
            section, ReviewHighlightGeometryTests.Key("Sheet 1", "U8"), out ReviewHighlightMove move));

        Assert.Equal("100", move.BeforeX);
        Assert.Equal("400", move.AfterX);
        Assert.Equal("50", move.BeforeY);
        Assert.Equal("250", move.AfterY);
    }

    [Fact]
    public void An_UNCHANGED_field_carries_the_SAME_value_on_both_sides()
    {
        // *** THE SUBTLE ONE. *** A field diff only lists what CHANGED, so a highlight that moved
        // horizontally carries no Y row at all. Reading a missing field as blank would place the
        // "before" rectangle at the top of the board and invent a vertical move that never
        // happened - inventing a change is worse than missing one, because the maintainer acts on it.
        ReviewSectionView section = ReviewHighlightGeometryTests.Section(
            new ReviewFieldChangeView("X", "100", "400"));

        Assert.True(ReviewHighlightGeometry.TryReadMove(
            section, ReviewHighlightGeometryTests.Key("Sheet 1", "U8"), out ReviewHighlightMove move));

        Assert.Equal("100", move.BeforeX);
        Assert.Equal("400", move.AfterX);

        // Y never changed, so both sides must report whatever it is - not blank, and identical.
        Assert.Equal(move.BeforeY, move.AfterY);
    }

    [Fact]
    public void A_RESIZED_highlight_is_a_move_too()
    {
        // Growing a highlight to cover a neighbouring pin is as much a change to what it marks as
        // sliding it, and is easier to miss in a row diff.
        ReviewSectionView section = ReviewHighlightGeometryTests.Section(
            new ReviewFieldChangeView("Width", "50", "120"));

        Assert.True(ReviewHighlightGeometry.TryReadMove(
            section, ReviewHighlightGeometryTests.Key("Sheet 1", "U8"), out ReviewHighlightMove move));

        Assert.Equal("50", move.BeforeWidth);
        Assert.Equal("120", move.AfterWidth);
    }

    [Fact]
    public void A_row_whose_geometry_did_NOT_change_is_not_a_move()
    {
        // A highlight row can change without moving - this section's key is SchematicName and
        // BoardLabel, so nothing else in it is geometry. Drawing an unmoved rectangle twice would
        // put a "before and after" in front of a maintainer where nothing moved at all.
        ReviewSectionView section = ReviewHighlightGeometryTests.Section(
            new ReviewFieldChangeView("Some other column", "a", "b"));

        Assert.False(ReviewHighlightGeometry.TryReadMove(section, ReviewHighlightGeometryTests.Key("Sheet 1", "U8"), out _));
    }

    [Fact]
    public void A_key_the_section_does_not_carry_is_refused()
    {
        ReviewSectionView section = ReviewHighlightGeometryTests.Section(
            new ReviewFieldChangeView("X", "1", "2"));

        Assert.False(ReviewHighlightGeometry.TryReadMove(section, ReviewHighlightGeometryTests.Key("Sheet 1", "U99"), out _));
    }

    [Fact]
    public void The_SCHEMATIC_is_read_off_the_row_key_so_the_move_is_drawn_on_the_right_board()
    {
        // The key is SchematicName|BoardLabel (BoardDraftNaturalKeys.ForComponentHighlight). A
        // board has many schematics, and drawing the move on the wrong one would show a maintainer a
        // rectangle sitting over unrelated circuitry.
        Assert.True(ReviewHighlightGeometry.TryReadKeyParts(
            ReviewHighlightGeometryTests.Key("Sheet 1", "U8"), out string schematic, out string label));

        Assert.Equal("Sheet 1", schematic);
        Assert.Equal("U8", label);
    }

    [Fact]
    public void A_key_with_no_separator_is_refused_rather_than_guessed_at()
    {
        Assert.False(ReviewHighlightGeometry.TryReadKeyParts("U8", out _, out _));
    }

    [Fact]
    public void The_section_name_matches_what_the_SERVER_calls_it()
    {
        // A contract with ReviewSummary.SectionComponentHighlights across a process boundary. If
        // these drift, the maintainer app finds no highlight section and silently draws no moves -
        // the screen looks fine and simply omits the thing it was built for.
        Assert.Equal(
            global::Handlers.DataHandling.ReviewSummary.SectionComponentHighlights,
            ReviewHighlightGeometry.SectionName);
    }
}
