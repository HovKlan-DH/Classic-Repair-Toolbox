using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// The reservation that stops Font Awesome icons losing their top pixel row.
//
// The cause, which recurs because it is invisible until someone looks closely at a screenshot: a
// font declares how far above the baseline its glyphs reach, and the text engine sizes the line
// box from that number - but individual glyphs can be drawn TALLER than the declaration. Font
// Awesome's padlocks reach 480 units on a 512-unit em against a declared ascent of 448, so 32
// units fall outside the box and are clipped away.
//
// It is per-GLYPH, not per-font: bug, file-lines and the plus/minus squares all stay inside 448
// and render perfectly, which is why "the last icon was fine" is no evidence about the next one.
//
// These pin the arithmetic. FontAwesomeAssetTests pins the font data the arithmetic is based on,
// so a font upgrade that moves a glyph cannot leave this table quietly wrong.
public class FontAwesomeGlyphMetricsTests
{
    private const int LockOpen = 0xF3C1;
    private const int Lock = 0xF023;
    private const int Bug = 0xF188;
    private const int FileLines = 0xF15C;
    private const int SquareMinus = 0xF146;

    // ------------------------------------------------------------- which glyphs need it

    [Theory]
    [InlineData(Lock)]
    [InlineData(LockOpen)]
    public void The_padlocks_overflow_their_declared_ascent(int codepoint)
    {
        Assert.True(FontAwesomeGlyphMetrics.OverflowsDeclaredAscent(codepoint));
    }

    // The icons that render correctly today must NOT be padded - reserving a row they do not need
    // pushes them down and off-centre, trading one visual defect for another.
    [Theory]
    [InlineData(Bug)]
    [InlineData(FileLines)]
    [InlineData(SquareMinus)]
    public void An_icon_that_stays_inside_the_ascent_needs_no_reservation(int codepoint)
    {
        Assert.False(FontAwesomeGlyphMetrics.OverflowsDeclaredAscent(codepoint));
        Assert.Equal(0.0, FontAwesomeGlyphMetrics.GetTopOverflowPadding(codepoint, 11));
    }

    // An unknown codepoint is assumed fine. The table lists the exceptions, so a glyph nobody has
    // measured gets no padding rather than a guess.
    [Fact]
    public void An_unlisted_glyph_gets_no_reservation()
    {
        Assert.Equal(0.0, FontAwesomeGlyphMetrics.GetTopOverflowPadding(0xF007, 11));
    }

    // ------------------------------------------------------------- how much

    // 32/512 of the em, rounded UP to a whole pixel - a fractional reservation still leaves part of
    // a device pixel outside the box, which is precisely the row that goes missing.
    [Theory]
    [InlineData(10.0, 1.0)]   // 0.625px -> 1
    [InlineData(11.0, 1.0)]   // 0.688px -> 1
    [InlineData(16.0, 1.0)]   // 1.0px   -> 1
    [InlineData(17.0, 2.0)]   // 1.063px -> 2
    [InlineData(64.0, 4.0)]   // 4.0px   -> 4
    public void The_reservation_is_the_overshoot_scaled_to_the_font_size(double fontSize, double expected)
    {
        Assert.Equal(expected, FontAwesomeGlyphMetrics.GetTopOverflowPadding(LockOpen, fontSize));
    }

    // The reservation grows with the font size. A fixed 1px would be silently wrong on a large
    // icon, which is the reason this is a calculation rather than a constant.
    [Fact]
    public void A_larger_icon_reserves_more_room()
    {
        double small = FontAwesomeGlyphMetrics.GetTopOverflowPadding(Lock, 11);
        double large = FontAwesomeGlyphMetrics.GetTopOverflowPadding(Lock, 48);

        Assert.True(large > small, $"48px reserved {large}, no more than 11px reserved {small}");
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-4.0)]
    public void A_meaningless_font_size_reserves_nothing(double fontSize)
    {
        Assert.Equal(0.0, FontAwesomeGlyphMetrics.GetTopOverflowPadding(LockOpen, fontSize));
    }

    // ------------------------------------------------------------- the Thickness form

    // Top only. Padding on any other edge would shift the icon sideways or add space below it,
    // neither of which has anything to do with the clipping being fixed.
    [Fact]
    public void The_thickness_reserves_the_top_edge_only()
    {
        var thickness = FontAwesomeGlyphMetrics.GetTopOverflowThickness(LockOpen, 11);

        Assert.Equal(1.0, thickness.Top);
        Assert.Equal(0.0, thickness.Left);
        Assert.Equal(0.0, thickness.Right);
        Assert.Equal(0.0, thickness.Bottom);
    }

    [Fact]
    public void An_unaffected_glyph_gets_an_empty_thickness()
    {
        Assert.Equal(default, FontAwesomeGlyphMetrics.GetTopOverflowThickness(Bug, 11));
    }

    // ------------------------------------------------------------- the text-based form

    // The form the UI uses, so a control's reservation comes from the glyph it is actually showing
    // rather than from a literal in markup that is right only for the size it was written against.
    [Theory]
    [InlineData("", 11.0, 1.0)]
    [InlineData("", 11.0, 1.0)]
    [InlineData("", 17.0, 2.0)]
    public void The_reservation_can_be_taken_from_the_text_a_control_shows(string text, double fontSize, double expected)
    {
        Assert.Equal(expected, FontAwesomeGlyphMetrics.GetTopOverflowThicknessForText(text, fontSize).Top);
    }

    // Safe to call on any icon: a glyph that does not overshoot reserves nothing, so the UI can
    // apply it uniformly without first asking which icons are troublesome.
    [Theory]
    [InlineData("")]
    [InlineData("")]
    [InlineData("")]
    public void An_icon_that_does_not_overshoot_reserves_nothing_from_its_text(string text)
    {
        Assert.Equal(default, FontAwesomeGlyphMetrics.GetTopOverflowThicknessForText(text, 11));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Blank_text_reserves_nothing(string? text)
    {
        Assert.Equal(default, FontAwesomeGlyphMetrics.GetTopOverflowThicknessForText(text, 11));
    }

    // 17px is the size at which a hardcoded 1px silently becomes wrong - the regression the helper
    // exists to prevent, asserted so the two forms cannot drift.
    [Fact]
    public void The_text_form_agrees_with_the_codepoint_form()
    {
        foreach (double size in new[] { 10.0, 11.0, 17.0, 48.0 })
        {
            Assert.Equal(
                FontAwesomeGlyphMetrics.GetTopOverflowThickness(LockOpen, size),
                FontAwesomeGlyphMetrics.GetTopOverflowThicknessForText("", size));
        }
    }
}
