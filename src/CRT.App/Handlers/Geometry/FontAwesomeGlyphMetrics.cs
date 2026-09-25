using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Why Font Awesome icons lose their top pixel row, and how much room to reserve so they do not.
    //
    // THE PROBLEM. A font declares an ascent - how far above the baseline its glyphs are supposed
    // to reach - and a text layout engine sizes the line box from that number. Font Awesome's Free
    // faces declare an ascent of 448 units on a 512-unit em, but a number of individual glyphs are
    // drawn TALLER than that: the padlocks (lock, lock-open) reach 480 units, 32 units past the
    // declared ascent. The engine has already decided the line is 448 units tall, so those 32 units
    // fall outside the box and are clipped - which on screen is the top one or two pixel rows of the
    // icon simply missing. The shackle of a padlock loses its curve; a bug loses the top of its head.
    //
    // It is not a rendering bug and it is not fixable by nudging alignment. The glyph genuinely
    // overflows its own line box, so the only fix is to give the box more room.
    //
    // WHICH ICONS ARE AFFECTED. Only those whose outline exceeds the declared ascent, which is a
    // property of each individual glyph rather than of the font. In the faces this app ships:
    //
    //     lock (U+F023)            yMax 480   overshoot 32   AFFECTED
    //     lock-open (U+F3C1)       yMax 480   overshoot 32   AFFECTED
    //     trash-can (U+F2ED)       yMax 464   overshoot 16   AFFECTED
    //     bug (U+F188)             yMax 448   overshoot  0   fine
    //     file-lines (U+F15C)      yMax 448   overshoot  0   fine
    //     circle-question (U+F059) yMax 448   overshoot  0   fine
    //     pen-to-square (U+F044)   yMax 444   overshoot  0   fine
    //     square-plus (U+F0FE)     yMax 416   overshoot  0   fine
    //     square-minus (U+F146)    yMax 416   overshoot  0   fine
    //     sort-down (U+F160)       yMax 416   overshoot  0   fine
    //     sort-up (U+F161)         yMax 416   overshoot  0   fine
    //
    // So "it worked last time" proves nothing about the next icon. Measure the glyph.
    //
    // trash-can is the one that proves the point. It sits 16 units over - half the padlocks'
    // overshoot, so the clipping is a single pixel row at UI sizes and reads as the lid simply
    // being drawn flat. It went unnoticed through the original fix because the test that should
    // have caught it treated "never measured" and "does not overshoot" as the same answer; see
    // FontAwesomeAssetTests.GetGlyphYMax.
    //
    // HOW TO MEASURE ONE. Read the shipped OTF directly - do not trust a value from memory, and do
    // not trust the headless test renderer, which returns zeroed glyph metrics and paints filled
    // boxes rather than real outlines, so it cannot show clipping at all:
    //
    //     python -c "
    //     from fontTools.ttLib import TTFont
    //     from fontTools.pens.boundsPen import BoundsPen
    //     f = TTFont('Assets/Fonts/Font Awesome 7 Free-Solid-900.otf')
    //     upem, asc = f['head'].unitsPerEm, f['hhea'].ascender
    //     gs, cmap = f.getGlyphSet(), f.getBestCmap()
    //     g = cmap[0xF3C1]; bp = BoundsPen(gs); gs[g].draw(bp)
    //     print(g, 'yMax', bp.bounds[3], 'ascent', asc, 'overshoot', bp.bounds[3] - asc)"
    //
    // THE FIX. Reserve the overshoot as top padding on the TextBlock, scaled to the font size:
    // GetTopOverflowPadding below turns a design-unit overshoot into pixels. One pixel is usually
    // enough at UI sizes, but the calculation is here so it stays right at any size.
    // ###########################################################################################
    public static class FontAwesomeGlyphMetrics
    {
        // The em square and declared ascent of the Font Awesome 7 Free faces (both Solid and
        // Regular carry the same values). Read out of the shipped OTFs.
        public const double FontAwesomeDesignEmHeight = 512.0;

        public const double FontAwesomeDeclaredAscent = 448.0;

        // How far past the declared ascent each glyph's outline actually reaches, in design units.
        // Only glyphs that overshoot need an entry; anything absent needs no padding.
        //
        // Keyed by codepoint so a caller can ask about the glyph it is about to render rather than
        // having to know which ones are troublesome - the whole point being that "which ones" is
        // not something anyone should be carrying in their head.
        private static readonly Dictionary<int, double> OvershootByCodepoint = new()
        {
            [0xF023] = 32.0, // lock
            [0xF3C1] = 32.0, // lock-open
            [0xF2ED] = 16.0, // trash-can
        };

        // ###########################################################################################
        // The top padding, in pixels, that a TextBlock rendering this glyph needs so its outline is
        // not clipped. Zero for a glyph that stays inside the declared ascent.
        //
        // Rounded UP to a whole pixel: a fractional reservation still leaves part of a device pixel
        // outside the box, which is exactly the row that goes missing.
        // ###########################################################################################
        public static double GetTopOverflowPadding(int codepoint, double fontSize)
        {
            if (fontSize <= 0 || !OvershootByCodepoint.TryGetValue(codepoint, out double overshoot))
            {
                return 0.0;
            }

            return Math.Ceiling(overshoot / FontAwesomeDesignEmHeight * fontSize);
        }

        // ###########################################################################################
        // The same reservation as a Thickness, ready to assign to TextBlock.Padding.
        // ###########################################################################################
        public static Thickness GetTopOverflowThickness(int codepoint, double fontSize) =>
            new(0, GetTopOverflowPadding(codepoint, fontSize), 0, 0);

        // ###########################################################################################
        // The reservation for the glyph a control is actually showing, taken from that control's own
        // text and font size.
        //
        // This is the form the UI should use. A literal Padding="0,1,0,0" in markup is right only
        // for the size it was written against - at FontSize 17 the padlocks need 2px - and markup
        // cannot call the calculation, so every hardcoded site is a clipped icon waiting for someone
        // to change a font size. Returns an empty Thickness for text that is not a listed glyph, so
        // it is safe to call on any icon.
        // ###########################################################################################
        public static Thickness GetTopOverflowThicknessForText(string? text, double fontSize)
        {
            if (string.IsNullOrEmpty(text))
            {
                return default;
            }

            return GetTopOverflowThickness(char.ConvertToUtf32(text, 0), fontSize);
        }

        // ###########################################################################################
        // Whether this glyph is drawn taller than its font claims, and so needs the reservation.
        // ###########################################################################################
        public static bool OverflowsDeclaredAscent(int codepoint) =>
            OvershootByCodepoint.ContainsKey(codepoint);
    }
}
