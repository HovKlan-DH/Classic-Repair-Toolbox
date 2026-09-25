using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // The Font Awesome padlocks every worklog Open/Closed surface draws - the worklog bar's status
    // pill, a workbook card's status pill, the Workbooks tab's top-line, a worklog entry's state
    // pill, and the "#N" badges on both the Schematics tab and the Workbooks board pane.
    //
    // They lived as six separate pairs of private constants across three classes - two of those
    // pairs in two partials of the SAME class, on the stated grounds that the files were
    // "independent partials of different classes", which they are not.
    //
    // Why one pair and not six: FontAwesomeGlyphMetrics.OvershootByCodepoint is keyed off these
    // exact codepoints. A site left behind on a different codepoint gets no overshoot padding and
    // silently clips the top pixel row of its padlock - the precise defect that class was written
    // to fix. Sitting the constants beside it makes the two impossible to change apart.
    //
    // Codepoints and glyph strings are both here because both are needed: the int feeds
    // GetTopOverflowThickness, the string feeds TextBlock.Text. Spelled as hex codepoints with the
    // strings derived from them, rather than as literal glyph characters, so this file stays plain
    // ASCII and the two can never disagree.
    // ###########################################################################################
    public static class WorklogGlyphs
    {
        // fa-solid lock-open - an OPEN workbook or an open entry.
        public const int OpenCodepoint = 0xF3C1;

        // fa-solid lock - a CLOSED workbook or a resolved entry.
        public const int ClosedCodepoint = 0xF023;

        public static readonly string OpenGlyph = char.ConvertFromUtf32(OpenCodepoint);

        public static readonly string ClosedGlyph = char.ConvertFromUtf32(ClosedCodepoint);

        public static int CodepointFor(bool isResolved) => isResolved ? ClosedCodepoint : OpenCodepoint;

        public static string GlyphFor(bool isResolved) => isResolved ? ClosedGlyph : OpenGlyph;
    }
}
