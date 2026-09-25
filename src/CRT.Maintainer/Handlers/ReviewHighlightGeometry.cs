using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // A MOVED HIGHLIGHT, turned into something drawable (NewContributeStrategy.md Phase 5, task 4).
    //
    // A component highlight is the rectangle CRT draws over a chip on a schematic. Its natural key
    // is SchematicName|BoardLabel, so X/Y/Width/Height are COMPARED FIELDS - which means a moved
    // highlight already arrives as an ordinary "changed" row with a field diff. What was missing
    // was the geometry to put it back on the board.
    //
    // *** THE PLACEMENT IS PROPORTIONAL, AND THAT IS THE ENTIRE DESIGN. *** A highlight is stored
    // in the schematic's own PIXEL coordinates, while the review panel draws that image at
    // whatever width the layout leaves it - a size decided during layout and never reported back.
    // So this converts pixels to FRACTIONS of the image and the drawing code multiplies by its own
    // size. No panel dimension appears here at all.
    //
    // The PDF export made the opposite mistake once and it is worth knowing about: it computed the
    // fractions correctly, then passed them to an API taking absolute lengths believing they were
    // percentages, so a highlight covering a tenth of the board was drawn covering most of it.
    // Nothing threw, and it was caught by holding the PDF next to the screen.
    //
    // *** AVALONIA-FREE ON PURPOSE. *** CRT's own HighlightRectBuilder does this with Avalonia's
    // Rect and lives in CRT.App, and CLAUDE.md says plainly not to drag those types across. This
    // is a small independent implementation rather than a shared one for that reason - and it
    // needs different behaviour anyway, since CRT resolves rects in pixels and this resolves them
    // in fractions.
    // ###########################################################################################
    public static class ReviewHighlightGeometry
    {
        // ###########################################################################################
        // The section the server reports highlights under.
        //
        // Read from CRT.Data rather than spelled out, because a literal on both sides of a process
        // boundary is a rename waiting to break exactly one screen: the maintainer app would find no
        // highlight section, draw no moves, and look perfectly correct while omitting the thing it
        // was built for.
        // ###########################################################################################
        public const string SectionName = global::Handlers.DataHandling.ReviewSummary.SectionComponentHighlights;

        // The field names as ReviewSummary reports them for a highlight row.
        private const string FieldX = "X";
        private const string FieldY = "Y";
        private const string FieldWidth = "Width";
        private const string FieldHeight = "Height";

        // ###########################################################################################
        // A highlight's rectangle as FRACTIONS of the image it sits on.
        //
        // Returns false rather than throwing for anything unusable - malformed numbers, a
        // zero-sized rect, an unknown image size, or a highlight that is entirely off the board.
        // Every one of those reaches here from contributed data in a manifest a stranger built.
        // ###########################################################################################
        public static bool TryBuild(
            string? x,
            string? y,
            string? width,
            string? height,
            double imageWidth,
            double imageHeight,
            out ReviewHighlightBox box)
        {
            box = default;

            // An unknown image size divides to infinity or NaN, and a NaN in a layout silently
            // collapses the control - the panel would show the picture with no highlight and look
            // entirely fine.
            if (imageWidth <= 0 || imageHeight <= 0)
                return false;

            if (!ReviewHighlightGeometry.TryParse(x, out double left) ||
                !ReviewHighlightGeometry.TryParse(y, out double top) ||
                !ReviewHighlightGeometry.TryParse(width, out double boxWidth) ||
                !ReviewHighlightGeometry.TryParse(height, out double boxHeight))
            {
                return false;
            }

            if (boxWidth <= 0 || boxHeight <= 0)
                return false;

            // *** CLIPPED TO THE IMAGE, NOT CLAMPED INWARD. *** Clamping only the origin keeps the
            // full width and slides the rectangle off the thing it marks - an entry at x=-50 w=100
            // on a 1000px board would be drawn 100 wide starting at 0, marking the wrong half. The
            // same rule ExportOverlayGeometry settled on, for the same reason.
            double right = Math.Min(left + boxWidth, imageWidth);
            double bottom = Math.Min(top + boxHeight, imageHeight);

            left = Math.Max(left, 0);
            top = Math.Max(top, 0);

            // Nothing of it is on the board. A zero-width box would draw as a hairline at the
            // edge, which reads as a real highlight in the wrong place.
            if (right <= left || bottom <= top)
                return false;

            box = new ReviewHighlightBox(
                left / imageWidth,
                top / imageHeight,
                (right - left) / imageWidth,
                (bottom - top) / imageHeight);

            return true;
        }

        // ###########################################################################################
        // The before and after coordinates of one highlight row.
        //
        // *** A FIELD DIFF ONLY LISTS WHAT CHANGED, and that is the subtlety here. *** A highlight
        // moved horizontally carries an X row and no Y row at all. Reading a missing field as
        // blank would place the "before" rectangle at the top of the board and invent a vertical
        // move that never happened - and inventing a change is worse than missing one, because the
        // maintainer acts on it. So an absent field reports the SAME value on both sides.
        //
        // Returns false when nothing geometric changed: this section's key is SchematicName and
        // BoardLabel, so a row can change without moving, and drawing an unmoved rectangle twice
        // would present a "before and after" where nothing moved.
        // ###########################################################################################
        public static bool TryReadMove(
            ReviewSectionView? section,
            string rowKey,
            out ReviewHighlightMove move)
        {
            move = default;

            if (section is null || string.IsNullOrWhiteSpace(rowKey))
                return false;

            if (!section.FieldChanges.TryGetValue(rowKey, out IReadOnlyList<ReviewFieldChangeView>? fields))
                return false;

            bool moved = false;

            // The unchanged value is not carried anywhere, so an absent field has to resolve to
            // something identical on both sides. Empty is the honest answer - it means "this
            // geometry did not change", and TryBuild refuses it, so a partially-known rectangle is
            // never drawn as though it were complete.
            (string before, string after) Read(string name)
            {
                ReviewFieldChangeView? field = fields.FirstOrDefault(candidate =>
                    string.Equals(candidate.Field, name, StringComparison.Ordinal));

                if (field is null)
                    return (string.Empty, string.Empty);

                moved = true;
                return (field.Before, field.After);
            }

            (string beforeX, string afterX) = Read(ReviewHighlightGeometry.FieldX);
            (string beforeY, string afterY) = Read(ReviewHighlightGeometry.FieldY);
            (string beforeWidth, string afterWidth) = Read(ReviewHighlightGeometry.FieldWidth);
            (string beforeHeight, string afterHeight) = Read(ReviewHighlightGeometry.FieldHeight);

            if (!moved)
                return false;

            move = new ReviewHighlightMove(
                beforeX, beforeY, beforeWidth, beforeHeight,
                afterX, afterY, afterWidth, afterHeight);

            return true;
        }

        // ###########################################################################################
        // Splits a highlight row key into the schematic and the component it names.
        //
        // The key is SchematicName|BoardLabel - BoardDraftNaturalKeys.ForComponentHighlight. A
        // board has many schematics, and drawing a move on the wrong one puts a rectangle over
        // unrelated circuitry.
        // ###########################################################################################
        public static bool TryReadKeyParts(string? rowKey, out string schematicName, out string boardLabel)
        {
            schematicName = string.Empty;
            boardLabel = string.Empty;

            if (string.IsNullOrWhiteSpace(rowKey))
                return false;

            int separator = rowKey.IndexOf(
                global::Handlers.DataHandling.BoardDraftNaturalKeys.Separator,
                StringComparison.Ordinal);

            if (separator <= 0 || separator >= rowKey.Length - 1)
                return false;

            schematicName = rowKey[..separator];
            boardLabel = rowKey[(separator + 1)..];

            return true;
        }

        // ###########################################################################################
        // INVARIANT CULTURE, always.
        //
        // These are STRINGS in BoardData, and this project has already been bitten three times by
        // culture-sensitive handling of them: "0.5" read under da-DK becomes 5, which is a
        // highlight ten times too far across - on a maintainer's machine only, and on nobody else's.
        // ###########################################################################################
        private static bool TryParse(string? text, out double value) =>
            double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    }

    // ###########################################################################################
    // A rectangle as FRACTIONS of the image it sits on - never pixels, never points.
    //
    // Every value is between 0 and 1, so the drawing code multiplies by whatever size it has
    // actually been given. See ReviewHighlightGeometry's header for why this is not a pixel rect.
    // ###########################################################################################
    public readonly record struct ReviewHighlightBox(double Left, double Top, double Width, double Height);

    // ###########################################################################################
    // Where a highlight was and where it is being put, as the raw stored strings.
    //
    // Kept as STRINGS rather than parsed here because the image size is not known at this point -
    // the picture has not been decoded yet. TryBuild turns a side into a box once it is.
    // ###########################################################################################
    public readonly record struct ReviewHighlightMove(
        string BeforeX,
        string BeforeY,
        string BeforeWidth,
        string BeforeHeight,
        string AfterX,
        string AfterY,
        string AfterWidth,
        string AfterHeight);
}
