using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** A PASTED PICTURE'S ALPHA, MENDED (owner report, 2026-10-15: a screenshot pasted into the
    // Feedback tab's text showed as a white box with only its corner handles). *** Avalonia 12.1.4's
    // Windows clipboard reader (OleDataObjectHelper) turns a CF_DIB or CF_BITMAP - what Windows'
    // own screenshot tools and many programs put on the clipboard, with no PNG beside it - into a
    // 32-bit bitmap labelled OPAQUE whose alpha bytes are all 0 (the fourth byte of a BI_RGB pixel,
    // which nobody sets). Skia promises nothing for an opaque bitmap with alpha below 255: drawn on
    // the CPU it shows, but on the screen it came out fully transparent - measured on the
    // reporting machine's own clipboard: 662 x 897, every alpha byte 0. Its PNG for the mail was
    // fine all along.
    //
    // So the four-byte pixels of a pasted bitmap are looked at once: when they are labelled opaque
    // but some alpha is not 255, or carry no alpha at all (every byte 0 - a picture nobody can see
    // is never what was meant), every alpha byte is set to 255. FeedbackComposer makes the new
    // bitmap; this decides and does the bytes.
    // ###########################################################################################
    public static class PastedBitmapAlpha
    {
        public static bool NeedsOpaqueAlpha(ReadOnlySpan<byte> fourBytePixels, bool labelledOpaque)
        {
            if (fourBytePixels.Length < 4)
                return false;

            for (int alpha = 3; alpha < fourBytePixels.Length; alpha += 4)
            {
                byte value = fourBytePixels[alpha];

                // Opaque must mean 255 everywhere; otherwise any alpha at all means it is meant.
                if (labelledOpaque ? value != 255 : value != 0)
                    return labelledOpaque;
            }

            return !labelledOpaque;
        }

        public static void SetOpaque(Span<byte> fourBytePixels)
        {
            for (int alpha = 3; alpha < fourBytePixels.Length; alpha += 4)
                fourBytePixels[alpha] = 255;
        }
    }
}
