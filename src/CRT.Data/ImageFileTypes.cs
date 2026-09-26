using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH FILES THE APPLICATIONS CAN DRAW AS A PICTURE - the one list (2026-09-26).
    //
    // Every picture is drawn through Avalonia's Bitmap (Skia), so this is the set it decodes. It is
    // deliberately narrower than the file types a submission may carry (SubmissionFileRules) and
    // than what ExternalTargetLauncher opens: a PDF or a text file is shown by opening it, not
    // drawn.
    //
    // It lived as two copies - the component editor's (ContributionPackaging) and the maintainer
    // application's image comparison (ReviewImageComparison) - which the board table's file
    // preview would have made three. Both now read this one.
    // ###########################################################################################
    public static class ImageFileTypes
    {
        public static readonly IReadOnlyList<string> DisplayableExtensions = new[]
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"
        };

        private static readonly HashSet<string> DisplayableExtensionSet =
            new(ImageFileTypes.DisplayableExtensions, StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // True when the name or path carries an extension that can be drawn. Blank input, a name
        // with no extension and every other type are false - the caller is deciding whether to try
        // decoding contributed bytes, so it fails closed.
        // ###########################################################################################
        public static bool IsDisplayable(string? pathValue)
        {
            string trimmed = pathValue?.Trim() ?? string.Empty;

            if (trimmed.Length == 0)
            {
                return false;
            }

            try
            {
                return ImageFileTypes.DisplayableExtensionSet.Contains(Path.GetExtension(trimmed));
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
    }
}
