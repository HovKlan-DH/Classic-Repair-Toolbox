using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Whether a submitted file's BYTES are what its NAME says (security review, 2026-09-25).
    //
    // *** THE EXTENSION IS CHOSEN BY THE CONTRIBUTOR; THE BYTES ARE WHAT GETS OPENED. *** Every
    // published file is synced to every user's disk, and CRT hands documents to the operating
    // system to open. A ".png" that is really an executable, or a ".txt" that is really a
    // script, only needs a user to open it. So the first bytes are checked against the format the
    // name claims, and a mismatch refuses the submission.
    //
    // WHAT THIS IS NOT: a decoder. It checks the signature every format opens with - enough to
    // refuse a file of the wrong KIND - but a well-formed signature on a deliberately malformed
    // image still passes. Decoding every image on the server would need an imaging library there,
    // which is a dependency decision for the project owner rather than something to add in passing.
    //
    // Pure: bytes in, verdict out. The caller reads the first SampleBytes of the file.
    // ###########################################################################################
    public static class SubmissionContentRules
    {
        // How much of the file the caller should read. Enough for every signature below and for a
        // meaningful look inside a text file, small enough to read for 1,200 files without noticing.
        public const int SampleBytes = 64 * 1024;

        // A PDF's "%PDF-" header may be preceded by junk; readers accept it within the first 1 KB,
        // so this does too - refusing a PDF every viewer opens would reject real board documents.
        private const int PdfHeaderWindow = 1024;

        private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        // ###########################################################################################
        // How many opening bytes Matches actually needs for this file's type - a signature for an
        // image, the header window for a PDF, the full sample for text (whose NUL test gets
        // stronger the more it sees). Lets a caller checking thousands of files read a few bytes
        // of each image rather than 64 KB of every one.
        // ###########################################################################################
        public static int BytesNeeded(string? path) =>
            Path.GetExtension(path ?? string.Empty).ToLowerInvariant() switch
            {
                ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".webp" => 32,
                ".pdf" => SubmissionContentRules.PdfHeaderWindow,
                _ => SubmissionContentRules.SampleBytes
            };

        // ###########################################################################################
        // Do these first bytes match the format the path's extension names?
        //
        // `head` is the first SampleBytes of the file (or all of it, when shorter). An extension
        // this class does not know is refused: SubmissionFileRules only admits extensions listed
        // here, so reaching the default arm means the two lists have drifted, which must fail
        // closed rather than wave the file through.
        // ###########################################################################################
        public static bool Matches(string? path, ReadOnlySpan<byte> head, out string reason)
        {
            reason = string.Empty;

            string extension = Path.GetExtension(path ?? string.Empty).ToLowerInvariant();

            // ###########################################################################################
            // *** AN IMAGE NAME ACCEPTS ANY REAL IMAGE, not only the format the extension names. ***
            //
            // Three images every Commodore board cites are misnamed in the published tree - a PNG
            // saved as "VIC-20 memory map basic screen color.jpg", and JPEGs saved as
            // "7700_251641-02_PLA.png" and "74LS125.png" (found by SubmissionRulesShippedDataTests).
            // Every viewer, the app and the review screen decode them by their content, so they
            // work; refusing them would have made every Commodore board unsubmittable. What
            // matters is that an image NAME carries image BYTES - a program renamed ".png" is
            // still refused.
            // ###########################################################################################
            bool matches = extension switch
            {
                ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".webp" => SubmissionContentRules.IsImage(head),
                ".pdf" => head[..Math.Min(head.Length, SubmissionContentRules.PdfHeaderWindow)].IndexOf("%PDF-"u8) >= 0,
                ".txt" or ".html" or ".htm" => SubmissionContentRules.LooksLikeText(head),

                // A board's KiCad data (2026-09-26). The board and schematic files are s-expression
                // text ("(kicad_pcb ..."), the project file is JSON - so each must be text AND open
                // with its format's own first character, which refuses a renamed executable and a
                // file of the wrong KiCad kind alike.
                ".kicad_pcb" or ".kicad_sch" =>
                    SubmissionContentRules.LooksLikeText(head) && SubmissionContentRules.OpensWith(head, (byte)'('),
                ".kicad_pro" =>
                    SubmissionContentRules.LooksLikeText(head) && SubmissionContentRules.OpensWith(head, (byte)'{'),

                _ => false
            };

            if (!matches)
            {
                reason = extension switch
                {
                    ".kicad_pcb" or ".kicad_sch" or ".kicad_pro" =>
                        "It does not look like a KiCad file of that kind. KiCad saves these as plain " +
                        "text; export the project again from KiCad rather than renaming a file.",
                    ".txt" or ".html" or ".htm" =>
                        "It does not look like a text file. Text files must be saved as plain text " +
                        "(UTF-8), not as a document or program.",
                    "" => "It has no file type.",
                    ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".webp" =>
                        "Its contents are not an image, although its name says it is one.",
                    _ => $"Its contents are not a {extension.TrimStart('.').ToUpperInvariant()} file, " +
                         "although its name says it is."
                };
            }

            return matches;
        }

        // The opening bytes of any of the five image formats board data carries.
        private static bool IsImage(ReadOnlySpan<byte> head) =>
            head.StartsWith(SubmissionContentRules.PngSignature) ||
            (head.Length >= 3 && head[0] == 0xFF && head[1] == 0xD8 && head[2] == 0xFF) ||
            head.StartsWith("GIF87a"u8) ||
            head.StartsWith("GIF89a"u8) ||
            (head.Length >= 26 && head.StartsWith("BM"u8)) ||
            (head.Length >= 12 && head.StartsWith("RIFF"u8) && head.Slice(8, 4).SequenceEqual("WEBP"u8));

        // ###########################################################################################
        // Does the text open with this character - after any whitespace and an optional UTF-8 BOM,
        // both of which a KiCad file saved through another editor can legitimately start with?
        // ###########################################################################################
        private static bool OpensWith(ReadOnlySpan<byte> head, byte opener)
        {
            ReadOnlySpan<byte> bom = [0xEF, 0xBB, 0xBF];

            if (head.StartsWith(bom))
                head = head[bom.Length..];

            foreach (byte value in head)
            {
                if (value is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n')
                    continue;

                return value == opener;
            }

            return false;
        }

        // ###########################################################################################
        // Text is text: no NUL byte anywhere in the sample. Every executable, archive, image and
        // office document carries NUL bytes within its first few kilobytes, while no plain text
        // file does - so this one test refuses all of them without needing to know which is which.
        //
        // UTF-16 text also carries NULs and is refused too, deliberately: the rest of the tree is
        // UTF-8, and the message says how to save it.
        // ###########################################################################################
        private static bool LooksLikeText(ReadOnlySpan<byte> head) => head.IndexOf((byte)0) < 0;
    }
}
