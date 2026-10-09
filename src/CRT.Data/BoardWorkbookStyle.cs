using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW A BOARD WORKBOOK LOOKS WHEN THE APPLICATION CREATES ONE (owner request,
    // 2026-09-23).
    //
    // *** WHY THIS EXISTS. *** BoardWorkbookWriter used to build an empty package and write bare
    // rows into it, so a published or newly created board lost the preamble, the coloured headers
    // and the column widths that every hand-maintained board carries.
    // Nothing was DELETED - it was simply never written - but the first publish of the C64 250407
    // turned a 190 KB owner-authored workbook into an 81 KB stripped one, and the header
    // block linking to the project's own documentation disappeared with it.
    //
    // *** THE VALUES ARE READ OFF THE REFERENCE BOARD, NOT INVENTED. *** Everything here was taken
    // from Assets/Data/Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx - the board CLAUDE.md and
    // the Wiki both name as the reference implementation - so a generated workbook matches the
    // hand-made ones rather than being a second house style.
    //
    // *** THE UUID COLUMN IS DELIBERATELY NOT REPRODUCED. *** Every sheet in that reference still
    // carries a "UUID v4" column, and it is retired (owner instruction, 2026-09-23): nothing
    // generates one and nothing reads one. Copying the reference faithfully would resurrect a dead
    // column in every board published from now on, so the layout is copied and that column is not.
    // ###########################################################################################
    public static class BoardWorkbookStyle
    {
        // ###########################################################################################
        // The preamble occupies the rows ABOVE the header, and its shape is the SAME ON EVERY SHEET:
        //
        //   1  # Hardware: <name>          16pt
        //   2  # Board: <name>             16pt
        //   3  # Revision date: <date>     16pt, the DATE ITSELF bold
        //   4  (blank)
        //   5  <section title>             filled
        //   6  <column headers>            filled - and the panes are FROZEN under this row
        //
        // *** NO DOCUMENTATION LINES ANY MORE (owner request, 2026-10-02). *** Three rows used to
        // sit between the blank line and the title: "Documentation of columns in this worksheet is
        // availble here:", a link to the old Wiki's anchor for the sheet, and another blank line.
        // The columns are maintained from within CRT now, so the project owner removed those rows
        // from every published board and asked for CRT to stop writing them.
        //
        // *** THE REVISION DATE IS ON EVERY SHEET (owner request, 2026-10-02). *** It used to be on
        // Board schematics alone, as the references had it, so the other sheets were a row shorter.
        // It is the same date on all of them, and READING is unchanged: the board's date is still
        // the one on Board schematics (BoardDataReader.ScanRevisionDate searches that sheet's
        // top-left 10x10), so a hand edit that updates only one sheet must update that one.
        //
        // *** THE READER SCANS FOR THE HEADER ROW, so none of this is load-bearing for parsing. ***
        // ReadSheetRows walks from row 1 looking for the row carrying the required column names.
        // That is what makes the preamble safe to change - a published board in the old shape and
        // one in the new shape read exactly alike.
        // ###########################################################################################
        public const int PreambleHardwareRow = 1;
        public const int PreambleBoardRow = 2;
        public const int PreambleRevisionDateRow = 3;

        // The section title band, and the column headers directly under it.
        public const int TitleRow = 5;
        public const int HeaderRow = 6;

        // 16pt, matching the reference's three identity lines.
        public const float IdentityFontSize = 16f;

        // ###########################################################################################
        // The title on the row directly above the column headers ("Components", "Board credits").
        //
        // Mostly the sheet's own name, but not always - the reference calls the Credits sheet's
        // band "Board credits" and the Important signals sheet's "KiCad references". Taken from the
        // reference rather than derived, because a derived name would be wrong for exactly those
        // two and right for no reason on the others.
        // ###########################################################################################
        private static readonly Dictionary<string, string> SectionTitles =
            new(StringComparer.OrdinalIgnoreCase)
            {
                [BoardWorkbookSchema.SheetBoardSchematics] = "Board schematic images",
                [BoardWorkbookSchema.SheetComponents] = "Components",
                [BoardWorkbookSchema.SheetComponentImages] = "Component images",
                [BoardWorkbookSchema.SheetComponentLocalFiles] = "Component local files",
                [BoardWorkbookSchema.SheetComponentLinks] = "Component links",
                [BoardWorkbookSchema.SheetBoardLocalFiles] = "Board local files",
                [BoardWorkbookSchema.SheetBoardLinks] = "Board links",
                [BoardWorkbookSchema.SheetKiCadImportantSignals] = "KiCad references",
                [BoardWorkbookSchema.SheetCredits] = "Board credits",
            };

        // ###########################################################################################
        // The two band colours, as literal RGB - and the values were VERIFIED against the reference
        // rather than eyeballed, because the first attempt at them was wrong twice over.
        //
        // The reference stores these as THEME references with a TINT, which is where the trap is:
        //
        //   header band        fgColor theme="0" tint="-0.15"   -> white darkened 15% -> D9D9D9
        //   section title band fgColor theme="1"  (no tint)     -> BLACK, font theme="0" -> WHITE
        //
        // Two things make that easy to get backwards. In the STYLES part theme="0" is the light
        // background (white) and theme="1" the dark text colour - the opposite order to the theme
        // part's own clrScheme listing, where dk1 comes first. And a tint is not stored as a
        // colour at all, so reading the slot alone gives white for something that renders grey.
        //
        // *** AND IT WAS GOT BACKWARDS, for the title band (owner report, 2026-09-30). *** This
        // header said theme="1" is the dark colour and then called the band white, so every workbook
        // CRT wrote - drafts, and every board published to BETA - carried a white "Components"
        // strip where the shipped boards have a black one with white text. Checked again in the
        // XML of BOTH references (C64 250407, C128 310378): every sheet's title band is theme 1
        // (black) with theme 0 (white) text, across all its columns.
        //
        // Written as literal RGB rather than theme+tint so a generated workbook, which carries the
        // DEFAULT theme rather than the reference's, still renders the same colours.
        // ###########################################################################################

        // Black, with white text - the band carrying the sheet's title, directly above the header.
        public const string SectionTitleFillRgb = "FF000000";

        public const string SectionTitleFontRgb = "FFFFFFFF";

        // ###########################################################################################
        // THE BOARD SCHEMATICS SHEET IS BANDED BY MEANING, not painted in one colour
        // (owner report, 2026-09-24).
        //
        // Its title row groups the columns into three bands, and the header row picks the CAD-name
        // column out in blue:
        //
        //   title row   A-B  BLACK          the sheet's own title
        //               C-G  MID GREY       "Highlights in Main or Thumbnail" - one heading spanning five columns
        //               H    PALE BLUE      CAD name, which belongs to the KiCad import rather than to the board
        //   header row  A-G  LIGHT GREY     ordinary headers
        //               H    BLUE           CAD name again
        //
        // Every other sheet is a single light-grey header with no title banding at all, which is
        // why this lives here as a special case rather than in the generic path.
        //
        // The values are the reference's theme+tint references RESOLVED - Excel's tint maths, not a
        // guess: a negative tint darkens toward black and a positive one lightens toward white, so
        // "theme 1 at +0.4999" is black lightened halfway, which is #7F7F7F. Writing them literally
        // means a generated workbook renders the same under the default theme.
        // ###########################################################################################
        public const string TitleBandBlackRgb = "FF000000";

        public const string TitleBandGreyRgb = "FF7F7F7F";

        public const string TitleBandCadRgb = "FFDADDE1";

        public const string HeaderCadFillRgb = "FF8FAADC";

        // The heading that spans the five highlight columns on the schematics sheet.
        public const string HighlightsBandTitle = "Highlights in Main or Thumbnail";

        // ###########################################################################################
        // THE HEADER ROW'S HEIGHT, and why it is set explicitly (owner report, 2026-09-24).
        //
        // The reference wraps its column headers - "Schematic highlight opacity" over three lines -
        // and fixes the row at 57.6 points to fit them. Without the wrap those headers force their
        // columns wide enough to hold the whole phrase on one line, which is exactly what the
        // project owner's screenshot showed: a sheet three times wider than the reference with the
        // same data in it.
        //
        // Set rather than left to autofit because Excel does not re-measure a wrapped row's height
        // reliably when the file is written by a library rather than by Excel itself - the row can
        // come out one line tall with the text clipped.
        //
        // *** PER SHEET, AS THE REFERENCES HAVE IT (owner report, 2026-09-30). *** 57.6 is the
        // schematics sheet's height; it was used for EVERY sheet, which drew a tall grey band over
        // "Components" and every other header. The references (C64 250407, C128 310378) agree:
        // Board schematics 57.6, Components and Component images 28.8 (two lines - "Short one-liner
        // description" breaks onto a second), and every other sheet an ordinary one-line row with no
        // wrap. HeaderRowHeightFor answers null for those: the row's own default.
        // ###########################################################################################
        public const double HeaderRowHeight = 57.6;

        public const double TwoLineHeaderRowHeight = 28.8;

        public static double? HeaderRowHeightFor(string sheetName)
        {
            if (string.Equals(sheetName, BoardWorkbookSchema.SheetBoardSchematics, StringComparison.OrdinalIgnoreCase))
                return BoardWorkbookStyle.HeaderRowHeight;

            if (string.Equals(sheetName, BoardWorkbookSchema.SheetComponents, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(sheetName, BoardWorkbookSchema.SheetComponentImages, StringComparison.OrdinalIgnoreCase))
            {
                return BoardWorkbookStyle.TwoLineHeaderRowHeight;
            }

            return null;
        }

        // ###########################################################################################
        // THE TEXT A HEADER CELL CARRIES - the schema's column name, except where the references
        // break it over two lines (owner report, 2026-09-30): "Short one-liner description" and
        // "(one short line only!)" are separated by an Alt+Enter in both. Written on one line, the
        // header ran across a column far wider than its reference width - and that width, keyed on
        // the broken form, was never found at all.
        //
        // Safe to read: BoardDataReader.NormalizeHeader turns a line break in a header into a
        // space, which is the schema's name exactly.
        // ###########################################################################################
        public static string HeaderTextFor(string columnName) =>
            string.Equals(columnName, BoardWorkbookSchema.ColDescription, StringComparison.Ordinal)
                ? "Short one-liner description\n(one short line only!)"
                : columnName;

        // ###########################################################################################
        // THE FONT. Calibri 11 throughout, which is what the reference's default font is - and
        // stating it explicitly matters because a generated package does NOT inherit it
        // (owner report, 2026-09-24).
        //
        // EPPlus creates a workbook whose default font depends on its own defaults rather than on
        // Excel's, so a file written without setting it comes out in something else entirely and
        // looks wrong beside every hand-made board.
        // ###########################################################################################
        public const string FontName = "Calibri";

        public const float FontSize = 11f;

        // ###########################################################################################
        // EVERY COLUMN'S WIDTH, taken from the reference and keyed BY HEADER NAME.
        //
        // *** BY NAME RATHER THAN BY POSITION, deliberately. *** The reference still carries a
        // retired "UUID v4" column which the writer no longer emits, so the Nth written column is
        // not the Nth reference column on most sheets. Keying on the name means the mapping stays
        // right however the order changes - and it already has changed once, when CAD name moved
        // to the end of the schematics sheet to match the reference.
        //
        // A column with no entry here falls back to fitting its data, which is the right answer
        // for anything added later that the reference has never seen.
        // ###########################################################################################
        private static readonly Dictionary<string, double> ColumnWidths =
            new(StringComparer.OrdinalIgnoreCase)
            {
                // Board schematics
                ["Schematic name"] = 32.1,
                ["Schematic image file"] = 62.2,
                ["Schematic highlight color"] = 10.4,
                ["Schematic highlight opacity"] = 10.9,
                ["Opposite trace highlight color"] = 10.6,
                ["Thumbnail highlight color"] = 11.8,
                ["Thumbnail highlight opacity"] = 11.6,
                ["CAD name"] = 23.3,

                // Components
                ["Friendly name"] = 21.3,
                ["Technical name or value"] = 26.3,
                ["Part-number"] = 21.1,
                // The header genuinely carries an Alt+Enter line break in the workbook, so the key
                // does too - WidthFor normalises CRLF to LF before looking it up.
                ["Short one-liner description\n(one short line only!)"] = 41.2,

                // Component images
                ["Pin"] = 3.4,
                ["Expected oscilloscope reading"] = 22.8,
                ["T/DIV"] = 6.9,
                ["V/DIV"] = 7.0,
                ["T.LVL"] = 6.8,
                ["Note"] = 67.4,

                // Important signals
                ["Display name"] = 26.1,
                ["KiCad net name"] = 32.3,

                // Credits
                ["Sub-category"] = 24.6,
                ["Name or handle"] = 39.1,
                ["Contact (email or web page)"] = 60.1,
            };

        // ###########################################################################################
        // A few names appear on SEVERAL sheets at DIFFERENT widths - "Name" is 36.7 on component
        // local files and 46.0 on board local files - so those are keyed by sheet AND name.
        //
        // Looked up before the shared table above, so a per-sheet value always wins.
        // ###########################################################################################
        private static readonly Dictionary<string, double> ColumnWidthsBySheet =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Components|Board label"] = 15.0,
                ["Components|Category"] = 10.0,
                ["Components|Region"] = 10.0,

                ["Component images|Board label"] = 15.0,
                ["Component images|Region"] = 8.1,
                ["Component images|Name"] = 24.1,
                ["Component images|File"] = 73.7,

                ["Component local files|Board label"] = 12.3,
                ["Component local files|Name"] = 36.7,
                ["Component local files|File"] = 87.9,

                ["Component links|Board label"] = 12.2,
                ["Component links|Name"] = 39.6,
                ["Component links|URL"] = 100.3,

                ["Board local files|Category"] = 24.4,
                ["Board local files|Name"] = 46.0,
                ["Board local files|File"] = 107.8,

                ["Board links|Category"] = 25.8,
                ["Board links|Name"] = 55.9,
                ["Board links|URL"] = 86.6,

                ["Credits|Category"] = 35.1,
            };

        // The reference's width for this column, or null to fit the data instead.
        public static double? WidthFor(string sheetName, string columnName)
        {
            // CRLF to LF, so a header carrying an Alt+Enter break matches its key whichever line
            // ending the workbook happens to store.
            string normalised = (columnName ?? string.Empty)
                .Replace("\r\n", "\n", StringComparison.Ordinal);

            if (BoardWorkbookStyle.ColumnWidthsBySheet.TryGetValue(
                    sheetName + "|" + normalised, out double perSheet))
            {
                return perSheet;
            }

            return BoardWorkbookStyle.ColumnWidths.TryGetValue(normalised, out double shared)
                ? shared
                : null;
        }

        // What a column is worth on a sheet with no rows yet and no reference width. A brand-new
        // board has nothing to measure, and autofitting the wrapped headers instead would widen
        // every column to hold a whole header phrase on one line - the opposite of what the wrap
        // is for.
        public const double EmptySheetColumnWidth = 22;

        // White at 15% darker, which is what the column-header row actually shows.
        public const string HeaderFillRgb = "FFD9D9D9";

        public static string SectionTitleFor(string sheetName)
        {
            return BoardWorkbookStyle.SectionTitles.TryGetValue(sheetName, out string? title)
                ? title
                : sheetName;
        }

        // ###########################################################################################
        // THE PUBLISHED REVISION DATE, in the shape the boards already use: "2026-May-12".
        //
        // *** MONTH BY NAME, INVARIANT CULTURE. *** The value is hand-maintained free text that a
        // human reads at the top of a sheet, and the shipped boards all use this shape. Invariant
        // matters on a Danish machine, where "MMMM" renders a lowercase Danish month name and the
        // board would suddenly say "2026-maj-12".
        //
        // No leading zero on the day, matching SubmissionReceipt.FormatDate's own reasoning about
        // these dates being read by eye rather than sorted.
        // ###########################################################################################
        public static string FormatRevisionDate(DateTimeOffset published)
        {
            return published.ToString(
                "yyyy-MMMM-d",
                System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
