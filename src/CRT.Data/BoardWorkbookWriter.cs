using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Writes a BoardData back out as a board workbook (NewContributeStrategy.md Phase 5, task 6 -
    // the publishing step). The inverse of BoardDataReader, sharing its schema through
    // BoardWorkbookSchema so a column cannot be renamed on one side only.
    //
    // *** THIS IS NOT A RESTORATION OF BoardDataWriter. *** That class existed to let the APP
    // edit a board in place and was retired in Phase 2 session 2c, deliberately: authoring now
    // produces draft ROWS, and a second mechanism that mutated workbooks would have been a shadow
    // draft system beside draft.json. Nothing here changes that - this writer exists for the
    // SERVER, at publish time, writing a board that has already been reviewed and accepted. If a
    // future change makes the app want to write a workbook, that is a decision to re-open with the
    // project owner, not a call site to add.
    //
    // *** EVERY CELL IS WRITTEN AS TEXT, AND THIS IS THE SUBTLE PART. *** BoardData holds every
    // field as a string, including ones that look numeric: highlight opacities ("0.35"), scope
    // settings ("2.5"), trigger levels ("-1.2"). Writing those with EPPlus's ordinary Value
    // setter would let Excel store them as NUMBERS, and the reader takes .Text back - which is
    // formatted by the CURRENT CULTURE. On a comma-decimal machine "0.35" would round-trip as
    // "0,35", and the invariant-culture parse that reads it later would then fail or, worse,
    // produce 35. So cells are written as strings with an explicit Text number format. The same
    // class of bug is already documented twice in this codebase (highlight coordinates in
    // SubmissionValidator, and WorkbookSummary's Stat parts), and this is the third face of it.
    //
    // *** THE UUID COLUMN IS GONE ENTIRELY (2026-09-23). *** Phase 4 retired UuidV4 as an identity
    // - rows pair on natural keys - and this writer used to carry whatever a row held through
    // verbatim, on the strategy document's "stop WRITING new ones; keep READING existing ones".
    // The project owner has since dropped the column from the data outright: nothing generated one,
    // nothing compared one, and the only code still touching it was a validator warning that a
    // dead field needed fixing. So a written workbook has no UUID column at all.
    //
    // *** THE WORKBOOK IS REPLACED, NEVER WRITTEN IN PLACE (owner report, 2026-09-28). *** It is
    // written complete beside the target and renamed over it (FileReplacer), so a reader mid-sync
    // never sees half a workbook, and a workbook copied into the tree by hand as another user -
    // which the service may replace but not open for writing - is no obstacle. This used to be
    // left to the caller, and the publishing step wrote straight onto the served path; it only
    // survived hand-copied files because EPPlus happens to delete an existing file first.
    // ###########################################################################################
    public static class BoardWorkbookWriter
    {
        // ###########################################################################################
        // *** THE HEADER ROW IS NO LONGER A CONSTANT (owner request, 2026-09-23). ***
        //
        // It used to be row 2 on every sheet, with a bare revision-date marker above it on the
        // schematics sheet and nothing else anywhere. That produced a workbook carrying only data:
        // no identity lines, no documentation links, no shaded header band, no column widths. The
        // first real publish turned a 190 KB hand-maintained board into an 81 KB stripped one, and
        // the project owner asked for the presentation to be kept.
        //
        // The header now sits below a PREAMBLE whose height differs by one row between the
        // schematics sheet (which carries the revision date) and the rest, so the position comes
        // from BoardWorkbookStyle.HeaderRowFor rather than from a constant here.
        //
        // NOTHING ABOUT READING DEPENDS ON IT. BoardDataReader.ReadSheetRows scans from row 1 for
        // the row carrying the required column names, and ScanRevisionDate searches the schematics
        // sheet's top-left 10x10 - so the preamble is free to move without breaking a reader.
        // ###########################################################################################

        // ###########################################################################################
        // Writes the whole board to a new workbook at the given path, replacing anything there.
        //
        // The caller is responsible for the path being the right GENERATION - see
        // DataGenerationRules, and the project owner's rule that publishing writes only the newest
        // one and never an older, frozen one.
        // ###########################################################################################
        public static void Write(string excelPath, BoardData data)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
                throw new ArgumentException("A workbook path is required.", nameof(excelPath));

            ArgumentNullException.ThrowIfNull(data);

            EpplusLicense.Ensure();

            using var package = new ExcelPackage();

            // ###########################################################################################
            // *** THE DEFAULT FONT IS SET EXPLICITLY (owner report, 2026-09-24). ***
            //
            // A generated package does not inherit Excel's Calibri 11 - EPPlus applies its own
            // default - so a workbook written without this came out in a different typeface from
            // every hand-made board. Set once on the workbook rather than per cell, which is both
            // cheaper and what makes a hand-typed cell inherit it too.
            // ###########################################################################################
            package.Workbook.Styles.UpdateXml();
            package.Workbook.Styles.NamedStyles[0].Style.Font.Name = BoardWorkbookStyle.FontName;
            package.Workbook.Styles.NamedStyles[0].Style.Font.Size = BoardWorkbookStyle.FontSize;

            foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
            {
                BoardWorkbookWriter.WriteSheet(package, sheet, data);
            }

            var file = new FileInfo(excelPath);

            // A publish target's folder may not exist yet - a brand-new system is the ordinary
            // case, not an exception.
            file.Directory?.Create();

            // Complete and deterministic at a temporary path first, then renamed into place - see
            // the class header.
            FileReplacer.Replace(file.FullName, temporary =>
            {
                package.SaveAs(new FileInfo(temporary));
                BoardWorkbookWriter.MakeDeterministic(temporary);
            });

            CrtLog.Info($"Wrote board workbook [{excelPath}]");
        }

        // ###########################################################################################
        // THE SAME BOARD MUST PRODUCE THE SAME BYTES, EVERY TIME (found 2026-09-22).
        //
        // *** AN .xlsx IS A ZIP, AND EPPlus STAMPS EVERY ENTRY WITH THE CURRENT CLOCK. *** So two
        // publishes of an identical board, seconds apart, produced byte-different workbooks and
        // therefore different SHA-256 hashes. Nothing in the file's CONTENT differed - only the
        // timestamps in the archive directory.
        //
        // That is not cosmetic, because the workbook's hash is folded into the system's
        // ContentHash (PublishPlan.DescriptorWithWorkbook), which is what every CRT client uses to
        // decide whether to re-download a board. The consequences, in order of how much they cost:
        //
        //   - RE-PUBLISHING AN UNCHANGED SYSTEM MOVED ITS CONTENT HASH, so every user on every
        //     machine re-downloaded a board that had not changed. SystemDescriptorRules already
        //     states the intended guarantee in its own comments ("republishing an unchanged system
        //     does not produce a file that differs"); the ZIP timestamp silently broke it.
        //   - Re-running an interrupted publish - the documented recovery - produced a different
        //     result from the one it was recovering, which is the opposite of idempotent.
        //   - It surfaced as PublishExecutorTests.Re_running_the_same_publish_is_safe failing
        //     intermittently. The test was RIGHT and the code was wrong; it only looked flaky
        //     because two publishes inside the same clock second happen to agree.
        //
        // *** THE FIX IS A FIXED TIMESTAMP ON EVERY ENTRY, NOT A RE-COMPRESSION. *** The entries
        // are copied across verbatim, so the compressed bytes are untouched and only the archive's
        // own metadata is normalised. The constant is arbitrary but must never change: moving it
        // would re-hash every published workbook at once and trigger a download of the entire data
        // tree for every user.
        //
        // Failure is SOFT. A workbook that could not be normalised is still a perfectly valid
        // workbook - it just hashes differently next time - and losing the board over a tidy-up
        // would be far worse than the re-download it prevents.
        // ###########################################################################################
        private static void MakeDeterministic(string excelPath)
        {
            // 2000-01-01, comfortably inside the range a ZIP's DOS timestamp can express (it
            // cannot represent anything before 1980) and old enough that no tool mistakes it for
            // "just modified".
            var fixedStamp = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);

            try
            {
                string temporary = excelPath + ".deterministic.tmp";

                using (var source = ZipFile.Open(excelPath, ZipArchiveMode.Read))
                using (var destination = ZipFile.Open(temporary, ZipArchiveMode.Create))
                {
                    // *** ORDER IS PRESERVED, not sorted. *** An .xlsx is not an arbitrary bag of
                    // files: [Content_Types].xml must come first, and EPPlus writes the rest in an
                    // order Excel is known to accept. Sorting them would produce a file that hashes
                    // consistently and might not open.
                    foreach (ZipArchiveEntry entry in source.Entries)
                    {
                        ZipArchiveEntry copy = destination.CreateEntry(entry.FullName);
                        copy.LastWriteTime = fixedStamp;

                        using Stream input = entry.Open();
                        using Stream output = copy.Open();

                        input.CopyTo(output);
                    }
                }

                File.Move(temporary, excelPath, overwrite: true);
            }
            catch (Exception ex)
            {
                CrtLog.Warning(
                    $"Could not normalise workbook timestamps [{excelPath}]: [{ex.Message}] - " +
                    "the workbook is valid but may hash differently on the next publish");
            }
        }

        // ###########################################################################################
        // Writes one sheet: the revision-date marker (schematics sheet only), the header row, then
        // one row per entry.
        //
        // A SHEET WITH NO ROWS IS STILL CREATED, WITH ITS HEADERS. The reader warns and returns an
        // empty list for a missing sheet, which is survivable - but a board that genuinely has no
        // credits should read as "no credits", not as "this workbook is malformed". Writing the
        // header also means the next person to edit it by hand has the columns in front of them.
        // ###########################################################################################
        private static void WriteSheet(
            ExcelPackage package,
            BoardWorkbookSchema.SheetDefinition sheet,
            BoardData data)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheet.SheetName);

            int headerRow = BoardWorkbookWriter.WritePreamble(worksheet, sheet, data);
            int columnCount = sheet.ColumnOrder.Count;

            // ---- the column headers ------------------------------------------------------------
            for (int column = 0; column < columnCount; column++)
            {
                BoardWorkbookWriter.WriteTextCell(
                    worksheet,
                    headerRow,
                    column + 1,
                    sheet.ColumnOrder[column]);
            }

            // ###########################################################################################
            // *** THE HEADERS ARE NOT BOLD, and they WRAP (corrected 2026-09-24). ***
            //
            // Both were wrong in the first version and the project owner spotted them side by side.
            // Nothing on the reference's title or header rows is bold - the shading is what marks
            // them out - and the headers wrap onto up to three lines at a fixed row height, which
            // is what keeps the sheet a readable width. Without the wrap, "Schematic highlight
            // opacity" widens its column to fit on one line and the sheet ends up three times
            // wider than every shipped board.
            // ###########################################################################################
            ExcelRange headerBand = worksheet.Cells[headerRow, 1, headerRow, columnCount];
            headerBand.Style.WrapText = true;
            BoardWorkbookWriter.Fill(headerBand, BoardWorkbookStyle.HeaderFillRgb);

            worksheet.Row(headerRow).Height = BoardWorkbookStyle.HeaderRowHeight;

            // *** THE CAD NAME COLUMN IS BLUE, on every sheet that has one. *** It is the one
            // column filled in by the KiCad import rather than by hand, and the reference picks it
            // out so a reader can see at a glance which column is not theirs to type into.
            int cadColumn = BoardWorkbookWriter.ColumnIndex(sheet, BoardWorkbookSchema.ColCadName);

            if (cadColumn >= 0)
            {
                BoardWorkbookWriter.Fill(
                    worksheet.Cells[headerRow, cadColumn + 1],
                    BoardWorkbookStyle.HeaderCadFillRgb);
            }

            // ---- the data ----------------------------------------------------------------------
            IReadOnlyList<IReadOnlyDictionary<string, string>> rows =
                BoardWorkbookSchema.BuildRows(sheet, data);

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                IReadOnlyDictionary<string, string> row = rows[rowIndex];

                for (int column = 0; column < columnCount; column++)
                {
                    string value = row.TryGetValue(sheet.ColumnOrder[column], out string? cell)
                        ? cell
                        : string.Empty;

                    BoardWorkbookWriter.WriteTextCell(
                        worksheet,
                        headerRow + 1 + rowIndex,
                        column + 1,
                        value);
                }
            }

            // ###########################################################################################
            // Column widths, fitted to the DATA rather than to the preamble.
            //
            // *** AUTOFIT OVER THE WHOLE SHEET WOULD BE RUINED BY THE PREAMBLE. *** Column A now
            // holds a 100-character documentation URL, so fitting the used range would widen that
            // column to the maximum on every sheet and push every other column off the screen. The
            // fit is therefore measured over the header and data block only.
            // ###########################################################################################
            // ###########################################################################################
            // *** AND IT EXCLUDES THE HEADER ROW TOO (corrected 2026-09-24). ***
            //
            // The headers WRAP, so measuring them defeats the wrap entirely: autofit widens each
            // column until its header fits on one line, which is the very thing the wrap exists to
            // avoid. Fitting the DATA alone is what gives the reference its compact columns with
            // three-line headers above them.
            //
            // A sheet with no data keeps a sensible default width rather than collapsing to the
            // minimum - there is nothing to measure, and a brand-new system is exactly that case.
            // ###########################################################################################
            // ###########################################################################################
            // *** THE REFERENCE'S OWN WIDTHS, column by column (owner request, 2026-09-24). ***
            //
            // Fitting the data produced sensible-but-different columns, and the project owner asked for
            // the shipped boards' actual widths. They are keyed by HEADER NAME in
            // BoardWorkbookStyle - see there for why position would be wrong.
            //
            // Anything the reference has no width for falls back to fitting its data, so a column
            // added later is still readable rather than stuck at a default.
            // ###########################################################################################
            var unsized = new List<int>();

            for (int column = 0; column < columnCount; column++)
            {
                double? width = BoardWorkbookStyle.WidthFor(sheet.SheetName, sheet.ColumnOrder[column]);

                if (width.HasValue)
                {
                    worksheet.Column(column + 1).Width = width.Value;
                }
                else
                {
                    unsized.Add(column + 1);
                }
            }

            foreach (int column in unsized)
            {
                if (rows.Count > 0)
                {
                    worksheet.Cells[headerRow + 1, column, headerRow + rows.Count, column]
                        .AutoFitColumns(10, 60);
                }
                else
                {
                    worksheet.Column(column).Width = BoardWorkbookStyle.EmptySheetColumnWidth;
                }
            }
        }

        // ###########################################################################################
        // The identity block, documentation links and section title above the header, matching the
        // hand-maintained boards. Returns the row the column headers go on.
        //
        // See BoardWorkbookStyle for where every value here came from - all of it was read off the
        // reference board rather than invented, so a generated workbook and a hand-made one look
        // the same.
        // ###########################################################################################
        private static int WritePreamble(
            ExcelWorksheet worksheet,
            BoardWorkbookSchema.SheetDefinition sheet,
            BoardData data)
        {
            int row = BoardWorkbookStyle.PreambleHardwareRow;

            // The rows are always RESERVED even when a name is unknown, so the block keeps its
            // shape and the header lands on the same row for every board - but the marker itself is
            // only written when there is something to say. "# Hardware: " with nothing after it is
            // a caption that has lost its subject.
            BoardWorkbookWriter.WriteIdentityCell(
                worksheet, row++, BoardWorkbookSchema.HardwareMarkerPrefix, data.HardwareName);

            BoardWorkbookWriter.WriteIdentityCell(
                worksheet, row++, BoardWorkbookSchema.BoardMarkerPrefix, data.BoardName);

            if (BoardWorkbookStyle.HasRevisionDateRow(sheet.SheetName))
            {
                // ###########################################################################################
                // *** THE DATE ITSELF IS BOLD, THE LABEL IS NOT - which needs RICH TEXT. ***
                //
                // The reference stores this cell as two runs: a plain "# Revision date: " followed
                // by the date in bold. A single bolded cell would be visibly different, and a plain
                // one loses the emphasis the project owner asked to keep.
                //
                // The reader is unaffected either way: ScanRevisionDate takes the cell's TEXT, which
                // is the concatenation of the runs, so the formatting is presentation only.
                // ###########################################################################################
                // The same two-run shape as the hardware and board lines above - one helper, so the
                // three cannot drift apart again.
                BoardWorkbookWriter.WriteRichLabelAndValue(
                    worksheet,
                    row,
                    BoardWorkbookSchema.RevisionDateMarkerPrefix,
                    (data.RevisionDate ?? string.Empty).Trim());

                row++;
            }

            row++;   // the blank line the reference leaves under the identity block

            BoardWorkbookWriter.WriteTextCell(worksheet, row++, 1, BoardWorkbookStyle.DocumentationLeadIn);
            BoardWorkbookWriter.WriteTextCell(
                worksheet, row++, 1, BoardWorkbookStyle.DocumentationUrlFor(sheet.SheetName));

            row++;   // and the blank line under the documentation link

            // The section title band, directly above the headers.
            BoardWorkbookWriter.WriteTextCell(
                worksheet, row, 1, BoardWorkbookStyle.SectionTitleFor(sheet.SheetName));

            // Not bold - see the header band below. The shading carries the emphasis.
            ExcelRange titleBand = worksheet.Cells[row, 1, row, sheet.ColumnOrder.Count];
            BoardWorkbookWriter.Fill(titleBand, BoardWorkbookStyle.SectionTitleFillRgb);

            BoardWorkbookWriter.ApplyTitleBanding(worksheet, sheet, row);

            row++;

            return row;
        }

        // ###########################################################################################
        // The schematics sheet's three-colour title band (owner report, 2026-09-24).
        //
        // Every other sheet's title row is a plain white strip carrying the sheet name, which the
        // generic path above has already written. This one groups its columns instead - see
        // BoardWorkbookStyle for the bands and why the colours are literal RGB - and the middle
        // band carries a heading of its own spanning the five highlight columns.
        //
        // Silently does nothing for a sheet with no CAD-name column, which is every sheet but this
        // one, so the caller needs no branch of its own.
        // ###########################################################################################
        private static void ApplyTitleBanding(
            ExcelWorksheet worksheet,
            BoardWorkbookSchema.SheetDefinition sheet,
            int titleRow)
        {
            int cadColumn = BoardWorkbookWriter.ColumnIndex(sheet, BoardWorkbookSchema.ColCadName);

            if (cadColumn < 0)
            {
                return;
            }

            // The highlight columns are everything between the image file and the CAD name - taken
            // from the column order rather than hardcoded, so adding a highlight column extends the
            // band instead of leaving it painted over the wrong cells.
            int firstHighlight = BoardWorkbookWriter.ColumnIndex(
                sheet, BoardWorkbookSchema.ColSchematicHighlightColor);

            if (firstHighlight < 0 || firstHighlight >= cadColumn)
            {
                return;
            }

            // A-B: the sheet's own title, in black with white text.
            ExcelRange titlePart = worksheet.Cells[titleRow, 1, titleRow, firstHighlight];
            BoardWorkbookWriter.Fill(titlePart, BoardWorkbookStyle.TitleBandBlackRgb);
            titlePart.Style.Font.Color.SetColor(System.Drawing.Color.White);

            // C-G: the highlight columns, under one heading of their own.
            ExcelRange highlightPart =
                worksheet.Cells[titleRow, firstHighlight + 1, titleRow, cadColumn];

            BoardWorkbookWriter.Fill(highlightPart, BoardWorkbookStyle.TitleBandGreyRgb);
            highlightPart.Style.Font.Color.SetColor(System.Drawing.Color.White);

            BoardWorkbookWriter.WriteTextCell(
                worksheet, titleRow, firstHighlight + 1, BoardWorkbookStyle.HighlightsBandTitle);

            // ###########################################################################################
            // *** THE BAND IS MERGED, which is WHY it is centred (corrected 2026-09-24). ***
            //
            // The reference merges C8:G8 into one cell spanning the five highlight columns, and the
            // centring is that merged cell's own alignment. Centring a single unmerged cell instead
            // - which is what the first version did - puts the text in the middle of ONE column and
            // reads as a mistake, which the project owner said outright.
            //
            // Merging reproduces the reference exactly, so the centring is right again rather than
            // being something to remove.
            // ###########################################################################################
            worksheet.Cells[titleRow, firstHighlight + 1, titleRow, cadColumn].Merge = true;

            worksheet.Cells[titleRow, firstHighlight + 1].Style.HorizontalAlignment =
                ExcelHorizontalAlignment.Center;

            worksheet.Cells[titleRow, firstHighlight + 1].Style.Font.Color
                .SetColor(System.Drawing.Color.White);

            // H: the CAD name column, pale blue.
            BoardWorkbookWriter.Fill(
                worksheet.Cells[titleRow, cadColumn + 1],
                BoardWorkbookStyle.TitleBandCadRgb);
        }

        // The position of a named column in this sheet's write order, or -1. Ordinal, matching how
        // the reader maps headers, and written out because IReadOnlyList has no IndexOf.
        private static int ColumnIndex(BoardWorkbookSchema.SheetDefinition sheet, string columnName)
        {
            for (int i = 0; i < sheet.ColumnOrder.Count; i++)
            {
                if (string.Equals(sheet.ColumnOrder[i], columnName, StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // "# Hardware: " in plain text followed by the NAME IN BOLD (corrected 2026-09-24).
        //
        // The reference stores all three identity lines this way - two runs in one cell - and only
        // the revision date was being written like that. The other two came out entirely plain,
        // which the project owner spotted against the real file.
        // ###########################################################################################
        private static void WriteIdentityCell(
            ExcelWorksheet worksheet,
            int row,
            string markerPrefix,
            string? value)
        {
            string trimmed = (value ?? string.Empty).Trim();

            if (trimmed.Length == 0)
            {
                return;
            }

            BoardWorkbookWriter.WriteRichLabelAndValue(worksheet, row, markerPrefix, trimmed);
        }

        // ###########################################################################################
        // One preamble cell as TWO RUNS: a plain label and a bold value.
        //
        // Shared by all three identity lines so they cannot drift apart - the revision date used to
        // have its own copy of this, which is how the other two ended up plain.
        // ###########################################################################################
        private static void WriteRichLabelAndValue(
            ExcelWorksheet worksheet,
            int row,
            string markerPrefix,
            string value)
        {
            ExcelRange cell = worksheet.Cells[row, 1];
            cell.Style.Numberformat.Format = "@";

            cell.IsRichText = true;
            cell.RichText.Clear();

            ExcelRichText label = cell.RichText.Add(markerPrefix + " ");
            label.Bold = false;
            label.Size = BoardWorkbookStyle.IdentityFontSize;
            label.FontName = BoardWorkbookStyle.FontName;

            if (value.Length == 0)
            {
                return;
            }

            ExcelRichText bold = cell.RichText.Add(value);
            bold.Bold = true;
            bold.Size = BoardWorkbookStyle.IdentityFontSize;
            bold.FontName = BoardWorkbookStyle.FontName;
        }

        // A solid fill in one literal colour - see BoardWorkbookStyle for why these are RGB rather
        // than the theme references the reference workbook stores.
        private static void Fill(ExcelRange range, string argb)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(
                System.Drawing.ColorTranslator.FromHtml("#" + argb.Substring(2)));
        }

        // ###########################################################################################
        // Writes one cell as TEXT - see the class header. Both halves matter: the number format
        // stops Excel coercing a numeric-looking string, and assigning to .Value as a string stops
        // EPPlus inferring a type from the content.
        //
        // A blank is written as an EMPTY cell rather than an empty string, so a freshly written
        // workbook has no cells that exist only to hold nothing. The reader answers "" for both.
        // ###########################################################################################
        private static void WriteTextCell(ExcelWorksheet worksheet, int row, int column, string? value)
        {
            if (string.IsNullOrEmpty(value))
                return;

            ExcelRange cell = worksheet.Cells[row, column];

            cell.Style.Numberformat.Format = "@";
            cell.Value = value;
        }
    }
}
