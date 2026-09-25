using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The CONTENT of a workbook export, assembled and ordered, with no idea how it will be drawn.
    //
    // WHY THIS IS SEPARATE FROM THE PDF WRITER: what goes into a customer's document, in what
    // order, under what headings, and what an absent field turns into, are all decisions worth
    // testing - and none of them need QuestPDF, a display, or a file on disk. WorkbookPdfExporter
    // takes one of these and only paints it. The ZIP export prints the same model, so the PDF
    // inside the archive and the archive's own manifest cannot describe different jobs.
    //
    // Pure apart from resolving attachment PATHS (string work on paths the caller supplies; it
    // reads no file - the writer does). Everything else is the records the caller already read.
    // ###########################################################################################
    public static class WorkbookExportModel
    {
        // What the customer sees at the top of the document and in the archive's file name.
        public sealed class Document
        {
            public string Title { get; init; } = string.Empty;

            public string BoardKey { get; init; } = string.Empty;

            public string Note { get; init; } = string.Empty;

            public string Status { get; init; } = string.Empty;

            public int WorkbookId { get; init; }

            public DateTime StartDate { get; init; }

            // When this workbook was closed, or null if it never was - so the header can report a
            // finished job by when it FINISHED, exactly as the worklog bar and the workbook cards
            // do on screen. An exported document that says "started" for a workbook the app shows
            // as "ended" is the same disagreement, just harder to notice because the PDF is read
            // away from the app. Nullable for the reason WorkbookRecord.EndDate is: null is "never
            // closed", which every workbook written before that field existed deserializes to.
            public DateTime? EndDate { get; init; }

            public DateTime GeneratedAt { get; init; }

            public WorkbookSummary.Totals Totals { get; init; } = WorkbookSummary.Summarize(null);

            // The ISO 4217 code every cost in this document is printed with - captured when the
            // document is BUILT rather than read from UserSettings while it is painted. The paint
            // runs on a background thread (the export handler's Task.Run), and a setting the user
            // could change mid-write would let one document print two currencies. It also keeps
            // WorkbookPdfExporter free of a settings read, so its tests need no settings file.
            public string CurrencyCode { get; init; } = WorklogCurrency.DefaultCode;

            // Grouped by schematic and ordered, so the document walks the board a section at a
            // time rather than jumping between schematics in entry-id order - a repair reads as a
            // narrative per area of the board.
            public IReadOnlyList<Section> Sections { get; init; } = Array.Empty<Section>();
        }

        public sealed class Section
        {
            public string SchematicName { get; init; } = string.Empty;

            // The schematic's own image, when the board data names one that exists. Null is normal
            // and must not stop the export: entries whose schematic image is missing from the data
            // root are still worth printing, just without the picture.
            public string? SchematicImagePath { get; init; }

            public IReadOnlyList<Entry> Entries { get; init; } = Array.Empty<Entry>();
        }

        public sealed class Entry
        {
            public WorklogEntryRecord Record { get; init; } = new();

            // Absolute paths to the entry's photo files, in the entry's own display order, with
            // any whose file is missing from disk already dropped - the writer must never be
            // handed a path it then has to decide about.
            public IReadOnlyList<Attachment> Photos { get; init; } = Array.Empty<Attachment>();

            public IReadOnlyList<Attachment> Files { get; init; } = Array.Empty<Attachment>();
        }

        public sealed class Attachment
        {
            public string FileName { get; init; } = string.Empty;

            public string Comment { get; init; } = string.Empty;

            public string FullPath { get; init; } = string.Empty;
        }

        // ###########################################################################################
        // Assembles the document for one workbook.
        //
        // schematicImagePathsByName maps a schematic name to its image path on disk - the caller
        // resolves these from board data, because this class deliberately knows nothing about
        // BoardData or DataManager's data root. Names are matched case-insensitively, matching how
        // schematic names are compared everywhere else in this app.
        //
        // attachmentsFolderForEntry resolves an entry's "worklog_<id>" folder. Passed in as a
        // function rather than called directly so a test can point it at a temp folder without
        // WorklogManager's real AppData root.
        //
        // An entry whose SchematicName is blank still appears, under the UnassignedSectionName
        // heading declared below - it is the user's own worklog, and silently dropping it from
        // their customer document is the one failure mode this export must not have.
        // ###########################################################################################
        public static Document Build(
            WorkbookRecord workbook,
            IEnumerable<WorklogEntryRecord> entries,
            IReadOnlyDictionary<string, string>? schematicImagePathsByName,
            Func<int, string?> attachmentsFolderForEntry,
            DateTime generatedAt,
            string? currencyCode)
        {
            var entryList = (entries ?? Array.Empty<WorklogEntryRecord>())
                .Where(e => e != null)
                .ToList();

            var sections = entryList
                .GroupBy(
                    e => string.IsNullOrWhiteSpace(e.SchematicName) ? UnassignedSectionName : e.SchematicName,
                    StringComparer.OrdinalIgnoreCase)

                // Sections alphabetically, entries within a section by id: the id is the number the
                // pills on screen show, so a reader holding the app open follows the same order.
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .Select(g => new Section
                {
                    SchematicName = g.Key,
                    SchematicImagePath = ResolveSchematicImage(schematicImagePathsByName, g.Key),
                    Entries = g.OrderBy(e => e.Id)
                        .Select(e => BuildEntry(e, attachmentsFolderForEntry))
                        .ToList()
                })
                .ToList();

            return new Document
            {
                WorkbookId = workbook.Id,
                Title = string.IsNullOrWhiteSpace(workbook.Title) ? "(untitled)" : workbook.Title,
                BoardKey = workbook.BoardKey,
                Note = workbook.Note ?? string.Empty,
                Status = workbook.Status,
                StartDate = workbook.StartDate,
                EndDate = workbook.EndDate,
                GeneratedAt = generatedAt,
                Totals = WorkbookSummary.Summarize(entryList),
                CurrencyCode = WorklogCurrency.NormalizeCode(currencyCode),
                Sections = sections
            };
        }

        // The heading an entry with no schematic name is filed under. Parenthesised, matching the
        // "(untitled)" / "(no description)" placeholders the rest of the worklog UI uses, so it
        // reads as the app describing an absence rather than as a schematic actually called this.
        public const string UnassignedSectionName = "(no schematic)";

        private static Entry BuildEntry(WorklogEntryRecord entry, Func<int, string?> attachmentsFolderForEntry)
        {
            string? folder = null;
            try
            {
                folder = attachmentsFolderForEntry(entry.Id);
            }
            catch (Exception)
            {
                // A resolver that throws (an unreadable workbook folder) costs this entry its
                // attachments, not the whole export. The text of the entry is the part the
                // customer document cannot do without.
                folder = null;
            }

            return new Entry
            {
                Record = entry,
                Photos = ResolveAttachments(entry.Photos, folder),
                Files = ResolveAttachments(entry.Files, folder)
            };
        }

        // ###########################################################################################
        // Turns attachment records into absolute paths, dropping any whose file is not on disk.
        //
        // Dropped rather than passed through as a broken path: an attachment row is written to
        // entries.json when the file is added, and the file itself can go missing afterwards (a
        // hand-edited workbook folder, a failed copy, a restore from a partial backup). The export
        // must produce a document either way, and a page reading "photo-3.jpg" over a blank frame
        // is worse than a document that simply does not mention it.
        //
        // Ordered by DisplayOrder, matching how the editor lists them, so the exported document
        // shows the photos in the sequence the user arranged.
        // ###########################################################################################
        private static List<Attachment> ResolveAttachments(IEnumerable<WorklogAttachmentRecord>? records, string? folder)
        {
            var result = new List<Attachment>();
            if (records == null || string.IsNullOrWhiteSpace(folder))
                return result;

            foreach (var record in records.OrderBy(r => r.DisplayOrder))
            {
                if (string.IsNullOrWhiteSpace(record.FileName))
                    continue;

                string path;
                try
                {
                    path = Path.Combine(folder, record.FileName);
                    if (!File.Exists(path))
                        continue;
                }
                catch (Exception)
                {
                    // An invalid character in a stored file name makes Path.Combine throw. Same
                    // reasoning as a missing file: drop the row, keep the export.
                    continue;
                }

                result.Add(new Attachment
                {
                    FileName = record.FileName,
                    Comment = record.Comment ?? string.Empty,
                    FullPath = path
                });
            }

            return result;
        }

        private static string? ResolveSchematicImage(IReadOnlyDictionary<string, string>? paths, string schematicName)
        {
            if (paths == null || string.IsNullOrWhiteSpace(schematicName))
                return null;

            if (!paths.TryGetValue(schematicName, out var path) || string.IsNullOrWhiteSpace(path))
                return null;

            try
            {
                return File.Exists(path) ? path : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        // ###########################################################################################
        // The suggested file name for an export, without an extension:
        // "Workbook_3_C64_250469_20260904" - id, hardware, board, and the date it was produced.
        //
        // The BoardKey it splits is "Hardware|Board" (Main.GetCurrentBoardKey), so the two halves
        // become their own underscore-separated segments; a key with no separator contributes one
        // segment, and an empty one contributes none rather than leaving "__" in the middle.
        //
        // The workbook TITLE is deliberately left out: it is a sentence ("SN: HB4 1325119; bought
        // 2023-08-15 from Esben"), which makes an unwieldy file name and routinely carries a
        // customer's own details - on a file that is about to be emailed. Id, board and date
        // identify it unambiguously.
        //
        // YYYYMMDD with no separators, and underscores between segments, so the name sorts
        // chronologically per board in a folder listing and carries no character that needs
        // escaping in a shell or a URL.
        // ###########################################################################################
        // ###########################################################################################
        // The chosen path with the export's OWN extension on it.
        //
        // The format comes from which button was pressed, not from what the user typed, so a name
        // carrying a DIFFERENT known export extension has it replaced rather than appended to:
        // typing "repair.pdf" into the ZIP dialog produced "repair.pdf.zip", and - worse - the
        // picker's overwrite prompt had been shown for "repair.pdf", so an existing
        // "repair.pdf.zip" was silently overwritten without asking.
        //
        // Only the two extensions this feature owns are replaced. An unrelated one is APPENDED, not
        // stripped: a workbook titled "Rev 2.5" saved as "board rev 2.5" must not lose its tail to
        // something that merely looks like an extension.
        // ###########################################################################################
        public static string EnsureFileExtension(string path, string extension)
        {
            if (string.IsNullOrWhiteSpace(path))
                return path;

            string wanted = "." + extension.TrimStart('.');

            if (path.EndsWith(wanted, StringComparison.OrdinalIgnoreCase))
                return path;

            // The other format's extension, replaced rather than stacked on.
            foreach (string owned in OwnedExtensions)
            {
                if (path.EndsWith(owned, StringComparison.OrdinalIgnoreCase))
                    return path.Substring(0, path.Length - owned.Length) + wanted;
            }

            return path + wanted;
        }

        // The extensions this export owns, and therefore the only ones it may replace.
        private static readonly string[] OwnedExtensions = { ".pdf", ".zip" };

        public static string BuildFileBaseName(Document document)
        {
            var segments = new List<string>
            {
                "Workbook",
                document.WorkbookId.ToString(CultureInfo.InvariantCulture)
            };

            // '|' is the separator Main.GetCurrentBoardKey joins on. Split rather than sanitised
            // into a single token, so "C64/250469"-style keys yield "C64_250469" rather than a
            // hyphenated run - each half is a name in its own right.
            foreach (string part in (document.BoardKey ?? string.Empty).Split('|'))
            {
                string clean = SanitizeForFileName(part);
                if (!string.IsNullOrEmpty(clean))
                    segments.Add(clean);
            }

            segments.Add(document.GeneratedAt.ToString("yyyyMMdd", CultureInfo.InvariantCulture));

            return string.Join("_", segments);
        }

        // ###########################################################################################
        // Reduces free text to a file-name-safe token: the invalid characters of EVERY platform (not
        // just this one - Path.GetInvalidFileNameChars is host-specific, and an export made on Linux
        // may well be opened on Windows), plus the path separators, collapsed to single hyphens.
        //
        // Applied PER SEGMENT by BuildFileBaseName, which joins the results with underscores - so a
        // hyphen here can never be confused with the separator between segments.
        // ###########################################################################################
        public static string SanitizeForFileName(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var chars = text.Trim().Select(c =>
                char.IsLetterOrDigit(c) || c == '-' || c == '_' || c == '.' ? c : '-').ToArray();

            string collapsed = new string(chars);
            while (collapsed.Contains("--", StringComparison.Ordinal))
                collapsed = collapsed.Replace("--", "-", StringComparison.Ordinal);

            return collapsed.Trim('-', '.');
        }
    }
}
