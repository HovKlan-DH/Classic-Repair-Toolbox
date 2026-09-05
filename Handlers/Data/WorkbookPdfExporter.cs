using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Avalonia.Platform;
using Handlers.Geometry;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Renders a WorkbookExportModel.Document to a PDF, and packs a PDF plus the workbook's original
    // attachments into a ZIP.
    //
    // WHAT THIS IS FOR: handing a finished repair to the customer who paid for it. That shapes
    // every decision here - it prints the worklog as a readable narrative per schematic, shows the
    // board image with the marked areas, and totals the hours and cost. It is NOT a backup format:
    // it deliberately drops the app's internal ids, coordinates and display orders, which mean
    // nothing to the recipient. (The ZIP's attachment copies are the closest thing to a backup, and
    // are there so the customer HAS the photos, not so the workbook could be reconstructed.)
    //
    // The document model does the deciding (what goes in, in what order, what an absent field
    // becomes - see WorkbookExportModel); this file only paints it, which is why the model is unit
    // tested and this is not: everything below needs QuestPDF to produce bytes, and asserting on
    // PDF bytes tests the library rather than this app.
    //
    // QuestPDF requires QuestPDF.Settings.License to be set before it generates anything, and
    // throws otherwise. It is set once at startup in App.OnFrameworkInitializationCompleted rather
    // than here, so a missing licence line fails on launch in development rather than in a user's
    // hands at their first export.
    // ###########################################################################################
    public static class WorkbookPdfExporter
    {
        // Muted greys for the labels and rules, chosen to print legibly on a monochrome printer -
        // this document is as likely to be printed and handed over as it is to be emailed, so the
        // layout carries the structure and colour only reinforces it. The category/state colours
        // are the app's own, resolved by the caller-independent helpers below rather than from the
        // theme: an exported document must look the same whatever theme the app happens to be in.
        private const string LabelColor = "#666666";

        private const string RuleColor = "#DDDDDD";

        private const string HeadingColor = "#222222";

        private const string PanelColor = "#F7F7F7";

        // The 1px outline drawn around every schematic image, matching the black boundary the
        // Workbooks board pane draws around each preview - a schematic with a lot of white in it is
        // otherwise hard to tell apart from the page behind it.
        private const string ImageOutlineColor = "#000000";

        // An informational pill's fill. The app resolves Form_Bg from the theme; an exported
        // document is a fixed artefact that gets printed, so it takes the light theme's value.
        private const string PillFillColor = "#FFFFFF";

        // Corner radii, in points, mirroring WorklogInfoPillBuilder's own 10px pill / 3px chip
        // split - a status pill is fully rounded, a category chip only softened, and that
        // difference is how the two are told apart when they sit side by side.
        private const float PillCornerRadius = 7f;

        private const float ChipCornerRadius = 2f;

        private const float BadgeCornerRadius = 7f;

        // The white disc inside a "#N" badge, and the padlock in it. Sized so the disc reads as a
        // circle at the badge's own small scale rather than as a dot.
        private const float StateDiscSize = 9f;

        private const float StateDiscGlyphSize = 5f;

        // The marked area's outline. Thinner than the on-screen 1px because the exported image is
        // drawn far larger than any preview, where a full point of border starts to obscure the
        // copper underneath it.
        private const float AreaBorderThickness = 0.75f;

        // A web link's colour. Blue and underlined is the convention every reader already knows,
        // and the ONLY thing that marks a link in a printed or exported document - a PDF viewer
        // shows no hover cue and no status bar, so an unstyled hyperlink is indistinguishable from
        // prose until someone happens to click it. On paper it is the only cue that survives at all.
        private const string LinkColor = "#1A5FB4";

        // The panel each photo sits in, so the picture and the caption under it read as one item
        // rather than as a loose image with some text near it.
        private const string PhotoPanelBorderColor = "#7A7A7A";

        // ###########################################################################################
        // THE ICON FONT: Font Awesome, registered with QuestPDF so the exported document carries the
        // app's own padlocks and category icons rather than words standing in for them.
        //
        // WHY THE BYTES ARE READ SEPARATELY FROM THE REGISTRATION, and why EnsureIconFontLoaded must
        // be called from the UI thread:
        //
        // The .otf is an AvaloniaResource compiled into the assembly (see the csproj) - there is no
        // font file on disk beside the executable, and it is NOT a plain manifest resource either,
        // so the ONLY way to read it is Avalonia's AssetLoader. But AssetLoader resolves
        // `Avalonia.Platform.IAssetLoader` out of Avalonia's locator, which is not available on an
        // arbitrary background thread: the export runs its write inside Task.Run, and doing the load
        // there throws `InvalidOperationException: Unable to locate 'Avalonia.Platform.IAssetLoader'`.
        //
        // That threw INSIDE the try below, so the first version of this simply logged a warning and
        // silently exported every document with no icons at all - which is exactly what it did until
        // the PDF was inspected and found to carry only its text fonts.
        //
        // So: the caller loads the bytes while still on the UI thread (EnsureIconFontLoaded), they
        // are cached in a byte[], and the registration itself - which needs no Avalonia service -
        // happens lazily wherever the export runs.
        //
        // FAILURE IS NOT FATAL: if the asset cannot be read or QuestPDF rejects it, IconFontAvailable
        // stays false and every glyph site draws nothing. An unregistered font renders as a blank box
        // in most readers, which looks like a defect in the document; a missing icon beside a label
        // that already reads "Open" loses nothing.
        // ###########################################################################################
        private const string IconFontName = "CRT Export Icons";

        private static readonly Uri IconFontUri =
            new("avares://Classic-Repair-Toolbox/Assets/Fonts/Font Awesome 7 Free-Solid-900.otf");

        // The font's bytes, read once through AssetLoader on a thread where Avalonia is available.
        private static byte[]? thisIconFontBytes;

        private static bool thisIconFontRegistered;

        private static bool thisIconFontRegistrationFailed;

        private static readonly object IconFontLock = new();

        // ###########################################################################################
        // Reads the icon font into memory, if it has not been read already.
        //
        // MUST be called from the UI thread - see the note above. The export call site does this
        // before it hands the work to Task.Run. Safe to call repeatedly; the second call is a no-op.
        //
        // Deliberately swallows its failure rather than throwing: an export with no icons is still a
        // perfectly good document, and refusing to export because a decoration is missing would be a
        // poor trade.
        // ###########################################################################################
        public static void EnsureIconFontLoaded()
        {
            lock (IconFontLock)
            {
                if (thisIconFontBytes != null || thisIconFontRegistrationFailed)
                    return;

                try
                {
                    using var stream = AssetLoader.Open(IconFontUri);
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    thisIconFontBytes = buffer.ToArray();
                }
                catch (Exception ex)
                {
                    thisIconFontRegistrationFailed = true;
                    Logger.Warning(
                        $"Workbook export: could not read the icon font - icons will be omitted - [{ex.Message}]");
                }
            }
        }

        // Registers the cached bytes with QuestPDF on first use. No Avalonia service is touched here,
        // so this half is safe on any thread.
        private static bool IconFontAvailable
        {
            get
            {
                lock (IconFontLock)
                {
                    if (thisIconFontRegistered)
                        return true;

                    if (thisIconFontRegistrationFailed || thisIconFontBytes == null)
                        return false;

                    try
                    {
                        using var stream = new MemoryStream(thisIconFontBytes);
                        FontManager.RegisterFontWithCustomName(IconFontName, stream);
                        thisIconFontRegistered = true;
                        return true;
                    }
                    catch (Exception ex)
                    {
                        thisIconFontRegistrationFailed = true;
                        Logger.Warning(
                            $"Workbook export: could not register the icon font - icons will be omitted - [{ex.Message}]");
                        return false;
                    }
                }
            }
        }

        // ###########################################################################################
        // Writes the document as a PDF to the given path, overwriting it.
        //
        // Throws on an unwritable path or an unreadable image - the caller shows the failure to the
        // user, because an export that half-worked must not look like it succeeded.
        // ###########################################################################################
        public static void WritePdf(WorkbookExportModel.Document document, string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(1.5f, Unit.Centimetre);
                    page.DefaultTextStyle(x => x.FontSize(10).FontColor(Colors.Black));

                    page.Header().Element(e => ComposeHeader(e, document));
                    page.Content().Element(e => ComposeContent(e, document));

                    // A page number on every page: this document is printed and handed over, and a
                    // dropped sheet from a ten-page repair report is otherwise undetectable.
                    page.Footer().AlignCenter().Text(text =>
                    {
                        text.DefaultTextStyle(x => x.FontSize(8).FontColor(LabelColor));
                        text.CurrentPageNumber();
                        text.Span(" / ");
                        text.TotalPages();
                    });
                });
            }).GeneratePdf(outputPath);
        }

        // ###########################################################################################
        // Writes a ZIP holding the PDF plus every photo and file attached to the workbook, each
        // under a folder named for its entry.
        //
        // WHY THE ORIGINALS AND NOT JUST THE PDF: the PDF embeds photos downscaled to fit a page, and
        // an attached file (a datasheet, a scope capture, an invoice) cannot be embedded at all. A
        // customer who needs the actual bytes - or a second repairer picking the job up - needs the
        // files themselves, and asking the user to gather them by hand from an AppData folder
        // defeats the point of an export button.
        //
        // Entry folders are named "worklog_{id}" - the SAME name the entry's attachments have in
        // the local Workbooks folder (WorklogManager.BuildEntryAttachmentsFolderName), so what the
        // recipient unpacks matches what the repairer sees on their own disk. The worklog title is
        // deliberately NOT in the folder name: it is free text the user typed, often carrying a
        // customer's own details, and the PDF beside it already says which worklog is which.
        //
        // A file whose name collides with one already written gets a numeric suffix rather than
        // overwriting - two entries may legitimately both hold a "front.jpg".
        // ###########################################################################################
        public static void WriteZip(WorkbookExportModel.Document document, string outputPath)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            // The PDF is generated to a temp file and copied in, rather than being written into the
            // archive stream: QuestPDF's GeneratePdf(path) is the API this app uses everywhere else,
            // and a temp file keeps a failed PDF from leaving a half-written entry in the archive.
            string tempPdf = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.pdf");

            try
            {
                WritePdf(document, tempPdf);

                if (File.Exists(outputPath))
                    File.Delete(outputPath);

                using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Create);

                // The names written so far, tracked HERE rather than read back off the archive:
                // ZipArchive in Create mode is write-forward only and throws
                // NotSupportedException ("Cannot access entries in Create mode") from GetEntry and
                // from the Entries collection. An earlier version called GetEntry to test for a
                // collision and crashed the app on the first export of a workbook that had any
                // attachment at all - a workbook with none never reached the call.
                //
                // OrdinalIgnoreCase because the archive may well be unpacked on Windows or macOS,
                // where "Front.jpg" and "front.jpg" collide even though the zip format itself would
                // keep them apart.
                var usedEntryNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                string pdfEntryName = $"{WorkbookExportModel.BuildFileBaseName(document)}.pdf";
                archive.CreateEntryFromFile(tempPdf, pdfEntryName);
                usedEntryNames.Add(pdfEntryName);

                foreach (var section in document.Sections)
                {
                    foreach (var entry in section.Entries)
                    {
                        string folder = BuildEntryFolderName(entry);

                        foreach (var attachment in entry.Photos.Concat(entry.Files))
                        {
                            string entryPath = UniqueEntryName(usedEntryNames, $"{folder}/{attachment.FileName}");

                            // The name is RESERVED BEFORE the write, not recorded after it.
                            // CreateEntryFromFile creates the archive entry and then streams the
                            // source into it, so a file that is locked or truncated part-way
                            // through throws with the entry already in the archive. Recording the
                            // name only on success would then hand the same path to a later
                            // attachment - writing a duplicate entry, which is exactly what this
                            // set exists to prevent, and which most unzip tools resolve by
                            // silently overwriting one with the other.
                            usedEntryNames.Add(entryPath);

                            try
                            {
                                archive.CreateEntryFromFile(attachment.FullPath, entryPath);
                            }
                            catch (Exception ex)
                            {
                                // One unreadable attachment (locked by another process, or removed
                                // between the model being built and this loop) must not lose the
                                // customer the whole archive - the PDF is already in it.
                                Logger.Warning(
                                    $"Workbook export: could not add attachment [{attachment.FullPath}] to the archive - [{ex.Message}]");
                            }
                        }
                    }
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPdf))
                        File.Delete(tempPdf);
                }
                catch (Exception)
                {
                    // A temp file left behind is not worth failing an otherwise-successful export.
                }
            }
        }

        // Through WorklogManager's own helper rather than a second copy of the format string: the
        // archive folder and the on-disk folder are meant to be the same name, and two independent
        // interpolations would be free to drift apart.
        private static string BuildEntryFolderName(WorkbookExportModel.Entry entry) =>
            WorklogManager.BuildEntryAttachmentsFolderName(entry.Record.Id);

        // ###########################################################################################
        // A name not yet used in the archive, suffixed if the desired one is taken.
        //
        // Two entries can legitimately hold identically-named files ("front.jpg" in each), and
        // CreateEntryFromFile would happily write a second archive entry under the same path -
        // which most unzip tools then either overwrite or refuse.
        //
        // Takes the set of names ALREADY WRITTEN rather than the archive itself: ZipArchive in
        // Create mode cannot be queried for its entries at all (see WriteZip). The set is also the
        // honest thing to consult, since it holds exactly what has been committed.
        // ###########################################################################################
        private static string UniqueEntryName(HashSet<string> usedEntryNames, string desired)
        {
            if (!usedEntryNames.Contains(desired))
                return desired;

            string directory = Path.GetDirectoryName(desired)?.Replace('\\', '/') ?? string.Empty;
            string stem = Path.GetFileNameWithoutExtension(desired);
            string extension = Path.GetExtension(desired);

            for (int i = 2; i < 1000; i++)
            {
                string candidate = string.IsNullOrEmpty(directory)
                    ? $"{stem}-{i}{extension}"
                    : $"{directory}/{stem}-{i}{extension}";

                if (!usedEntryNames.Contains(candidate))
                    return candidate;
            }

            // A thousand files of the same name under one entry is not a real workbook; returning
            // the desired name lets the archive finish rather than failing the whole export.
            return desired;
        }

        // ###########################################################################################
        // The document header: what this is, which board, and when it was produced.
        //
        // The generation date is stated explicitly because a worklog keeps changing after an export
        // is sent - without a date, a customer holding a printout has no way to tell it is not the
        // current state of their repair.
        //
        // The workbook's Open/Closed status is drawn as the SAME outlined pill the app shows, not
        // as a word in the run of metadata - see ComposeStatePill.
        // ###########################################################################################
        private static void ComposeHeader(IContainer container, WorkbookExportModel.Document document)
        {
            container.Column(column =>
            {
                column.Item().Row(row =>
                {
                    row.AutoItem().AlignMiddle().Text(document.Title).FontSize(16).Bold().FontColor(HeadingColor);
                    row.AutoItem().PaddingLeft(8).AlignMiddle().Element(e => ComposeStatePill(e, document.Status));
                    row.RelativeItem();
                });

                column.Item().PaddingTop(2).Text(text =>
                {
                    text.DefaultTextStyle(x => x.FontSize(9).FontColor(LabelColor));
                    text.Span($"Workbook #{document.WorkbookId.ToString(CultureInfo.InvariantCulture)}");

                    if (!string.IsNullOrWhiteSpace(document.BoardKey))
                        text.Span($"  \u00b7  {document.BoardKey.Replace("|", " \u00b7 ")}");

                    text.Span($"  \u00b7  started {FormatDate(document.StartDate)}");
                    text.Span($"  \u00b7  exported {FormatDate(document.GeneratedAt)}");
                });

                column.Item().PaddingTop(6).LineHorizontal(1).LineColor(RuleColor);
            });
        }

        private static void ComposeContent(IContainer container, WorkbookExportModel.Document document)
        {
            container.PaddingTop(10).Column(column =>
            {
                column.Spacing(14);

                if (!string.IsNullOrWhiteSpace(document.Note))
                {
                    // The workbook Note is free text that regularly carries a link (the maintainer's
                    // own weblog entry for the repair, in the reported case) - linkified for the
                    // same reason the sub-lists are.
                    column.Item().Element(e => ComposeLinkedText(e, document.Note, 10, HeadingColor));
                }

                column.Item().Element(e => ComposeSummary(e, document.Totals));

                if (document.Sections.Count == 0)
                {
                    column.Item().PaddingTop(10).Text("This workbook has no worklogs.")
                        .FontSize(10).FontColor(LabelColor).Italic();
                    return;
                }

                for (int i = 0; i < document.Sections.Count; i++)
                {
                    // EVERY schematic starts on a fresh page, asked for directly: a section is a
                    // full-width board image plus its worklogs, and one starting halfway down a page
                    // pushed its own image onto the next one, separating the picture from the list
                    // of what is marked on it. The FIRST section does not get a break, or the
                    // document would open on a blank page.
                    if (i > 0)
                        column.Item().PageBreak();

                    column.Item().Element(e => ComposeSection(e, document.Sections[i]));
                }
            });
        }

        // ###########################################################################################
        // The same numbers the app's own summary strip shows, from the same WorkbookSummary - so a
        // customer's document and the repairer's screen cannot disagree about the totals.
        //
        // EVERY NUMBER IS BOLD AND THE WORDS ARE NOT, exactly as on screen. That is why this walks
        // WorkbookSummary's `Stat` PARTS (prefix / number / suffix) rather than calling one of the
        // Format* helpers: a finished string has no way to say which characters were the digits,
        // and re-finding them afterwards would have to guess about values like "0.5 h". The screen
        // has the identical problem and solves it the identical way - see TabWorkbooks.Summary.cs's
        // ApplyStatRuns, which builds Avalonia Runs from these same parts.
        //
        // The category and state counts are the app's own counted pills (count first, no icon), and
        // the STATE pills sit on their OWN LINE under the categories rather than trailing them: two
        // different kinds of pill running together in one row read as one long list of five, with
        // nothing marking where "what kind of work" ends and "how it is progressing" begins.
        // ###########################################################################################
        private static void ComposeSummary(IContainer container, WorkbookSummary.Totals totals)
        {
            container.Background(PanelColor).CornerRadius(3).Padding(10).Column(column =>
            {
                column.Spacing(5);
                column.Item().Text("Summary").FontSize(11).Bold().FontColor(HeadingColor);
                column.Item().Element(e => ComposeStatLine(e, WorkbookSummary.BuildHeadlineStats(totals), 10));

                column.Item().PaddingTop(2).Row(row =>
                {
                    row.Spacing(4);

                    foreach (var (label, count) in WorkbookSummary.BuildCategoryCounts(totals))
                        row.AutoItem().Element(e => ComposeCountChip(e, count, label, CategoryHexColor(label)));

                    row.RelativeItem();
                });

                // The states on their own line - see the note above.
                column.Item().Row(row =>
                {
                    row.Spacing(4);

                    foreach (var (label, count) in WorkbookSummary.BuildStateCounts(totals))
                        row.AutoItem().Element(e => ComposeCountPill(e, count, label, StateHexColor(label)));

                    row.RelativeItem();
                });

                column.Item().PaddingTop(2)
                    .Element(e => ComposeStatLine(e, WorkbookSummary.BuildAttachmentStats(totals), 9));

                if (totals.ComponentCount > 0)
                {
                    column.Item().Element(e => ComposeStatLine(e, WorkbookSummary.BuildComponentStats(totals), 9));
                }
            });
        }

        // ###########################################################################################
        // One line of summary stats with only the NUMBERS in bold.
        //
        // Built span by span from WorkbookSummary's Stat parts, which is the whole reason that type
        // exists - see ComposeSummary. The separator matches the app's own " . " (a middle dot).
        // ###########################################################################################
        private static void ComposeStatLine(IContainer container, IReadOnlyList<WorkbookSummary.Stat> stats, float fontSize)
        {
            container.Text(text =>
            {
                text.DefaultTextStyle(x => x.FontSize(fontSize).FontColor(HeadingColor));

                for (int i = 0; i < stats.Count; i++)
                {
                    if (i > 0)
                        text.Span("  \u00b7  ").FontColor(LabelColor);

                    var stat = stats[i];

                    if (!string.IsNullOrEmpty(stat.Prefix))
                        text.Span(stat.Prefix);

                    text.Span(stat.Number).Bold();

                    if (!string.IsNullOrEmpty(stat.Suffix))
                        text.Span(stat.Suffix);
                }
            });
        }

        // ###########################################################################################
        // One schematic: its name, the board image at FULL PAGE WIDTH with the worklog areas and
        // "#N" pills drawn on it, then each worklog underneath.
        //
        // Full width was asked for directly, and it is what makes the picture usable: a marked area
        // is often a single chip on a board scan, and at the half-page size this used to draw, a
        // reader could see that something was marked but not what.
        // ###########################################################################################
        private static void ComposeSection(IContainer container, WorkbookExportModel.Section section)
        {
            container.Column(column =>
            {
                column.Spacing(8);

                column.Item().Text(section.SchematicName).FontSize(13).Bold().FontColor(HeadingColor);
                column.Item().LineHorizontal(1).LineColor(RuleColor);

                if (section.SchematicImagePath != null)
                {
                    column.Item().Element(e => ComposeSchematicImage(e, section));
                }

                foreach (var entry in section.Entries)
                {
                    column.Item().Element(e => ComposeEntry(e, entry));
                }
            });
        }

        // ###########################################################################################
        // The board image with each worklog's marked area and "#N" pill drawn over it - the same
        // picture the Schematics tab and the Workbooks board pane show, with a 1px outline around
        // the image itself exactly as those two surfaces draw one.
        //
        // HOW THE OVERLAY IS POSITIONED, and the mistake this replaces:
        //
        // An entry's area is stored in the schematic's own PIXEL coordinates, while the page draws
        // the image at whatever width the margins leave - a size QuestPDF decides during layout and
        // never reports. So everything here is PROPORTIONAL: ExportOverlayGeometry turns a pixel
        // rect into fractions of the image, and each band is then expressed as an ASPECT RATIO,
        // which QuestPDF can satisfy against any width it is given. No page dimension appears in
        // this method at all, which is what makes it correct at any margin or paper size.
        //
        // The version this replaces computed the same fractions and passed them to PaddingLeft and
        // PaddingTop multiplied by 100, believing those took a percentage. They do not - every
        // QuestPDF padding is an absolute LENGTH, so "58% across" became "58 points across" and an
        // area covering a tenth of the board was drawn covering most of it. It was reported by
        // holding the PDF next to the screen. ExportOverlayGeometry now owns the maths, and its
        // tests assert the fractions against a real entry's stored coordinates.
        //
        // An entry with ShowMarkedArea OFF gets NO rectangle and its pill is PARKED top-right,
        // mirroring what all three on-screen surfaces do (see ParkedBadgeGeometry) - an entry the
        // user hid the area for must not have it reappear in the document they send out.
        // ###########################################################################################
        private static void ComposeSchematicImage(IContainer container, WorkbookExportModel.Section section)
        {
            var imageSize = TryReadImageSize(section.SchematicImagePath!);
            if (imageSize == null)
            {
                // Unreadable dimensions means no overlay can be placed, but the picture itself may
                // still draw - so it is shown plain rather than dropped.
                container.Border(1).BorderColor(ImageOutlineColor).Image(section.SchematicImagePath!).FitWidth();
                return;
            }

            var (pixelWidth, pixelHeight) = imageSize.Value;

            // Anchored entries carry a drawable area; everything else parks. Built once here so the
            // rectangle pass, the badge pass and the parked stack all agree about which is which.
            var anchored = new List<(WorklogEntryRecord Record, ExportOverlayGeometry.AreaFractions Fractions)>();
            var parked = new List<WorklogEntryRecord>();

            foreach (var entry in section.Entries)
            {
                var record = entry.Record;

                var fractions = record.ShowMarkedArea
                    ? ExportOverlayGeometry.TryBuildAreaFractions(
                        record.AreaX, record.AreaY, record.AreaWidth, record.AreaHeight, pixelWidth, pixelHeight)
                    : null;

                if (fractions.HasValue)
                    anchored.Add((record, fractions.Value));
                else
                    parked.Add(record);
            }

            container
                .Border(1).BorderColor(ImageOutlineColor)
                .AspectRatio(pixelWidth / (float)pixelHeight, AspectRatioOption.FitWidth)
                .Layers(layers =>
                {
                    // The image is the PRIMARY layer: it decides the stack's size, and everything
                    // after it is drawn on top.
                    layers.PrimaryLayer().Image(section.SchematicImagePath!).FitArea();

                    foreach (var (record, fractions) in anchored)
                    {
                        string color = CategoryHexColor(record.Category);

                        layers.Layer().Element(e => ComposeBandPositioned(
                            e, fractions, pixelWidth, pixelHeight,
                            cell => cell
                                .Background(TranslucentAreaColor(color))
                                .Border(AreaBorderThickness).BorderColor(color)));
                    }

                    // Pills after the areas, so one never sits under a wash.
                    foreach (var (record, fractions) in anchored)
                    {
                        layers.Layer().Element(e => ComposeAnchoredBadge(e, record, fractions, pixelWidth, pixelHeight));
                    }

                    if (parked.Count > 0)
                    {
                        layers.Layer().AlignRight().AlignTop().Padding(4)
                            .Element(e => ComposeParkedBadgeBlock(e, parked));
                    }
                });
        }

        // ###########################################################################################
        // The parked "#N" pills, in the image's own top-right corner.
        //
        // WRAPPED INTO A GRID, not stacked in one column. The number of columns comes from
        // ParkedBadgeGeometry.GetGridShape - the same pure helper the three on-screen surfaces use
        // - so a schematic with a dozen hidden-area entries lays them out here the way the
        // Schematics tab and the Workbooks board pane already do.
        //
        // A single unbounded column was the alternative, and it grows past the bottom of the image
        // on a workbook with enough entries: QuestPDF then either clips the overflow or fails the
        // layout outright, which for this library means abandoning the whole document. The grid
        // keeps the block roughly square however many pills there are, which is the property
        // ThumbnailWorklogPillsTests pins on the on-screen side.
        //
        // Only the SHAPE is shared, not the positions: those are pixel offsets for a canvas, and
        // here the rows and columns are laid out by QuestPDF itself.
        // ###########################################################################################
        private static void ComposeParkedBadgeBlock(IContainer container, IReadOnlyList<WorklogEntryRecord> parked)
        {
            var (rowCount, columnCount) = ParkedBadgeGeometry.GetGridShape(parked.Count);

            if (rowCount <= 0 || columnCount <= 0)
                return;

            container.Column(rows =>
            {
                rows.Spacing(3);

                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    int start = rowIndex * columnCount;
                    int count = Math.Min(columnCount, parked.Count - start);

                    if (count <= 0)
                        break;

                    rows.Item().AlignRight().Row(row =>
                    {
                        row.Spacing(3);

                        for (int i = 0; i < count; i++)
                        {
                            var record = parked[start + i];
                            row.AutoItem().Element(e => ComposeEntryBadge(e, record));
                        }
                    });
                }
            });
        }

        // ###########################################################################################
        // Places content over the image at the given fractions, using nothing but aspect ratios.
        //
        // The image is split into vertical bands by a Row (whose RelativeItem weights are exactly
        // the fractions), and the chosen column is then split horizontally by a Column whose bands
        // are sized by AspectRatio - Column has no relative sizing of its own, so the vertical axis
        // has to be expressed as "this band is X wide and Y tall, i.e. this ratio", which is what
        // ExportOverlayGeometry.TryBuildBandAspectRatio computes.
        //
        // Every empty band is omitted rather than emitted at zero: a zero-weight RelativeItem and a
        // zero-ratio AspectRatio are both degenerate, and QuestPDF rejects the whole layout.
        // ###########################################################################################
        private static void ComposeBandPositioned(
            IContainer container,
            ExportOverlayGeometry.AreaFractions fractions,
            int pixelWidth,
            int pixelHeight,
            Action<IContainer> content)
        {
            container.Row(row =>
            {
                if (fractions.Left > 0)
                    row.RelativeItem((float)fractions.Left);

                row.RelativeItem((float)fractions.Width).Column(column =>
                {
                    double? topRatio = ExportOverlayGeometry.TryBuildBandAspectRatio(
                        fractions.Width, fractions.Top, pixelWidth, pixelHeight);

                    if (topRatio.HasValue)
                        column.Item().AspectRatio((float)topRatio.Value, AspectRatioOption.FitWidth);

                    double? areaRatio = ExportOverlayGeometry.TryBuildBandAspectRatio(
                        fractions.Width, fractions.Height, pixelWidth, pixelHeight);

                    if (areaRatio.HasValue)
                    {
                        content(column.Item().AspectRatio((float)areaRatio.Value, AspectRatioOption.FitWidth));
                    }
                });

                if (fractions.RemainingRight > 0)
                    row.RelativeItem((float)fractions.RemainingRight);
            });
        }

        // ###########################################################################################
        // An anchored "#N" pill, hung off its area's top-left corner.
        //
        // On screen the badge STRADDLES that corner - half of it either side (see BadgeGeometry's
        // centre offset, which WorklogBadgeLayout applies for both the Schematics tab and the
        // Workbooks board pane). Expressing that here would need a negative offset, and a negative
        // offset pushes content outside the layer bounds, which QuestPDF rejects by failing the
        // ENTIRE document rather than clipping. So the badge's own box is aligned to the corner and
        // overhangs down-and-right instead: the same corner, the same size, half a pill's offset
        // from the screen's placement, which is far closer than the version that drew it inside the
        // area entirely.
        //
        // The pill is a FIXED size in points on purpose - it marks a point on the board, so it must
        // not grow with the page the way the image does.
        // ###########################################################################################
        private static void ComposeAnchoredBadge(
            IContainer container,
            WorklogEntryRecord record,
            ExportOverlayGeometry.AreaFractions fractions,
            int pixelWidth,
            int pixelHeight)
        {
            container.Column(outer =>
            {
                // The vertical spacer down to the area's top edge, sized by aspect ratio against
                // the FULL image width - the same technique the rectangle uses, just measured over
                // the whole width rather than a band of it.
                double? topRatio = ExportOverlayGeometry.TryBuildBandAspectRatio(
                    1.0, fractions.Top, pixelWidth, pixelHeight);

                if (topRatio.HasValue)
                    outer.Item().AspectRatio((float)topRatio.Value, AspectRatioOption.FitWidth);

                // The badge's own row, which takes only the height it needs. It must NOT be given
                // an aspect-ratio band of its own: the remaining bands already account for the
                // whole image, so a further sized item leaves this one zero height and QuestPDF
                // then refuses to render the text inside the pill - failing the entire document
                // rather than clipping. Letting the row size itself is what keeps the badge
                // drawable while still starting at the right vertical offset.
                outer.Item().Row(row =>
                {
                    if (fractions.Left > 0)
                        row.RelativeItem((float)fractions.Left);

                    // What remains to the right of the corner. Never zero: an area flush against
                    // the right edge would otherwise ask for a zero-weight band, which QuestPDF
                    // treats as degenerate.
                    row.RelativeItem((float)Math.Max(0.0001, 1.0 - fractions.Left))
                       .AlignLeft().AlignTop()
                       .Element(e => ComposeEntryBadge(e, record));
                });
            });
        }

        // ###########################################################################################
        // The "#N" pill that marks a worklog on the board - the filled, category-coloured badge
        // with a WHITE DISC holding the state padlock, which is exactly what WorklogBadgeBuilder
        // draws on screen.
        //
        // The disc stays white rather than taking the state colour for the same reason it does on
        // screen: the badge behind it is already filled with the category colour, and a
        // state-coloured disc on that puts two saturated colours together with the glyph lost
        // between them.
        //
        // FULLY ROUNDED, like every other pill in this document - the first version drew all of
        // these as square boxes, which was reported. QuestPDF does have CornerRadius; it simply
        // was not used.
        //
        // Small and fixed-size on purpose: it marks a POINT on the board, so it must not grow with
        // the page the way the image does, or a full-width A4 schematic would carry pills the size
        // of a postage stamp each.
        // ###########################################################################################
        private static void ComposeEntryBadge(IContainer container, WorklogEntryRecord record)
        {
            string color = CategoryHexColor(record.Category);

            container
                .Background(color)
                .CornerRadius(BadgeCornerRadius)
                .PaddingVertical(2).PaddingHorizontal(5)
                .Row(row =>
                {
                    row.Spacing(3);

                    // NO AlignMiddle on either item, deliberately. A vertical alignment inside a
                    // row whose own height is still being negotiated measures its child against
                    // ZERO height, and QuestPDF then reports "not sufficient to render even a
                    // single line of text" and fails the WHOLE document rather than clipping one
                    // pill. The row already centres these two against each other by baseline, so
                    // the alignment bought nothing and cost the export.
                    row.AutoItem()
                        .Text($"#{record.Id.ToString(CultureInfo.InvariantCulture)}")
                        .FontSize(7).Bold().FontColor(Colors.White);

                    // The white disc. A fixed Width/Height with a radius of half of it is what
                    // makes it a circle rather than a rounded square.
                    //
                    // The padlock inside takes the STATE colour, not the badge's category colour -
                    // green for Closed, red for Open - exactly as WorklogBadgeBuilder paints it on
                    // screen (categoryColor for the fill, stateColor for the glyph). The two are
                    // deliberately different channels: the badge says what KIND of work this is,
                    // the padlock says whether it is DONE. Colouring the glyph to match its own
                    // background made the badge report the category twice and the state not at
                    // all, which was reported - a green padlock on a red Issue badge is the whole
                    // point of having both.
                    //
                    // Drawn ONLY when the icon font is available: with no glyph to hold, the disc
                    // is a meaningless white dot, and a fixed-size box around an EMPTY TextBlock
                    // is a layout QuestPDF refuses outright ("not sufficient to render even a
                    // single line of text") - failing the whole export rather than the one pill.
                    // That is reachable in normal use, not only in tests: EnsureIconFontLoaded
                    // degrades to no icons whenever the asset cannot be read.
                    if (IconFontAvailable)
                    {
                        row.AutoItem()
                            .Width(StateDiscSize).Height(StateDiscSize)
                            .Background(Colors.White)
                            .CornerRadius(StateDiscSize / 2f)
                            .AlignCenter().AlignMiddle()
                            .Element(e => ComposeStateGlyph(
                                e, record.State, StateDiscGlyphSize, StateHexColor(record.State)));
                    }
                });
        }

        // ###########################################################################################
        // An outlined Open/Closed pill - the app's own informational visual (see
        // WorklogInfoPillBuilder): a 1px border in the state's colour, a Form_Bg fill, the padlock
        // and label in that same colour, and FULLY ROUNDED corners.
        //
        // The rounding is the point: on screen a status pill has a 10px radius while a category
        // chip has 3px, and that difference is how the two are told apart at a glance when they
        // sit side by side. The first exported version drew both as plain rectangles.
        // ###########################################################################################
        private static void ComposeStatePill(IContainer container, string state)
        {
            string color = StateHexColor(state);

            container
                .Background(PillFillColor)
                .Border(1).BorderColor(color)
                .CornerRadius(PillCornerRadius)
                .PaddingVertical(2).PaddingHorizontal(6)
                .Row(row =>
                {
                    row.Spacing(3);
                    row.AutoItem().AlignMiddle().Element(e => ComposeStateGlyph(e, state, 7f, color));
                    row.AutoItem().AlignMiddle()
                        .Text(PillLabel(state, "Open")).FontSize(8).SemiBold().FontColor(color);
                });
        }

        // ###########################################################################################
        // A counted STATE pill for the summary - "2 Open" - outlined, fully rounded, and with NO
        // icon, exactly as the app's summary strip draws it (a padlock between a number and its
        // label reads as a third piece of information rather than as decoration).
        //
        // The count leads and is BOLD, matching WorklogInfoPillBuilder.BuildCountLabel: across the
        // whole summary the numbers are the content and the words merely label them.
        // ###########################################################################################
        private static void ComposeCountPill(IContainer container, int count, string label, string color) =>
            ComposeCountShape(container, count, label, color, PillCornerRadius);

        // The same, for a CATEGORY - "3 Note" - at the chip's smaller corner radius, which is the
        // one visual difference between the two shapes on screen.
        private static void ComposeCountChip(IContainer container, int count, string label, string color) =>
            ComposeCountShape(container, count, label, color, ChipCornerRadius);

        private static void ComposeCountShape(
            IContainer container, int count, string label, string color, float cornerRadius)
        {
            container
                .Background(PillFillColor)
                .Border(1).BorderColor(color)
                .CornerRadius(cornerRadius)
                .PaddingVertical(2).PaddingHorizontal(6)
                .Row(row =>
                {
                    row.Spacing(3);
                    row.AutoItem().AlignMiddle()
                        .Text(count.ToString(CultureInfo.InvariantCulture)).FontSize(8).Bold().FontColor(color);
                    row.AutoItem().AlignMiddle()
                        .Text(PillLabel(label, "-")).FontSize(8).FontColor(color);
                });
        }

        // ###########################################################################################
        // The Font Awesome padlock for a state, in the registered icon font - the SAME codepoints
        // every on-screen worklog surface uses (WorklogGlyphs), so the exported document carries the
        // app's own icons rather than a written-out word.
        //
        // Falls back to nothing at all if the icon font could not be registered: a missing glyph
        // renders as a blank box in most readers, which looks like a defect. The label beside every
        // one of these already says Open or Closed, so dropping the icon loses no information.
        // ###########################################################################################
        private static void ComposeStateGlyph(IContainer container, string state, float size, string color)
        {
            if (!IconFontAvailable)
            {
                // NOT container.Text(string.Empty): an empty TextBlock still demands a line box,
                // and inside a fixed-size container QuestPDF reports it as unrenderable and fails
                // the entire document. Collapsing the element to nothing is what makes "no icon
                // font" a cosmetic degradation rather than a broken export.
                container.Height(0);
                return;
            }

            string glyph = WorklogGlyphs.GlyphFor(WorklogManager.IsResolvedState(state));
            container.Text(glyph).FontFamily(IconFontName).FontSize(size).FontColor(color);
        }

        // ###########################################################################################
        // One worklog entry: its heading, description, the components it covers, and its work
        // done / comments / links / photos.
        //
        // ShowEntire keeps an entry from being split across a page break where it will fit whole -
        // a repair reads as one item, and a heading stranded at the foot of a page reads as a
        // missing page to someone holding the printout.
        //
        // The heading carries the same "#N" pill the board image does, plus the category chip and
        // state pill, so an entry in the list is recognisably the same object as its mark on the
        // picture above.
        // ###########################################################################################
        private static void ComposeEntry(IContainer container, WorkbookExportModel.Entry entry)
        {
            var record = entry.Record;

            container.ShowEntire().PaddingLeft(4).BorderLeft(2).BorderColor(CategoryHexColor(record.Category))
                .PaddingLeft(8).PaddingBottom(4).Column(column =>
            {
                column.Spacing(4);

                // The TITLE takes the RelativeItem and the badge the AutoItem, not the other way
                // round. A long title in an AutoItem measures at its full unwrapped width, claims
                // the whole row, and leaves the badge beside it with ZERO space - which QuestPDF
                // reports as a conflicting size constraint and refuses to lay out at all, failing
                // the entire export. It surfaced the moment the badge gained a fixed-size white
                // disc that cannot shrink to fit; the previous all-text badge simply squashed.
                // A RelativeItem title wraps instead, which is what a sentence-length worklog
                // title should do anyway.
                column.Item().Row(row =>
                {
                    row.Spacing(6);
                    row.AutoItem().AlignTop().Element(e => ComposeEntryBadge(e, record));
                    row.RelativeItem().AlignMiddle()
                        .Text(string.IsNullOrWhiteSpace(record.Title) ? "(untitled)" : record.Title)
                        .FontSize(11).Bold().FontColor(HeadingColor);
                });

                column.Item().Row(row =>
                {
                    row.Spacing(4);
                    row.AutoItem().Element(e => ComposeCategoryChip(e, record.Category));
                    row.AutoItem().Element(e => ComposeStatePill(e, record.State));
                    row.RelativeItem();
                });

                if (!string.IsNullOrWhiteSpace(record.Description))
                {
                    column.Item().PaddingTop(2)
                        .Element(e => ComposeLinkedText(e, record.Description, 10, HeadingColor));
                }

                // Components in scope, with the completed ones marked - this is the part a customer
                // reads as "what was actually touched on my board".
                if (record.ComponentLabels.Count > 0)
                {
                    string components = string.Join(", ", record.ComponentLabels.Select(label =>
                        record.CompletedComponentLabels.Contains(label, StringComparer.OrdinalIgnoreCase)
                            ? $"{label} (done)"
                            : label));

                    column.Item().Element(e => ComposeLabelled(e, "Components", components));
                }

                if (record.WorkDoneItems.Count > 0)
                {
                    column.Item().Element(e => ComposeSubList(e, "Work done", record.WorkDoneItems.Select(w =>
                    {
                        string totals = $"{w.HoursSpent.ToString("0.##", CultureInfo.InvariantCulture)} h \u00b7 {w.Cost.ToString("0.##", CultureInfo.InvariantCulture)}";
                        return $"{FormatDate(w.Date)} - {w.Text}  ({totals})";
                    })));
                }

                if (record.Comments.Count > 0)
                {
                    column.Item().Element(e => ComposeSubList(e, "Comments",
                        record.Comments.Select(c => $"{FormatDate(c.Date)} - {c.Text}")));
                }

                if (record.Links.Count > 0)
                {
                    column.Item().Element(e => ComposeLinkRows(e, record.Links));
                }

                if (entry.Files.Count > 0)
                {
                    // Named but not embedded: a PDF cannot show a datasheet or a scope capture. The
                    // ZIP export carries the actual bytes, which is what the list here refers to.
                    column.Item().Element(e => ComposeSubList(e, "Files", entry.Files.Select(f =>
                        string.IsNullOrWhiteSpace(f.Comment) ? f.FileName : $"{f.FileName} - {f.Comment}")));
                }

                if (entry.Photos.Count > 0)
                {
                    column.Item().Element(e => ComposePhotos(e, entry));
                }
            });
        }

        // ###########################################################################################
        // The category chip for ONE entry, in the outlined informational visual WITH its icon -
        // unlike the summary strip's counted chips, this one names what a single entry IS, which is
        // exactly the case the icon belongs on (see WorklogInfoPillBuilder, which branches the same
        // way on whether a count is present).
        //
        // Softened corners rather than the status pill's full rounding, matching the on-screen
        // chip's smaller radius.
        // ###########################################################################################
        private static void ComposeCategoryChip(IContainer container, string category)
        {
            string color = CategoryHexColor(category);

            container
                .Background(PillFillColor)
                .Border(1).BorderColor(color)
                .CornerRadius(ChipCornerRadius)
                .PaddingVertical(2).PaddingHorizontal(6)
                .Row(row =>
                {
                    row.Spacing(3);

                    if (IconFontAvailable && CategoryGlyphs.TryGetValue(category, out int codepoint))
                    {
                        row.AutoItem().AlignMiddle()
                            .Text(char.ConvertFromUtf32(codepoint))
                            .FontFamily(IconFontName).FontSize(7).FontColor(color);
                    }

                    row.AutoItem().AlignMiddle()
                        .Text(PillLabel(category, "Note")).FontSize(8).FontColor(color);
                });
        }

        // ###########################################################################################
        // The entry's photos, each in its OWN PANEL - a grey-bordered box holding the picture, the
        // file name, and the comment.
        //
        // The panel was asked for directly, and the reason is that a bare image followed by a line
        // of text gives a reader no way to tell which caption belongs to which picture once two sit
        // side by side - the gap between a photo and ITS text is the same as the gap to the next
        // photo's. A border makes the grouping structural rather than something the reader has to
        // infer from spacing.
        //
        // The FILE NAME is printed as well as the comment, so a recipient reading the PDF can find
        // that exact photo in the ZIP export's "worklog_{id}" folder (or the repairer in their own
        // Workbooks folder). It is shown even when the photo has no comment, which is the common
        // case and the one where the name is the only handle on the file.
        //
        // Two to a row: one per row wastes half a page on a phone photo, and more than two makes
        // the detail in a board close-up unreadable.
        // ###########################################################################################
        private static void ComposePhotos(IContainer container, WorkbookExportModel.Entry entry)
        {
            container.PaddingTop(4).Column(column =>
            {
                column.Spacing(6);
                column.Item().Text("Photos").FontSize(8).Bold().FontColor(LabelColor);

                foreach (var chunk in entry.Photos.Chunk(2))
                {
                    column.Item().Row(row =>
                    {
                        row.Spacing(8);

                        foreach (var photo in chunk)
                        {
                            row.RelativeItem().Element(e => ComposePhotoPanel(e, photo));
                        }

                        // A trailing empty cell so a lone photo on the last row keeps its half-width
                        // rather than stretching across the page at a different scale to the others.
                        if (chunk.Length == 1)
                            row.RelativeItem();
                    });
                }
            });
        }

        // One photo's panel: the picture, its file name, and its comment if it has one.
        private static void ComposePhotoPanel(IContainer container, WorkbookExportModel.Attachment photo)
        {
            container
                .Border(1).BorderColor(PhotoPanelBorderColor)
                .CornerRadius(2)
                .Padding(5)
                .Column(cell =>
                {
                    cell.Spacing(3);

                    try
                    {
                        cell.Item().MaxHeight(6, Unit.Centimetre).Image(photo.FullPath).FitArea();
                    }
                    catch (Exception ex)
                    {
                        // A file that exists but is not a decodable image - a renamed non-image, or
                        // a truncated copy. Named rather than dropped, so the reader knows a photo
                        // was meant to be here.
                        Logger.Warning(
                            $"Workbook export: could not embed photo [{photo.FullPath}] - [{ex.Message}]");
                        cell.Item().Text("[this photo could not be displayed]")
                            .FontSize(8).Italic().FontColor(LabelColor);
                    }

                    // The file name, so the picture can be found on disk later. Not linkified: it
                    // is a name, not a destination, and the file lives in the ZIP beside the PDF.
                    cell.Item().Text(photo.FileName).FontSize(7).FontColor(LabelColor);

                    if (!string.IsNullOrWhiteSpace(photo.Comment))
                    {
                        cell.Item().Element(e => ComposeLinkedText(e, photo.Comment, 8, HeadingColor));
                    }
                });
        }

        private static void ComposeLabelled(IContainer container, string label, string value)
        {
            container.Text(text =>
            {
                text.Span($"{label}: ").FontSize(9).Bold().FontColor(LabelColor);
                text.Span(value).FontSize(9);
            });
        }

        private static void ComposeSubList(IContainer container, string heading, IEnumerable<string> lines)
        {
            container.PaddingTop(2).Column(column =>
            {
                column.Spacing(1);
                column.Item().Text(heading).FontSize(8).Bold().FontColor(LabelColor);

                foreach (var line in lines)
                {
                    // Linkified: these rows are the work-done / comment / link / file lines, all of
                    // which are free text the user typed and routinely carry a URL. The app makes
                    // those clickable on screen (TextLinkRenderer), so the exported document does
                    // too, through the SAME TextLinkFinder rules.
                    column.Item().Element(e => ComposeLinkedText(e, $"\u2022 {line}", 9, HeadingColor));
                }
            });
        }

        // ###########################################################################################
        // The entry's LINK rows - the ones the user added through "Add link", each a headline plus
        // a URL.
        //
        // These get their own composer rather than going through ComposeSubList because the URL
        // here is a DECLARED destination, not a URL spotted inside prose. TextLinkFinder is
        // deliberately conservative and rejects a bare "example.com" - correct when scanning repair
        // notes full of part numbers, wrong for a field whose entire purpose is to hold a link. The
        // add-link dialog stores whatever the user typed without normalising a scheme onto it, so
        // that shape genuinely occurs.
        //
        // The whole "headline - url" run is therefore made one hyperlink, with the scheme filled in
        // for the click target when the stored text lacks one. The visible text is left exactly as
        // the user typed it.
        // ###########################################################################################
        private static void ComposeLinkRows(IContainer container, IReadOnlyList<WorklogLinkRecord> links)
        {
            container.PaddingTop(2).Column(column =>
            {
                column.Spacing(1);
                column.Item().Text("Links").FontSize(8).Bold().FontColor(LabelColor);

                foreach (var link in links)
                {
                    string url = link.Url ?? string.Empty;
                    string label = string.IsNullOrWhiteSpace(link.Headline) ? url : $"{link.Headline} - {url}";
                    string? target = BuildLinkTarget(url);

                    if (target == null)
                    {
                        // No usable URL at all - shown as plain text rather than as a link that
                        // goes nowhere.
                        column.Item().Text($"\u2022 {label}").FontSize(9).FontColor(HeadingColor);
                        continue;
                    }

                    column.Item().Text(composed =>
                    {
                        composed.DefaultTextStyle(x => x.FontSize(9).FontColor(HeadingColor));
                        composed.Span("\u2022 ");
                        composed.Hyperlink(label, target).FontColor(LinkColor).Underline();
                    });
                }
            });
        }

        // An absolute http/https target for a stored link URL, or null when there is nothing usable.
        // A stored value with no scheme gets https, matching what TextLinkFinder does for a "www."
        // run it finds in prose - a PDF hyperlink with no scheme is simply ignored by readers.
        private static string? BuildLinkTarget(string url)
        {
            string trimmed = (url ?? string.Empty).Trim();

            if (trimmed.Length == 0)
                return null;

            if (trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return trimmed;

            // Anything else that could not be a host is not worth linking.
            return trimmed.Contains(' ') ? null : $"https://{trimmed}";
        }

        // ###########################################################################################
        // Free text with any web links in it drawn as REAL, VISIBLE hyperlinks - blue, underlined,
        // and clickable.
        //
        // WHICH RUNS ARE LINKS is decided by TextLinkFinder, the same pure helper the on-screen
        // renderer uses, rather than by a second rule invented here. That matters: the finder is
        // deliberately conservative because repair prose is full of things shaped like domains
        // ("74LS08.pin3", "5.0V", "notes.txt"), and a document that linkified those where the app
        // does not would disagree with the screen it was exported from.
        //
        // BLUE AND UNDERLINED IS NOT DECORATION. QuestPDF's Hyperlink makes a run clickable but
        // styles it exactly like the text around it, so a link was reachable only by a reader who
        // happened to click the right words - reported as "it is clickable but does not show as a
        // link". A PDF viewer offers no hover cue of its own, and on a printed page the styling is
        // the ONLY thing that survives.
        // ###########################################################################################
        private static void ComposeLinkedText(IContainer container, string text, float fontSize, string color)
        {
            var spans = TextLinkFinder.FindSpans(text);

            // The overwhelmingly common case - no link at all - stays one plain span rather than
            // paying for a composed text block.
            if (spans.Count == 1 && !spans[0].IsLink)
            {
                container.Text(text).FontSize(fontSize).FontColor(color);
                return;
            }

            container.Text(composed =>
            {
                composed.DefaultTextStyle(x => x.FontSize(fontSize).FontColor(color));

                foreach (var span in spans)
                {
                    string run = text.Substring(span.Start, span.Length);

                    if (span.IsLink)
                        composed.Hyperlink(run, span.Url!).FontColor(LinkColor).Underline();
                    else
                        composed.Span(run);
                }
            });
        }

        // ###########################################################################################
        // The image's pixel dimensions, read from the FILE HEADER rather than by decoding it.
        //
        // The overlay maths needs the size in the same pixel coordinates an entry's area is stored
        // in. Decoding a 4220x2941 board scan just to read two numbers would cost ~47 MB per
        // schematic; the header of a PNG or JPEG carries them in its first few dozen bytes.
        //
        // Null when the file is not one of those formats or its header is malformed - the caller
        // then draws the picture with no overlay rather than failing the export.
        // ###########################################################################################
        private static (int Width, int Height)? TryReadImageSize(string path)
        {
            try
            {
                using var stream = File.OpenRead(path);
                using var reader = new BinaryReader(stream);

                var signature = reader.ReadBytes(8);

                // PNG: an 8-byte signature, then an IHDR chunk whose width and height are big-endian
                // 32-bit integers at a fixed offset.
                if (signature.Length == 8 &&
                    signature[0] == 0x89 && signature[1] == 0x50 && signature[2] == 0x4E && signature[3] == 0x47)
                {
                    stream.Position = 16;
                    int width = ReadBigEndianInt32(reader);
                    int height = ReadBigEndianInt32(reader);
                    return width > 0 && height > 0 ? (width, height) : null;
                }

                // JPEG: a chain of markers; the SOFn frame header carries the dimensions. Walked
                // rather than assumed at a fixed offset, since the number and size of the preceding
                // segments (EXIF, colour profiles) vary per file.
                //
                // WHICH MARKERS ARE FRAME HEADERS is the whole subtlety here, and getting it wrong
                // is silent: the walker stops at the wrong marker, reads four arbitrary payload
                // bytes as the dimensions, and every marked area on that schematic is then placed
                // against a bogus image size - no exception, nothing logged, an overlay simply
                // nowhere near the copper it marks.
                //
                // Of 0xC0..0xCF, FOUR are not frame headers and must be skipped as ordinary
                // segments: C4 (DHT), C8 (JPG), CC (DAC) - and the range does NOT extend past CF
                // into frame territory either, because CD/CE/CF are DNL/DHP/EXP, which carry no
                // dimensions at the SOF offset. So the frame set is C0-C3, C5-C7, C9-CB only.
                if (signature.Length >= 2 && signature[0] == 0xFF && signature[1] == 0xD8)
                {
                    stream.Position = 2;

                    while (stream.Position < stream.Length - 1)
                    {
                        if (reader.ReadByte() != 0xFF)
                            continue;

                        byte marker = reader.ReadByte();

                        // A run of 0xFF bytes is legal padding before a marker - skip to the last.
                        while (marker == 0xFF && stream.Position < stream.Length)
                            marker = reader.ReadByte();

                        if (IsJpegFrameHeaderMarker(marker))
                        {
                            reader.ReadBytes(3); // segment length (2) + sample precision (1)
                            int height = (reader.ReadByte() << 8) | reader.ReadByte();
                            int width = (reader.ReadByte() << 8) | reader.ReadByte();
                            return width > 0 && height > 0 ? (width, height) : null;
                        }

                        // STANDALONE markers carry no length word at all, so reading two bytes
                        // after one of these would consume image data as a segment length and seek
                        // to an arbitrary offset. 0x01 is TEM and 0xD0-0xD7 are the restart
                        // markers; 0xD8 (SOI) and 0xD9 (EOI) are standalone too.
                        if (marker == 0x00 || marker == 0x01 || (marker >= 0xD0 && marker <= 0xD9))
                            continue;

                        // Start of scan: entropy-coded image data follows, which is not a segment
                        // chain and must not be walked. A SOF always precedes it in a valid file,
                        // so reaching here means there is nothing more to find.
                        if (marker == 0xDA)
                            return null;

                        int length = (reader.ReadByte() << 8) | reader.ReadByte();
                        if (length < 2)
                            return null;

                        stream.Position += length - 2;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Workbook export: could not read image dimensions for [{path}] - [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // Whether a JPEG marker is a FRAME HEADER (SOFn), i.e. one whose payload begins with the
        // sample precision then the image height and width.
        //
        // C0-C3, C5-C7 and C9-CB are frame headers. The gaps are the ones that merely LOOK like
        // part of the run: C4 is DHT (Huffman tables), C8 is JPG (reserved), CC is DAC (arithmetic
        // coding conditioning) - and CD/CE/CF are DNL/DHP/EXP, none of which carry dimensions here.
        // ###########################################################################################
        private static bool IsJpegFrameHeaderMarker(byte marker) =>
            (marker >= 0xC0 && marker <= 0xC3) ||
            (marker >= 0xC5 && marker <= 0xC7) ||
            (marker >= 0xC9 && marker <= 0xCB);

        private static int ReadBigEndianInt32(BinaryReader reader)
        {
            var bytes = reader.ReadBytes(4);
            return bytes.Length < 4 ? 0 : (bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3];
        }

        // ###########################################################################################
        // A pill's visible label, never blank.
        //
        // An empty string inside a bordered, padded box is the one input QuestPDF answers by
        // failing the WHOLE document ("not sufficient to render even a single line of text") rather
        // than by drawing nothing - the same trap the icon-font fallback and the badge's white disc
        // are both guarded against elsewhere in this file.
        //
        // It is reachable without a bug on this side: State and Category are plain strings in
        // entries.json, and a hand-edited file or a record written by an older build can carry an
        // empty one. The fallback matches what the rest of the app assumes of a blank value -
        // WorklogManager.IsResolvedState treats anything not "Closed" as open, and the category
        // resolvers default to Note.
        // ###########################################################################################
        private static string PillLabel(string value, string fallback) =>
            string.IsNullOrWhiteSpace(value) ? fallback : value;

        // ###########################################################################################
        // The category's colour as a hex string.
        //
        // These are App.axaml's Worklog_Category_* values written out, NOT resolved from the theme
        // at export time, and that is deliberate twice over. First, this class must not depend on
        // Application.Current - the export is ordinary logic and should not need a running UI.
        // Second, an exported document is a fixed artefact: it goes to a customer and gets printed,
        // and its colours should not depend on which theme the repairer had on when they pressed
        // the button.
        //
        // Both of this app's themes currently define these identically (see App.axaml), so there is
        // nothing to choose between them. If a theme ever diverges, THIS is the light theme's set,
        // which is the one that prints.
        // ###########################################################################################
        // CASE-INSENSITIVE, like every other category comparison in this app (WorkbookSummary's
        // own tally, WorklogInfoPillBuilder.CategoryIconsByName, WorklogSearchIndex). A plain
        // switch here was case-SENSITIVE while the CategoryGlyphs dictionary beside it was not, so
        // an entry stored as "note" drew the right icon in the wrong colour - and its badge, its
        // left rule and its area wash on the board image all went grey while the summary chip for
        // the same category stayed blue.
        private static string CategoryHexColor(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
                return NoteCategoryColor;

            if (string.Equals(category, "Note", StringComparison.OrdinalIgnoreCase))
                return NoteCategoryColor;

            if (string.Equals(category, "Cosmetic", StringComparison.OrdinalIgnoreCase))
                return "#C8880E";

            // App.axaml spells this one as the named colour "IndianRed"; QuestPDF wants hex.
            if (string.Equals(category, "Issue", StringComparison.OrdinalIgnoreCase))
                return "#CD5C5C";

            // An unrecognised category - a value from a future build, or a hand-edited entry - gets
            // a neutral rule rather than being coloured as one of the three it is not.
            return "#7A7A7A";
        }

        // Also the colour a blank category takes, matching PillLabel's own "Note" fallback so the
        // chip's text and its colour cannot disagree about what an empty value means.
        private const string NoteCategoryColor = "#2F6FB5";

        // Worklog_Status_Closed / _Open, same reasoning as the categories above. Anything that is
        // not a resolved state reads as open, matching ResolveWorklogStateColor everywhere else.
        private static string StateHexColor(string state) =>
            WorklogManager.IsResolvedState(state) ? "#4C8C31" : "#CD5C5C";

        // The area wash drawn under a marked rectangle: the category colour at low alpha, which is
        // what WorklogEntriesOverlay paints on screen (a SolidColorBrush at FillOpacity). QuestPDF
        // takes colours as #AARRGGBB, so the alpha is prefixed onto the same hex.
        private static string TranslucentAreaColor(string hexColor) =>
            "#33" + hexColor.TrimStart('#');

        // fa-regular note-sticky / fa-solid paint-roller / fa-solid triangle-exclamation - the same
        // codepoints WorklogInfoPillBuilder uses, so an exported chip carries the app's own icon.
        private static readonly Dictionary<string, int> CategoryGlyphs = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Note"] = 0xF15C,
            ["Cosmetic"] = 0xF5D0,
            ["Issue"] = 0xF188
        };

        private static string FormatDate(DateTime value) =>
            value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }
}
