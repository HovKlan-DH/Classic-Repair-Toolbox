using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// What actually goes into an exported workbook document, in what order, and what an absent field
// turns into.
//
// This is the half of the export worth pinning down: WorkbookPdfExporter only paints what this
// model decides, and asserting on the PDF bytes it produces would test QuestPDF rather than this
// app. The failure modes that matter are all here - a missing photo file, an entry with no
// schematic, a board Excel file naming the same schematic twice - and every one of them is a state
// real user data reaches.
//
// Uses TempWorkspace for the attachment files, since ResolveAttachments genuinely checks the disk.
public class WorkbookExportModelTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private static WorkbookRecord Workbook(string title = "Repair job", string note = "") => new()
    {
        Id = 3,
        BoardKey = "C64/250469",
        Title = title,
        Note = note,
        Status = "Closed",
        StartDate = new DateTime(2026, 9, 3)
    };

    private static WorklogEntryRecord Entry(int id, string schematic, string title = "Bad cap") => new()
    {
        Id = id,
        SchematicName = schematic,
        Title = title,
        Category = "Issue",
        State = "Open"
    };

    private static WorkbookExportModel.Document Build(
        WorkbookRecord workbook,
        IEnumerable<WorklogEntryRecord> entries,
        IReadOnlyDictionary<string, string>? images = null,
        Func<int, string?>? attachments = null) =>
        WorkbookExportModel.Build(
            workbook,
            entries,
            images,
            attachments ?? (_ => null),
            new DateTime(2026, 9, 4),
            "DKK");

    // ###########################################################################################
    // Sections are per schematic and alphabetical; entries within one are in ID order - the number
    // the "#N" pills on screen show, so a reader with the app open follows the same sequence.
    // ###########################################################################################
    [Fact]
    public void Entries_are_grouped_into_alphabetical_sections_and_ordered_by_id_within_each()
    {
        var document = Build(Workbook(), new[]
        {
            Entry(3, "PCB Top"),
            Entry(1, "Sheet 1"),
            Entry(2, "PCB Top")
        });

        Assert.Equal(new[] { "PCB Top", "Sheet 1" }, document.Sections.Select(s => s.SchematicName));
        Assert.Equal(new[] { 2, 3 }, document.Sections[0].Entries.Select(e => e.Record.Id));
        Assert.Equal(new[] { 1 }, document.Sections[1].Entries.Select(e => e.Record.Id));
    }

    // ###########################################################################################
    // An entry with no schematic name is the one thing this export must never silently drop: it is
    // the user's own worklog, and its absence from a customer document would go unnoticed. It gets
    // a parenthesised heading instead, matching the "(untitled)" placeholder convention used
    // elsewhere in the worklog UI so it reads as an absence rather than a real schematic name.
    // ###########################################################################################
    [Fact]
    public void An_entry_with_no_schematic_is_still_exported_under_its_own_heading()
    {
        var document = Build(Workbook(), new[] { Entry(1, ""), Entry(2, "Sheet 1") });

        var unassigned = document.Sections.Single(s => s.SchematicName == WorkbookExportModel.UnassignedSectionName);
        Assert.Equal(1, unassigned.Entries.Single().Record.Id);
    }

    [Fact]
    public void The_document_carries_the_workbooks_own_identity_and_the_generation_date()
    {
        var document = Build(Workbook(note: "Found at the tip"), Array.Empty<WorklogEntryRecord>());

        Assert.Equal(3, document.WorkbookId);
        Assert.Equal("Repair job", document.Title);
        Assert.Equal("C64/250469", document.BoardKey);
        Assert.Equal("Found at the tip", document.Note);
        Assert.Equal("Closed", document.Status);
        Assert.Equal(new DateTime(2026, 9, 4), document.GeneratedAt);
    }

    // A blank title gets the same "(untitled)" placeholder the tab shows, rather than a blank
    // heading on the customer's document.
    [Fact]
    public void A_workbook_with_no_title_exports_as_untitled()
    {
        var document = Build(Workbook(title: "   "), Array.Empty<WorklogEntryRecord>());

        Assert.Equal("(untitled)", document.Title);
    }

    // The totals come from the shared WorkbookSummary, so the PDF's summary section and the tab's
    // summary strip are provably the same numbers.
    [Fact]
    public void The_document_carries_the_same_totals_the_summary_strip_shows()
    {
        var entries = new[] { Entry(1, "Sheet 1"), Entry(2, "Sheet 1") };

        var document = Build(Workbook(), entries);

        Assert.Equal(2, document.Totals.WorklogCount);
        Assert.Equal(2, document.Totals.EntriesByCategory["Issue"]);
    }

    // ###########################################################################################
    // A schematic image is included only when the board data names one AND the file is really
    // there. A data root out of sync with a board's Excel file is a state DataManager already
    // tolerates elsewhere, and it must cost the section its picture, not the export.
    // ###########################################################################################
    [Fact]
    public void A_schematic_image_is_included_when_it_exists_and_dropped_when_it_does_not()
    {
        string real = Path.Combine(this.thisWorkspace.Root, "sheet1.png");
        File.WriteAllBytes(real, new byte[] { 1, 2, 3 });

        var document = Build(Workbook(), new[] { Entry(1, "Sheet 1"), Entry(2, "Missing") },
            images: new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Sheet 1"] = real,
                ["Missing"] = Path.Combine(this.thisWorkspace.Root, "gone.png")
            });

        Assert.Equal(real, document.Sections.Single(s => s.SchematicName == "Sheet 1").SchematicImagePath);
        Assert.Null(document.Sections.Single(s => s.SchematicName == "Missing").SchematicImagePath);
    }

    // Schematic names are matched case-insensitively here as they are everywhere else in this app -
    // "sheet 1" in an entry must find "Sheet 1" in the board data.
    [Fact]
    public void A_schematic_image_is_matched_case_insensitively()
    {
        string real = Path.Combine(this.thisWorkspace.Root, "sheet1.png");
        File.WriteAllBytes(real, new byte[] { 1 });

        var document = Build(Workbook(), new[] { Entry(1, "sheet 1") },
            images: new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["Sheet 1"] = real });

        Assert.Equal(real, document.Sections.Single().SchematicImagePath);
    }

    // ###########################################################################################
    // Photos and files resolve to absolute paths, in DisplayOrder, with any whose file is missing
    // already dropped - the writer must never be handed a path it then has to decide about. A row
    // survives in entries.json after its file is gone (a hand-edited workbook folder, a restore
    // from a partial backup), and a page reading "photo-3.jpg" over a blank frame is worse than a
    // document that does not mention it.
    // ###########################################################################################
    [Fact]
    public void Attachments_resolve_in_display_order_with_missing_files_dropped()
    {
        string folder = Path.Combine(this.thisWorkspace.Root, "worklog_1");
        Directory.CreateDirectory(folder);
        File.WriteAllBytes(Path.Combine(folder, "second.png"), new byte[] { 1 });
        File.WriteAllBytes(Path.Combine(folder, "first.png"), new byte[] { 1 });

        var entry = Entry(1, "Sheet 1");
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "second.png", DisplayOrder = 2 });
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 2, FileName = "first.png", DisplayOrder = 1, Comment = "before" });
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 3, FileName = "gone.png", DisplayOrder = 3 });

        var document = Build(Workbook(), new[] { entry }, attachments: _ => folder);

        var photos = document.Sections.Single().Entries.Single().Photos;
        Assert.Equal(new[] { "first.png", "second.png" }, photos.Select(p => p.FileName));
        Assert.Equal("before", photos[0].Comment);
        Assert.Equal(Path.Combine(folder, "first.png"), photos[0].FullPath);
    }

    [Fact]
    public void Photos_and_files_are_kept_apart()
    {
        string folder = Path.Combine(this.thisWorkspace.Root, "worklog_1");
        Directory.CreateDirectory(folder);
        File.WriteAllBytes(Path.Combine(folder, "board.png"), new byte[] { 1 });
        File.WriteAllBytes(Path.Combine(folder, "datasheet.pdf"), new byte[] { 1 });

        var entry = Entry(1, "Sheet 1");
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "board.png" });
        entry.Files.Add(new WorklogAttachmentRecord { Id = 1, FileName = "datasheet.pdf" });

        var document = Build(Workbook(), new[] { entry }, attachments: _ => folder);

        var exported = document.Sections.Single().Entries.Single();
        Assert.Equal("board.png", Assert.Single(exported.Photos).FileName);
        Assert.Equal("datasheet.pdf", Assert.Single(exported.Files).FileName);
    }

    // An entry whose attachment folder cannot be resolved at all loses its attachments, not its
    // place in the document - the text of the entry is the part a customer document cannot do
    // without.
    [Fact]
    public void An_entry_whose_attachment_folder_cannot_be_resolved_still_exports_its_text()
    {
        var entry = Entry(1, "Sheet 1", "Still here");
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "board.png" });

        var document = Build(Workbook(), new[] { entry }, attachments: _ => throw new IOException("no folder"));

        var exported = document.Sections.Single().Entries.Single();
        Assert.Equal("Still here", exported.Record.Title);
        Assert.Empty(exported.Photos);
    }

    // ###########################################################################################
    // "Workbook_{id}_{Hardware}_{Board}_{YYYYMMDD}", the requested format.
    //
    // The BoardKey is "Hardware|Board" (Main.GetCurrentBoardKey), so its two halves become their own
    // underscore-separated segments rather than one hyphenated token.
    //
    // The workbook TITLE is deliberately absent: it is a sentence, often carrying a customer's own
    // details, on a file that is about to be emailed. Id, board and date identify it unambiguously.
    // ###########################################################################################
    [Fact]
    public void The_file_name_carries_the_id_hardware_board_and_date_but_not_the_title()
    {
        var workbook = Workbook(title: "SN: HB4 1325119; bought 2023-08-15");
        workbook.BoardKey = "C64|250469";

        var document = Build(workbook, Array.Empty<WorklogEntryRecord>());

        Assert.Equal("Workbook_3_C64_250469_20260904", WorkbookExportModel.BuildFileBaseName(document));
    }

    // A board key with no separator contributes ONE segment rather than an empty trailing one.
    [Fact]
    public void A_board_key_without_a_separator_contributes_a_single_segment()
    {
        var workbook = Workbook();
        workbook.BoardKey = "Amstrad CPC";

        var document = Build(workbook, Array.Empty<WorklogEntryRecord>());

        Assert.Equal("Workbook_3_Amstrad-CPC_20260904", WorkbookExportModel.BuildFileBaseName(document));
    }

    // An empty board key leaves no gap - "Workbook_3__20260904" would be the naive result.
    [Fact]
    public void A_workbook_with_no_board_key_still_gets_a_usable_file_name()
    {
        var workbook = Workbook();
        workbook.BoardKey = "";

        var document = Build(workbook, Array.Empty<WorklogEntryRecord>());

        Assert.Equal("Workbook_3_20260904", WorkbookExportModel.BuildFileBaseName(document));
    }

    // The date is TODAY's, in YYYYMMDD with no separators, so exports of one board sort
    // chronologically in a folder listing.
    [Fact]
    public void The_file_name_dates_the_export_not_the_workbooks_start()
    {
        var workbook = Workbook();
        workbook.StartDate = new DateTime(2020, 1, 2);

        var document = WorkbookExportModel.Build(
            workbook, Array.Empty<WorklogEntryRecord>(), null, _ => null, new DateTime(2026, 12, 31), "DKK");

        Assert.EndsWith("_20261231", WorkbookExportModel.BuildFileBaseName(document));
    }

    [Theory]
    [InlineData("C64/250469", "C64-250469")]
    [InlineData("a b  c", "a-b-c")]
    [InlineData("***", "")]
    [InlineData("  padded  ", "padded")]
    [InlineData(null, "")]
    // Runs of replaced characters collapse to ONE hyphen and the result is trimmed, so a name
    // never arrives as "board---name-" - legal, but plainly machine-made.
    public void File_name_sanitising_collapses_runs_and_trims(string? input, string expected)
    {
        Assert.Equal(expected, WorkbookExportModel.SanitizeForFileName(input));
    }

    // ###########################################################################################
    // The export's own extension is put on the chosen path, replacing the OTHER export format's
    // rather than stacking on top of it.
    //
    // THE BUG: the format comes from which button was pressed, not from the typed name, so typing
    // "repair.pdf" into the "Export to ZIP" dialog produced "repair.pdf.zip". Worse than ugly - the
    // picker had already shown its overwrite prompt for "repair.pdf", so an existing
    // "repair.pdf.zip" was overwritten with no prompt at all.
    // ###########################################################################################
    [Theory]
    [InlineData("C:/exports/repair.zip", "zip", "C:/exports/repair.zip")]
    [InlineData("C:/exports/repair", "zip", "C:/exports/repair.zip")]
    [InlineData("C:/exports/repair.ZIP", "zip", "C:/exports/repair.ZIP")]
    public void An_already_correct_or_missing_extension_is_handled(string path, string extension, string expected)
    {
        Assert.Equal(expected, WorkbookExportModel.EnsureFileExtension(path, extension));
    }

    // The reported case, both ways round.
    [Theory]
    [InlineData("C:/exports/repair.pdf", "zip", "C:/exports/repair.zip")]
    [InlineData("C:/exports/repair.zip", "pdf", "C:/exports/repair.pdf")]
    [InlineData("C:/exports/repair.PDF", "zip", "C:/exports/repair.zip")]
    public void The_other_export_formats_extension_is_replaced_not_appended(string path, string extension, string expected)
    {
        Assert.Equal(expected, WorkbookExportModel.EnsureFileExtension(path, extension));
    }

    // An extension this feature does not own is left alone and the real one appended - a workbook
    // saved as "board rev 2.5" must not lose its tail to something that only looks like a suffix.
    [Theory]
    [InlineData("C:/exports/board rev 2.5", "pdf", "C:/exports/board rev 2.5.pdf")]
    [InlineData("C:/exports/notes.txt", "pdf", "C:/exports/notes.txt.pdf")]
    public void An_unrelated_extension_is_kept_and_the_real_one_appended(string path, string extension, string expected)
    {
        Assert.Equal(expected, WorkbookExportModel.EnsureFileExtension(path, extension));
    }

    [Fact]
    public void A_blank_path_is_returned_unchanged()
    {
        Assert.Equal(string.Empty, WorkbookExportModel.EnsureFileExtension(string.Empty, "pdf"));
    }
}
