using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for BoardDataReader - the Excel side of a board contribution.
//
// Every board in Assets/Data is an .xlsx authored by a contributor, so the sheet names and
// column headers are a public contract. These tests build a workbook in that exact shape
// (see BoardWorkbookBuilder) and assert the reader maps it correctly; if a header constant is
// renamed without a data migration, they fail.
// BoardDataReader keeps its loaded boards in a static cache (a ConcurrentDictionary, because the
// app itself loads boards from the UI thread and background tasks at once). Thread-safe or not,
// the cache is still shared static state: this class and BoardDataWriterTests add to it, remove
// from it, and one test wipes it entirely with ClearAllCache, so they share one collection and
// run sequentially - the same treatment UserSettings and DataManager already get.
[Collection("BoardData")]
public sealed class BoardDataReaderTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();
    private readonly List<string> thisCacheKeys = new();

    public void Dispose()
    {
        // BoardDataReader caches by key in a static dictionary; leave none behind for other tests.
        foreach (string key in this.thisCacheKeys)
        {
            BoardDataReader.ClearCache(key);
        }

        this.thisWorkspace.Dispose();
    }

    private string NewCacheKey()
    {
        string key = "test-" + Guid.NewGuid().ToString("N");
        this.thisCacheKeys.Add(key);
        return key;
    }

    private Task<BoardData?> LoadAsync(string excelPath) =>
        BoardDataReader.LoadAsync(excelPath, this.NewCacheKey());

    private string CompleteBoard() =>
        BoardWorkbookBuilder.WriteCompleteBoard(this.thisWorkspace.Path_("board.xlsx"));

    // --------------------------------------------------------------------- loading

    [Fact]
    public async Task LoadAsync_returns_null_when_the_file_does_not_exist()
    {
        Assert.Null(await this.LoadAsync(this.thisWorkspace.Path_("missing.xlsx")));
    }

    [Fact]
    public async Task Concurrent_loads_and_cache_clears_do_not_corrupt_the_cache()
    {
        // This concurrency is not hypothetical: at every contributor-mode startup,
        // DataValidator.ValidateAllDataAsync walks EVERY board through LoadAsync from a
        // Task.Run background task while the UI thread loads the selected board through the
        // same static cache - and LoadAsync itself writes the cache from inside its own
        // Task.Run. ClearAllCache lands on top from the manual-sync path. The cache must
        // survive that; corruption shows up as a thrown "non-concurrent collections"
        // exception, a board silently coming back null, or a hang in a corrupted bucket
        // chain (which the timeout below converts into a plain failure).
        string excel = this.CompleteBoard();

        const int loadsPerRound = 16;
        const int rounds = 8;

        for (int round = 0; round < rounds; round++)
        {
            List<string> keys = Enumerable.Range(0, loadsPerRound)
                .Select(_ => this.NewCacheKey())
                .ToList();

            Task<BoardData?>[] loads = keys
                .Select(key => Task.Run(() => BoardDataReader.LoadAsync(excel, key)))
                .ToArray();

            Task<BoardData?[]> allLoads = Task.WhenAll(loads);
            Task finished = await Task.WhenAny(allLoads, Task.Delay(TimeSpan.FromSeconds(30)));

            Assert.True(finished == allLoads,
                "concurrent loads did not complete - the cache likely corrupted into a cycle");

            foreach (BoardData? data in await allLoads)
            {
                Assert.NotNull(data);
                Assert.Single(data!.Schematics);
            }

            // Cached re-reads must also survive; then clear so the next round races the
            // first-load inserts (and their resizes) again from an empty dictionary.
            foreach (string key in keys)
            {
                Assert.NotNull(await BoardDataReader.LoadAsync(excel, key));
            }

            BoardDataReader.ClearAllCache();
        }
    }

    [Fact]
    public async Task LoadAsync_returns_null_for_a_file_that_is_not_a_workbook()
    {
        string fake = this.thisWorkspace.WriteFile("broken.xlsx", "this is not a zip archive");
        Assert.Null(await this.LoadAsync(fake));
    }

    [Fact]
    public async Task LoadAsync_reads_a_complete_board()
    {
        BoardData? data = await this.LoadAsync(this.CompleteBoard());

        Assert.NotNull(data);
        Assert.Single(data!.Schematics);
        Assert.Equal(2, data.Components.Count);
        Assert.Single(data.ComponentImages);
        Assert.Single(data.ComponentLocalFiles);
        Assert.Single(data.ComponentLinks);
        Assert.Single(data.BoardLocalFiles);
        Assert.Single(data.BoardLinks);
        Assert.Single(data.Credits);
        Assert.Single(data.KiCadImportantSignals);
    }

    [Fact]
    public async Task The_revision_date_is_read_from_the_comment_above_the_header()
    {
        BoardData? data = await this.LoadAsync(
            BoardWorkbookBuilder.WriteCompleteBoard(this.thisWorkspace.Path_("board.xlsx"), "2026-03-01"));

        Assert.Equal("2026-03-01", data!.RevisionDate);
    }

    [Fact]
    public async Task The_header_row_is_found_below_free_text_comment_rows()
    {
        // Contributors put notes above the header; the reader scans for the header row rather
        // than assuming row 1.
        string path = new BoardWorkbookBuilder()
            .SheetWithLeadingRows(
                "Board schematics",
                new[] { "# Revision date: 2026-02-02", "# Notes: anything at all", "" },
                BoardWorkbookBuilder.SchematicsHeaders,
                new[] { "u", "Sheet 9", "", "s9.png", "", "", "", "", "" })
            .SaveTo(this.thisWorkspace.Path_("leading.xlsx"));

        BoardData? data = await this.LoadAsync(path);

        Assert.Equal("Sheet 9", Assert.Single(data!.Schematics).SchematicName);
    }

    // --------------------------------------------------------------------- mapping

    [Fact]
    public async Task Schematic_columns_are_mapped_to_the_schematic_model()
    {
        BoardData? data = await this.LoadAsync(this.CompleteBoard());

        BoardSchematicEntry schematic = Assert.Single(data!.Schematics);

        Assert.Equal("Sheet 1", schematic.SchematicName);
        Assert.Equal("board.kicad_pcb", schematic.CadName);
        Assert.Equal("sheet1.png", schematic.SchematicImageFile);
        Assert.Equal("#FF0000", schematic.SchematicHighlightColor);
        Assert.Equal("0.5", schematic.SchematicHighlightOpacity);
        Assert.Equal("#00FF00", schematic.OppositeTraceHighlightColor);
        Assert.Equal("#0000FF", schematic.ThumbnailHighlightColor);
    }

    [Fact]
    public async Task Component_columns_are_mapped_to_the_component_model()
    {
        BoardData? data = await this.LoadAsync(this.CompleteBoard());

        ComponentEntry pla = data!.Components.Single(c => c.BoardLabel == "U1");

        Assert.Equal("PLA", pla.FriendlyName);
        Assert.Equal("906114-01", pla.TechnicalNameOrValue);
        Assert.Equal("906114", pla.PartNumber);
        Assert.Equal("IC", pla.Category);
        Assert.Equal("PAL", pla.Region);
        Assert.Equal("Programmable logic array", pla.Description);
    }

    [Fact]
    public async Task Component_image_oscilloscope_columns_are_mapped()
    {
        // These three feed straight into ScopeValueMapper.
        BoardData? data = await this.LoadAsync(this.CompleteBoard());

        ComponentImageEntry image = Assert.Single(data!.ComponentImages);

        Assert.Equal("U1", image.BoardLabel);
        Assert.Equal("1", image.Pin);
        Assert.Equal("5ms", image.TimeDiv);
        Assert.Equal("500mV", image.VoltsDiv);
        Assert.Equal("1.65V", image.TriggerLevelVolts);
    }

    [Fact]
    public async Task Important_signal_rows_are_mapped()
    {
        BoardData? data = await this.LoadAsync(this.CompleteBoard());

        var signal = Assert.Single(data!.KiCadImportantSignals);

        Assert.Equal("Clock", signal.DisplayName);
        Assert.Equal("/Sheet1/CLK", signal.KiCadNetName);
    }

    [Fact]
    public async Task An_empty_cell_becomes_an_empty_string_not_null()
    {
        // The whole model is non-nullable strings; a blank Excel cell must not produce null.
        BoardData? data = await this.LoadAsync(this.CompleteBoard());

        ComponentEntry cap = data!.Components.Single(c => c.BoardLabel == "C1");

        Assert.Equal(string.Empty, cap.PartNumber);
        Assert.Equal(string.Empty, cap.Region);
    }

    [Fact]
    public async Task Fully_blank_rows_are_skipped()
    {
        string path = new BoardWorkbookBuilder()
            .Sheet("Components", BoardWorkbookBuilder.ComponentsHeaders,
                new[] { "u1", "U1", "PLA", "906114", null, null, null, null },
                new string?[] { null, null, null, null, null, null, null, null },
                new[] { "u2", "U2", "VIC", "6569", null, null, null, null })
            .SaveTo(this.thisWorkspace.Path_("blanks.xlsx"));

        BoardData? data = await this.LoadAsync(path);

        Assert.Equal(2, data!.Components.Count);
    }

    [Fact]
    public async Task A_missing_sheet_yields_an_empty_list_rather_than_failing_the_load()
    {
        // Older or partial contributions do not have every sheet.
        string path = new BoardWorkbookBuilder()
            .Sheet("Components", BoardWorkbookBuilder.ComponentsHeaders,
                new[] { "u1", "U1", "PLA", "906114", null, null, null, null })
            .SaveTo(this.thisWorkspace.Path_("partial.xlsx"));

        BoardData? data = await this.LoadAsync(path);

        Assert.NotNull(data);
        Assert.Single(data!.Components);
        Assert.Empty(data.Schematics);
        Assert.Empty(data.Credits);
        Assert.Empty(data.BoardLinks);
    }

    [Fact]
    public async Task A_sheet_whose_headers_do_not_match_yields_no_rows()
    {
        string path = new BoardWorkbookBuilder()
            .Sheet("Components", new[] { "Wrong", "Headers", "Entirely" },
                new[] { "a", "b", "c" })
            .SaveTo(this.thisWorkspace.Path_("badheaders.xlsx"));

        BoardData? data = await this.LoadAsync(path);

        Assert.NotNull(data);
        Assert.Empty(data!.Components);
    }

    [Fact]
    public async Task Column_order_does_not_matter()
    {
        // The reader maps by header name, so contributors can reorder columns.
        string[] reversed = BoardWorkbookBuilder.ComponentsHeaders.Reverse().ToArray();
        string?[] reversedRow = new string?[]
        {
            "Programmable logic array", "PAL", "IC", "906114", "906114-01", "PLA", "U1", "uuid-2"
        };

        string path = new BoardWorkbookBuilder()
            .Sheet("Components", reversed, reversedRow)
            .SaveTo(this.thisWorkspace.Path_("reordered.xlsx"));

        BoardData? data = await this.LoadAsync(path);

        ComponentEntry component = Assert.Single(data!.Components);
        Assert.Equal("U1", component.BoardLabel);
        Assert.Equal("PLA", component.FriendlyName);
        Assert.Equal("IC", component.Category);
    }

    // -------------------------------------------------------------- highlight JSON

    [Fact]
    public async Task Component_highlights_come_from_the_sidecar_json_not_from_excel()
    {
        // Highlights moved out of Excel into the board .json; the reader pulls them in.
        string excel = this.CompleteBoard();

        BoardComponentHighlightStorage.SaveComponentHighlights(
            excel,
            "Sheet 1",
            new[]
            {
                new LabelEditorSaveRow
                {
                    SchematicName = "Sheet 1", BoardLabel = "U1", X = 5, Y = 6, Width = 7, Height = 8
                }
            });

        BoardData? data = await this.LoadAsync(excel);

        ComponentHighlightEntry highlight = Assert.Single(data!.ComponentHighlights);
        Assert.Equal("Sheet 1", highlight.SchematicName);
        Assert.Equal("U1", highlight.BoardLabel);
        Assert.Equal("5", highlight.X);
    }

    [Fact]
    public async Task A_board_with_no_highlight_json_simply_has_none()
    {
        BoardData? data = await this.LoadAsync(this.CompleteBoard());
        Assert.Empty(data!.ComponentHighlights);
    }

    // -------------------------------------------------------------------- caching

    [Fact]
    public async Task A_second_load_with_the_same_cache_key_returns_the_cached_instance()
    {
        string excel = this.CompleteBoard();
        string key = this.NewCacheKey();

        BoardData? first = await BoardDataReader.LoadAsync(excel, key);
        BoardData? second = await BoardDataReader.LoadAsync(excel, key);

        Assert.NotNull(first);
        Assert.Same(first, second);
    }

    [Fact]
    public async Task The_cache_is_not_consulted_for_a_different_key()
    {
        string excel = this.CompleteBoard();

        BoardData? first = await BoardDataReader.LoadAsync(excel, this.NewCacheKey());
        BoardData? second = await BoardDataReader.LoadAsync(excel, this.NewCacheKey());

        Assert.NotSame(first, second);
    }

    [Fact]
    public async Task ClearCache_forces_the_next_load_to_re_read_the_file()
    {
        string excel = this.CompleteBoard();
        string key = this.NewCacheKey();

        BoardData? first = await BoardDataReader.LoadAsync(excel, key);
        BoardDataReader.ClearCache(key);
        BoardData? second = await BoardDataReader.LoadAsync(excel, key);

        Assert.NotNull(second);
        Assert.NotSame(first, second);
    }

    [Fact]
    public async Task A_cached_board_is_returned_even_when_the_file_has_since_been_deleted()
    {
        // Documents the current behaviour: the cache is checked before the file exists check.
        string excel = this.CompleteBoard();
        string key = this.NewCacheKey();

        BoardData? first = await BoardDataReader.LoadAsync(excel, key);
        File.Delete(excel);

        Assert.Same(first, await BoardDataReader.LoadAsync(excel, key));
    }

    // ###########################################################################################
    // *** THE DRAFT-OVERLAY TESTS ARE GONE, AND SO IS THE OVERLAY (Phase 6, 2026-09-23). ***
    //
    // Five tests lived here. They pinned that LoadAsync applied a BoardDraft over the parsed
    // board - on a fresh read, on a cache hit, without ever baking the draft into the cache, and
    // still correctly after a sync-style clear and rewrite. All of that was real and worth
    // pinning while a drafted board was a COMPUTATION.
    //
    // A draft is now a board folder with its own workbook, so loading a drafted board means
    // reading that file (DraftBoardSource decides which). There is no draft argument, no merge,
    // and nothing for these tests to assert. They are deleted rather than rewritten because the
    // behaviour they described no longer exists - this note is here so their absence reads as a
    // decision.
    //
    // What replaced them: DataManagerDraftOverlayTests (which workbook is read, app level),
    // DraftBoardSourceTests (the rule itself) and BoardDataDifferTests (what changed).
    //
    // The cache tests above are untouched and matter MORE than they did - the cache is now keyed
    // by the file that was read, precisely because a drafted board has two of them.
    // ###########################################################################################

    // ------------------------------------------------- CollectReferencedLocalFiles

    [Fact]
    public void CollectReferencedLocalFiles_returns_every_file_named_by_the_workbook()
    {
        string excel = this.CompleteBoard();

        HashSet<string> files = BoardDataReader.CollectReferencedLocalFiles(excel);

        Assert.Contains("sheet1.png", files);
        Assert.Contains("u1-pin1.png", files);
        Assert.Contains("906114.pdf", files);
        Assert.Contains("manual.pdf", files);
    }

    [Fact]
    public void CollectReferencedLocalFiles_is_empty_for_a_missing_workbook()
    {
        Assert.Empty(BoardDataReader.CollectReferencedLocalFiles(
            this.thisWorkspace.Path_("missing.xlsx")));
    }

    // ------------------------------------- The revision-only reader (session 2d)

    // THE anti-drift test between the two readers. Both scan for the same "# Revision date:" marker
    // and both feed the drift warning, so if they ever disagreed the app would compare a value
    // against itself read a different way. They share one private scan, and this proves it.
    [Fact]
    public async Task The_revision_only_reader_agrees_with_a_full_load()
    {
        string excel = this.CompleteBoard();

        BoardData? full = await this.LoadAsync(excel);

        Assert.Equal(full!.RevisionDate, BoardDataReader.ReadRevisionDateOnly(excel));
    }

    // The real-world shape, which the shipped boards in Assets/Data actually use - the fixtures'
    // own default is numeric, so without this neither reader would ever see a month name in a test.
    [Fact]
    public void The_revision_only_reader_handles_the_real_world_month_name_format()
    {
        string excel = BoardWorkbookBuilder.WriteCompleteBoard(
            this.thisWorkspace.Path_("month-name-board.xlsx"), "2026-August-21");

        Assert.Equal("2026-August-21", BoardDataReader.ReadRevisionDateOnly(excel));
    }

    [Fact]
    public void The_revision_only_reader_is_empty_for_a_missing_file()
    {
        Assert.Equal(
            string.Empty,
            BoardDataReader.ReadRevisionDateOnly(this.thisWorkspace.Path_("missing.xlsx")));
    }

    [Fact]
    public void The_revision_only_reader_is_empty_for_a_file_that_is_not_a_workbook()
    {
        string path = this.thisWorkspace.Path_("not-a-workbook.xlsx");
        File.WriteAllText(path, "this is not an Excel file");

        Assert.Equal(string.Empty, BoardDataReader.ReadRevisionDateOnly(path));
    }

    // It produces no BoardData, so it must leave the cache exactly as it found it - a board opened
    // for its revision alone must not look "loaded" to anything else.
    [Fact]
    public void The_revision_only_reader_never_populates_the_cache()
    {
        string excel = this.CompleteBoard();
        string cacheKey = this.NewCacheKey();

        BoardDataReader.ReadRevisionDateOnly(excel);

        Assert.Null(BoardDataReader.TryGetCachedRevisionDate(cacheKey));
    }

    [Fact]
    public async Task A_cached_board_reports_its_revision_without_touching_the_disk()
    {
        string cacheKey = this.NewCacheKey();
        string excel = BoardWorkbookBuilder.WriteCompleteBoard(
            this.thisWorkspace.Path_("cached-board.xlsx"), "2026-August-21");

        Assert.Null(BoardDataReader.TryGetCachedRevisionDate(cacheKey));

        await BoardDataReader.LoadAsync(excel, cacheKey);

        Assert.Equal("2026-August-21", BoardDataReader.TryGetCachedRevisionDate(cacheKey));
    }

    // ------------------------------------- A board that exists only as a draft (task 9)

    // The existing behaviour, pinned deliberately: for a board the main workbook DOES list, a
    // missing board file is a genuine sync failure. It must keep returning null rather than quietly
    // rendering as an empty board, which would hide the failure from the user entirely.
    [Fact]
    public async Task A_missing_board_file_is_still_null_when_a_missing_file_is_not_allowed()
    {
        var result = await BoardDataReader.LoadAsync(
            this.thisWorkspace.Path_("missing.xlsx"),
            this.NewCacheKey());

        Assert.Null(result);
    }

    // ###########################################################################################
    // A board created through "Add a new board" has no published .xlsx by construction, and an
    // EMPTY board is the truthful answer for it: officially, none of it exists yet.
    //
    // *** WHAT THIS TEST ASSERTED CHANGED WITH THE MECHANISM. *** It used to prove the DRAFT's
    // own rows came back, because the draft was merged onto an empty board here. The draft
    // workbook is now read directly by the caller (DraftBoardSource), so what this path is still
    // responsible for is narrower and is all that is left to pin: answering an empty board rather
    // than null, so a brand-new board can be opened at all.
    // ###########################################################################################
    // *** A MISSING-FILE BOARD IS NEVER CACHED. ***
    //
    // The cache holds boards parsed from a FILE. Caching this one would pin an empty board
    // against a path that may later hold a real workbook - a published board arriving by sync,
    // or a draft-only board being seeded - and every later load on that key would answer empty.
    //
    // Proved by creating the file between two loads on the same key: a cached first result would
    // make the second one stale.
    // ###########################################################################################
    [Fact]
    public async Task A_missing_file_board_is_never_cached_so_a_later_real_file_is_read()
    {
        string cacheKey = this.NewCacheKey();
        string path = this.thisWorkspace.Path_("appears-later.xlsx");

        var first = await BoardDataReader.LoadAsync(path, cacheKey, allowMissingOfficialFile: true);
        Assert.Empty(first!.Components);

        BoardWorkbookBuilder.WriteCompleteBoard(path);

        var second = await BoardDataReader.LoadAsync(path, cacheKey, allowMissingOfficialFile: true);

        Assert.NotEmpty(second!.Components);
    }
}
