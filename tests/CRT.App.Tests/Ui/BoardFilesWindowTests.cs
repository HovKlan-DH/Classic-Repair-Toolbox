using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// BoardFilesWindow - importing board images and KiCad data into a drafted board, and the report
// saying what the KiCad data actually lines up with (NewContributeStrategy.md Phase 2, session 2c,
// task 9).
//
// WHAT THESE COVER: how one already-built KiCadReferenceMatchReport is RENDERED (which sections
// appear, what each heading says), and the schematic-name collision rule. The matcher's own logic
// is covered exhaustively in CRT.Data.Tests.
//
// WHAT THEY DELIBERATELY DO NOT: the file copies and the platform file/folder pickers. Rule 6 puts
// real filesystem side effects and platform dialogs out of scope, and the copy destination itself
// is already covered by DraftFileResolverTests.
[Collection("HeadlessUi")]
public sealed class BoardFilesWindowTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public BoardFilesWindowTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static KiCadReferenceMatchReport Report(
        string[] matched,
        string[] unmatched,
        string[] unused,
        int pcbFileCount = 1,
        string[]? views = null) =>
        new(matched, unmatched, unused, pcbFileCount, views ?? Array.Empty<string>());

    private static string TextOf(BoardFilesWindow window, string name) =>
        window.GetControl<TextBlock>(name).Text ?? string.Empty;

    private static bool VisibleOf(BoardFilesWindow window, string name) =>
        window.GetControl<StackPanel>(name).IsVisible;

    // The references one of the report's badge lists is showing.
    private static List<string> ItemsOf(BoardFilesWindow window, string name) =>
        (window.GetControl<ItemsControl>(name).ItemsSource ?? Array.Empty<string>())
            .Cast<string>()
            .ToList();

    [Fact]
    public void The_window_constructs_without_throwing()
    {
        UiTest.Run(() => Assert.NotNull(new BoardFilesWindow()));
    }

    [Fact]
    public void The_report_panel_is_hidden_until_there_is_a_report()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            Assert.False(window.GetControl<Border>("ReportPanel").IsVisible);
        });
    }

    // The headline. A component in this list highlights on the image and lights up no copper, with
    // nothing anywhere else in the app to say so - it has to be visible and counted.
    [Fact]
    public void Unmatched_components_are_shown_and_counted()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8" },
                unmatched: new[] { "PLA/U17", "C15" },
                unused: Array.Empty<string>()));

            Assert.True(VisibleOf(window, "UnmatchedPanel"));
            Assert.Equal("2 components will NOT light up", TextOf(window, "UnmatchedHeaderText"));
            Assert.Equal(new[] { "PLA/U17", "C15" }, ItemsOf(window, "UnmatchedList"));
        });
    }

    // The section must only appear when something is genuinely wrong - a permanent "0 components
    // will NOT light up" heading would train the reader to ignore the one line that matters.
    [Fact]
    public void The_unmatched_section_is_hidden_when_everything_matches()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8", "C15" },
                unmatched: Array.Empty<string>(),
                unused: Array.Empty<string>()));

            Assert.False(VisibleOf(window, "UnmatchedPanel"));
            Assert.True(VisibleOf(window, "MatchedPanel"));
        });
    }

    [Fact]
    public void Matched_components_are_shown_and_counted()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8" },
                unmatched: Array.Empty<string>(),
                unused: Array.Empty<string>()));

            Assert.Equal("1 component matches", TextOf(window, "MatchedHeaderText"));
        });
    }

    // Not an error - this is the to-do list for a board being built up, which is exactly what a
    // brand-new board needs.
    [Fact]
    public void Unlabelled_components_are_shown_as_what_is_left_to_do()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: Array.Empty<string>(),
                unmatched: Array.Empty<string>(),
                unused: new[] { "U8", "C15" }));

            Assert.True(VisibleOf(window, "UnusedPanel"));
            Assert.Equal("2 components are not labelled yet", TextOf(window, "UnusedHeaderText"));
        });
    }

    // ###########################################################################################
    // *** EVERY UNLABELLED COMPONENT IS LISTED - NO CAP (owner request, 2026-09-24). ***
    //
    // This list used to stop after 50 with "and 315 more". On a brand-new board that is most of
    // the to-do list the section exists to show. It REPLACES a test that pinned the cap: the cap
    // was removed deliberately, not lost.
    // ###########################################################################################
    [Fact]
    public void A_very_long_unlabelled_list_is_shown_IN_FULL()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();
            string[] unused = Enumerable.Range(1, 365).Select(i => $"C{i}").ToArray();

            window.ShowReport(Report(
                matched: Array.Empty<string>(),
                unmatched: Array.Empty<string>(),
                unused: unused));

            Assert.Equal(unused, ItemsOf(window, "UnusedList"));
        });
    }

    // ###########################################################################################
    // *** EACH COMPONENT IS ITS OWN BADGE (owner request, 2026-09-24). ***
    //
    // Rendered for real (Show, so the WrapPanel realises its containers) and read back off the
    // visual tree: one ReferenceBadge per reference in EVERY component list, each holding exactly
    // its own reference. A comma-joined TextBlock - what this replaced - produces no badges at all.
    // ###########################################################################################
    [Fact]
    public void Every_component_in_every_list_is_drawn_as_its_own_badge()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8", "U9" },
                unmatched: new[] { "PLA/U17" },
                unused: Enumerable.Range(1, 60).Select(i => $"C{i}").ToArray()));

            window.Show();

            try
            {
                foreach ((string list, int expected) in new[] { ("MatchedList", 2), ("UnmatchedList", 1), ("UnusedList", 60) })
                {
                    List<string> badgeTexts = window.GetControl<ItemsControl>(list)
                        .GetVisualDescendants()
                        .OfType<Border>()
                        .Where(border => border.Classes.Contains("ReferenceBadge"))
                        .Select(border => Assert.IsType<TextBlock>(border.Child).Text ?? string.Empty)
                        .ToList();

                    Assert.Equal(expected, badgeTexts.Count);
                    Assert.Equal(ItemsOf(window, list), badgeTexts);
                }
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The Wiki currently tells contributors to read these out of the LOGFILE by hand, then type
    // them into the CAD name column. Showing them here removes that whole step.
    [Fact]
    public void The_cad_view_names_are_listed_so_they_do_not_have_to_be_read_from_a_logfile()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8" },
                unmatched: Array.Empty<string>(),
                unused: Array.Empty<string>(),
                views: new[] { "250407_ - PCB Top", "250407_ - PCB Bottom" }));

            Assert.True(VisibleOf(window, "ViewNamesPanel"));
            Assert.Contains("250407_ - PCB Top", TextOf(window, "ViewNamesListText"));
            Assert.Contains("250407_ - PCB Bottom", TextOf(window, "ViewNamesListText"));
        });
    }

    [Fact]
    public void The_view_names_section_is_hidden_when_the_project_generated_none()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8" },
                unmatched: Array.Empty<string>(),
                unused: Array.Empty<string>()));

            Assert.False(VisibleOf(window, "ViewNamesPanel"));
        });
    }

    [Fact]
    public void The_summary_line_is_shown_with_the_report()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();
            var report = Report(
                matched: new[] { "U8" },
                unmatched: new[] { "PLA/U17" },
                unused: Array.Empty<string>());

            window.ShowReport(report);

            Assert.Equal(report.Summary, TextOf(window, "ReportSummaryText"));
            Assert.True(window.GetControl<Border>("ReportPanel").IsVisible);
        });
    }

    // ------------------------------------------------------ The schematic name collision rule

    // The schematic name is the section's natural key, so a reused name REPLACES the earlier row -
    // importing two files with the same name would silently leave one schematic where two were
    // expected. Within one batch the list on screen has not been reloaded yet, which is precisely
    // why the taken-name set is threaded through the loop rather than read off the list.
    [Fact]
    public void A_free_schematic_name_is_the_desired_one_when_nothing_has_taken_it()
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board", BoardFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    [Fact]
    public void A_taken_schematic_name_gets_a_numbered_suffix()
    {
        var taken = new HashSet<string>(new[] { "Main board" }, StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board (2)", BoardFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    [Fact]
    public void The_suffix_keeps_climbing_past_names_already_taken()
    {
        var taken = new HashSet<string>(
            new[] { "Main board", "Main board (2)", "Main board (3)" },
            StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board (4)", BoardFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    [Fact]
    public void The_collision_check_ignores_casing()
    {
        var taken = new HashSet<string>(new[] { "MAIN BOARD" }, StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board (2)", BoardFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    // A file called ".png" has no name to take - it still needs one, or it would be written as a
    // row with a blank natural key that could never be found or removed again.
    [Fact]
    public void A_blank_desired_name_falls_back_to_a_usable_one()
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Board image", BoardFilesWindow.ResolveFreeSchematicName("   ", taken));
    }

    [Fact]
    public void A_desired_name_is_trimmed_before_it_is_used()
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board", BoardFilesWindow.ResolveFreeSchematicName("  Main board  ", taken));
    }

    // ------------------------------------------------------------------ the row PREVIEWS

    // A minimal but genuinely valid 1x1 PNG. Written as bytes rather than through Avalonia's
    // encoder because the code under test decodes from a plain FileStream - the fixture only has
    // to be something libpng accepts.
    private static readonly byte[] OnePixelPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82
    };

    private const string BoardKey = "Commodore/C64/250407/board.json";

    // Creates a real drafted board holding one schematic row, and optionally the image file that
    // row names. Goes through DraftSeeder and DraftWorkbookStore rather than hand-writing files, so
    // the window reads exactly the shape the application really produces.
    private void SeedBoardWithSchematic(string imageFileName, bool writeTheImageFile)
    {
        DraftSeeder.CreateNewBoard(
            DraftManager.DraftsRoot,
            new NewBoardRegistration
            {
                HardwareName = "C64",
                BoardName = "250407",
                ExcelDataFile = BoardFilesWindowTests.BoardKey,
            });

        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            BoardFilesWindowTests.BoardKey,
            board =>
            {
                // ###########################################################################
                // *** STORED AS THE APPLICATION WRITES IT. *** In CRT 4.0 (2026-10-24) a row
                // names a file of its board ("top.png"), never a path from the data root - the
                // Excel generations' "Commodore/C64/250407/top.png" once hid a defect here
                // (a preview combined by hand from the draft folder and the stored value), so
                // the fixture follows the real format, whatever it is.
                // ###########################################################################
                board.Schematics.Add(new BoardSchematicEntry
                {
                    SchematicName = "Top side",
                    SchematicImageFile = BoardFilesWindowTests.StoredImagePath(imageFileName),
                });

                return board;
            });

        if (writeTheImageFile)
        {
            // The bytes live directly in the draft's board folder, under the bare name.
            File.WriteAllBytes(
                Path.Combine(
                    DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey),
                    imageFileName),
                BoardFilesWindowTests.OnePixelPng);
        }
    }

    // The value a 4.0 board file really holds: the board's own name for the file.
    private static string StoredImagePath(string fileName) => fileName;

    // The window shows ONE section per opening (the Drafts tab has a button for each), and loads
    // only that section's data - so every test names the section it is about.
    private BoardFilesWindow OpenOnSeededBoard(BoardFilesSection section)
    {
        var window = new BoardFilesWindow();

        window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>(), section);
        BoardFilesWindowTests.WaitForOpening(window);

        return window;
    }

    // ###########################################################################################
    // The window reads the draft, the previews and the KiCad data OFF the UI thread, under its
    // "please wait" (owner request, 2026-10-10), so a test lets that finish - as the user waits for
    // it. The reads run on the pool; what they hand back runs here, through RunJobs.
    // ###########################################################################################
    private static void WaitForOpening(BoardFilesWindow window)
    {
        for (int i = 0; i < 1000 && !window.Opening.IsCompleted; i++)
        {
            Dispatcher.UIThread.RunJobs();
            Thread.Sleep(5);
        }

        Dispatcher.UIThread.RunJobs();

        Assert.True(window.Opening.IsCompleted, "The window should have finished reading what it shows.");
    }

    // ###########################################################################################
    // Every change the window makes to the draft is written OFF the UI thread, under its "please
    // wait", one at a time (code review, 2026-10-10 - BoardFilesWindow.DraftChanges.cs). A test lets
    // the queue empty, as the user waits for the overlay. Jobs first: a CAD name pick is queued by
    // a posted job.
    // ###########################################################################################
    private static void WaitForChanges(BoardFilesWindow window)
    {
        Dispatcher.UIThread.RunJobs();

        for (int i = 0; i < 1000 && !window.DraftChanging.IsCompleted; i++)
        {
            Thread.Sleep(5);
            Dispatcher.UIThread.RunJobs();
        }

        Dispatcher.UIThread.RunJobs();

        Assert.True(window.DraftChanging.IsCompleted, "The window should have finished writing the draft.");
    }

    // ------------------------------------------------------------------ ONE section per opening
    //
    // *** THE DRAFTS TAB HAS A BUTTON EACH FOR "Schematic images" AND "KiCad data" (owner
    // request, 2026-09-24), where it used to have one for both. *** Each opens this window on its
    // own section: that section is the only one shown, the title names it, and - the part a
    // screenshot cannot show - only that section's data is loaded. Opening "Schematic images" must
    // not parse the KiCad project, and opening "KiCad data" must not decode the board images.

    [Fact]
    public void Opening_on_SCHEMATIC_IMAGES_shows_and_loads_only_the_images()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);

            Assert.Equal("Schematic images", window.Title);
            Assert.True(window.GetControl<StackPanel>("SchematicsSection").IsVisible);
            Assert.False(window.GetControl<StackPanel>("KiCadSection").IsVisible);

            Assert.Single(window.Schematics);

            // The KiCad file is on disk, and deliberately not listed or parsed for a report. (Its
            // opening IS read, for the CAD names - see BoardFilesWindow.CadNames.cs.)
            Assert.Empty(window.KiCadFiles);
            Assert.False(window.GetControl<Border>("ReportPanel").IsVisible);
        });
    }

    [Fact]
    public void Opening_on_KICAD_DATA_shows_and_loads_only_the_KiCad_files()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            Assert.Equal("KiCad data", window.Title);
            Assert.False(window.GetControl<StackPanel>("SchematicsSection").IsVisible);
            Assert.True(window.GetControl<StackPanel>("KiCadSection").IsVisible);

            Assert.Single(window.KiCadFiles);

            // The schematic row exists in the draft, and its preview is deliberately not decoded.
            Assert.Empty(window.Schematics);
        });
    }

    // The line under the board name describes the section that is open, not both.
    [Fact]
    public void The_introduction_describes_only_the_section_that_is_open()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            string images = TextOf(this.OpenOnSeededBoard(BoardFilesSection.SchematicImages), "IntroText");
            string kiCad = TextOf(this.OpenOnSeededBoard(BoardFilesSection.KiCadData), "IntroText");

            Assert.Contains("schematic images", images, StringComparison.Ordinal);
            Assert.DoesNotContain("KiCad", images, StringComparison.Ordinal);

            Assert.Contains("KiCad project", kiCad, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // *** EACH SCHEMATIC ROW CARRIES A PREVIEW OF ITS OWN IMAGE (owner request, 2026-09-24).
    // ***
    //
    // The list used to name a file and nothing more, so the only way to find out whether the right
    // image had been imported - or whether it was the right way up, or the right side of the board
    // - was to go and open it. This is the "does that look right?" glance.
    // ###########################################################################################
    [Fact]
    public void A_schematic_row_carries_a_preview_of_its_image()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);

            BoardSchematicRow row = Assert.Single(window.Schematics);

            Assert.NotNull(row.Thumbnail);
            Assert.True(row.HasThumbnail);
            Assert.False(row.IsMissing);
        });
    }

    // ###########################################################################################
    // A draft can name an image whose file has been moved or deleted by hand, and this window is
    // the one place a contributor looks to find that out. So a missing file must produce a row
    // that SAYS so rather than an exception or a silently blank line.
    // ###########################################################################################
    [Fact]
    public void A_row_whose_image_file_is_MISSING_still_appears_and_reports_no_preview()
    {
        this.SeedBoardWithSchematic("gone.png", writeTheImageFile: false);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);

            BoardSchematicRow row = Assert.Single(window.Schematics);

            // The row is still listed - the schematic exists in the draft even if its file does not.
            Assert.Equal("Top side", row.SchematicName);
            Assert.Equal(BoardFilesWindowTests.StoredImagePath("gone.png"), row.ImageFile);

            Assert.Null(row.Thumbnail);
            Assert.True(row.IsMissing);
            Assert.False(row.HasThumbnail);
        });
    }

    // ###########################################################################################
    // A file that exists but is NOT a decodable image must not take the window down - a corrupt or
    // half-copied file is exactly what a contributor opens this window to discover.
    //
    // *** THIS ASSERTS SURVIVAL, NOT A NULL PREVIEW, and the distinction is deliberate. *** The
    // headless backend answers a bad decode with a placeholder bitmap rather than throwing, so
    // asserting "no preview" here would pin an Avalonia implementation detail that differs from
    // what a real display does. What this window actually owes the user is that the row is still
    // listed and the import is still usable, which is what is checked.
    // ###########################################################################################
    [Fact]
    public void A_row_whose_image_file_is_CORRUPT_does_not_break_the_window()
    {
        this.SeedBoardWithSchematic("broken.png", writeTheImageFile: false);

        File.WriteAllText(
            Path.Combine(
                DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey),
                "broken.png"),
            "this is not a png");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);

            BoardSchematicRow row = Assert.Single(window.Schematics);

            // Still listed, still named, still removable - the row is not lost with its bytes.
            Assert.Equal("Top side", row.SchematicName);
            Assert.Equal(BoardFilesWindowTests.StoredImagePath("broken.png"), row.ImageFile);

            // And whichever way the decode went, the two display flags still disagree, so the row
            // draws exactly one of the image and the caption.
            Assert.NotEqual(row.HasThumbnail, row.IsMissing);
        });
    }

    // ###########################################################################################
    // The two visibility flags must always disagree - the template shows the Image on one and the
    // "No preview" caption on the other, so if they ever agreed a row would draw both or neither.
    // ###########################################################################################
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Exactly_one_of_the_preview_states_is_true(bool imageExists)
    {
        this.SeedBoardWithSchematic("maybe.png", writeTheImageFile: imageExists);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);

            BoardSchematicRow row = Assert.Single(window.Schematics);

            Assert.NotEqual(row.HasThumbnail, row.IsMissing);
            Assert.Equal(imageExists, row.HasThumbnail);
        });
    }

    // ###########################################################################################
    // *** REBUILDING THE LIST MUST NOT LEAK THE PREVIOUS PASS BITMAPS. ***
    //
    // Every import rebuilds these rows wholesale, and a board scan decodes to tens of megabytes.
    // Reloading twice and finding the first pass bitmaps disposed is what proves the cleanup runs;
    // without it each import silently strands another set.
    // ###########################################################################################
    [Fact]
    public void Reloading_the_list_DISPOSES_the_previous_previews()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);

            Avalonia.Media.Imaging.Bitmap? first = Assert.Single(window.Schematics).Thumbnail;
            Assert.NotNull(first);

            // Re-initialising is what an import does once it has written its files.
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>(), BoardFilesSection.SchematicImages);
            BoardFilesWindowTests.WaitForOpening(window);

            Avalonia.Media.Imaging.Bitmap? second = Assert.Single(window.Schematics).Thumbnail;
            Assert.NotNull(second);

            // A fresh decode, and the old one released.
            Assert.NotSame(first, second);
            Assert.Throws<ObjectDisposedException>(() => first!.PixelSize);
        });
    }

    // ###########################################################################################
    // *** OPENING IS UNDER THE WINDOW'S "PLEASE WAIT" (owner request, 2026-10-10: "it can take a
    // few secs to display, so please do show that new standard wait method"). *** From the first
    // moment - before anything is read - until all of it is in, and then it lets go.
    // ###########################################################################################
    [Theory]
    [InlineData(BoardFilesSection.SchematicImages, CrtWaitWording.OpeningSchematicImages)]
    [InlineData(BoardFilesSection.KiCadData, CrtWaitWording.CheckingKiCadMatches)]
    public void Opening_reads_under_the_windows_please_wait_and_lets_go_when_done(BoardFilesSection section, string sentence)
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>(), section);

            BusyOverlay overlay = BusyOverlay.For(window)!;
            Assert.True(overlay.IsBusy);
            Assert.Equal(sentence, overlay.Message);

            BoardFilesWindowTests.WaitForOpening(window);

            Assert.False(overlay.IsBusy);
        });
    }

    // The previews are decoded off the UI thread; the rows swap in only once ALL of them are in.
    [Fact]
    public void The_rows_appear_with_their_previews_once_the_reading_is_done()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>(), BoardFilesSection.SchematicImages);

            BoardFilesWindowTests.WaitForOpening(window);

            Assert.NotNull(Assert.Single(window.Schematics).Thumbnail);
        });
    }

    // ###########################################################################################
    // *** OPENING AND CLOSING CHANGES NOTHING (owner report, 2026-10-10: closing the window
    // refreshed the detached schematic thumbnails although nothing had changed). *** The Drafts tab
    // reloads the board on screen only when ChangedTheDraft says so - and looking is no change,
    // however the CAD names and the KiCad report were read.
    // ###########################################################################################
    [Theory]
    [InlineData(BoardFilesSection.SchematicImages)]
    [InlineData(BoardFilesSection.KiCadData)]
    public void Opening_and_closing_without_a_change_says_nothing_changed(BoardFilesSection section)
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(section);
            window.Show();
            Dispatcher.UIThread.RunJobs();
            window.Close();

            Assert.False(window.ChangedTheDraft);
        });
    }

    // Closed while it was still reading: what the read hands back has nowhere to go, and is not
    // put on the closed window (its previews are released instead).
    [Fact]
    public void A_window_closed_while_it_reads_is_left_empty()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>(), BoardFilesSection.SchematicImages);
            window.Show();
            window.Close();

            BoardFilesWindowTests.WaitForOpening(window);

            Assert.Empty(window.Schematics);
        });
    }

    // ------------------------------------------------------------------ the imported KiCad FILES

    // Writes KiCad files into the seeded board's "KiCad data" folder, as an import would.
    private void WriteKiCadFiles(params string[] relativePaths)
    {
        string root = Path.Combine(
            DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey),
            "KiCad data");

        foreach (string relativePath in relativePaths)
        {
            string full = Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar));

            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            File.WriteAllText(full, "(kicad_sch (version 20230121))");
        }
    }

    // ###########################################################################################
    // *** THE WINDOW NAMES THE IMPORTED FILES, not just how many there are (owner report,
    // 2026-09-24). ***
    //
    // Reopening this window after an import said "25 KiCad files imported" and nothing else, so a
    // contributor coming back could not tell which project was in the board - an import of the
    // right folder and an import of the wrong one looked identical.
    // ###########################################################################################
    [Fact]
    public void The_imported_KiCad_files_are_LISTED_not_just_counted()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb", "Pages/vic.kicad_sch");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            Assert.Equal(
                new[] { "Open128.kicad_pcb", "Pages/vic.kicad_sch" },
                window.KiCadFiles.Select(row => row.RelativePath).OrderBy(p => p, StringComparer.Ordinal));
        });
    }

    // ###########################################################################################
    // A sub-folder is shown as part of the path, forward-slashed. A multi-sheet project keeps its
    // pages in one, and a flat list of bare file names would hide the structure entirely - which is
    // the very thing worth seeing after the import started following sub-folders.
    // ###########################################################################################
    [Fact]
    public void A_file_in_a_SUB_FOLDER_keeps_its_folder_in_the_listing()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Pages/cpu-8500.kicad_sch");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            BoardKiCadFileRow row = Assert.Single(window.KiCadFiles);

            Assert.Equal("Pages/cpu-8500.kicad_sch", row.RelativePath);

            // Forward slashes whatever the platform separator is - this is how KiCad and the Wiki
            // both write it, and a backslash here would read as a Windows-only detail leaking out.
            Assert.DoesNotContain('\\', row.RelativePath);
        });
    }

    // ###########################################################################################
    // *** THE LISTING HAS NO SCROLLER OF ITS OWN (owner report, 2026-09-24). ***
    //
    // It used to sit in its own ScrollViewer capped at 150px, inside the window's ScrollViewer - a
    // scroll inside a scroll, where the wheel moved whichever one the pointer was over. Every file
    // is now shown at full height and the WINDOW's scroller is the only one. Asserted on the
    // logical tree: exactly one ScrollViewer above the list, and no height cap on its panel.
    // ###########################################################################################
    [Fact]
    public void The_KiCad_file_listing_shows_every_file_with_no_scroller_of_its_own()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles(Enumerable.Range(1, 25).Select(i => $"Pages/sheet{i}.kicad_sch").ToArray());

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            Assert.Equal(25, window.KiCadFiles.Count);

            var list = window.GetControl<ItemsControl>("KiCadFilesItemsControl");

            Assert.Single(list.GetLogicalAncestors().OfType<ScrollViewer>());

            Assert.Equal(
                double.PositiveInfinity,
                window.GetControl<Border>("KiCadFilesPanel").MaxHeight);
        });
    }

    // Size is the cue that separates a real board from a stub export - a .kicad_pcb is tens of
    // megabytes, and a 2 KB one means something went wrong.
    [Fact]
    public void Each_listed_file_reports_its_size()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            Assert.NotEmpty(Assert.Single(window.KiCadFiles).SizeText);
        });
    }

    [Fact]
    public void A_board_with_NO_KiCad_data_lists_nothing_and_says_so()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            Assert.Empty(window.KiCadFiles);

            Assert.Equal(
                "No KiCad data imported yet.",
                window.GetControl<TextBlock>("KiCadStateText").Text);
        });
    }

    // ###########################################################################################
    // The listing is REBUILT rather than appended to. Re-initialising is what an import does once
    // its files are written, and a list that grew each time would show every file twice.
    // ###########################################################################################
    [Fact]
    public void Re_initialising_REPLACES_the_file_listing()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            Assert.Single(window.KiCadFiles);

            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>(), BoardFilesSection.KiCadData);

            Assert.Single(window.KiCadFiles);
        });
    }

    // ------------------------------------------------------------------ REMOVING a KiCad file
    //
    // *** EACH IMPORTED KiCad FILE CAN BE REMOVED ON ITS OWN (owner request, 2026-09-24), the
    // way a schematic image can. *** The delete itself - and what it refuses - is covered by
    // KiCadImportedFilesTests; these cover the window offering it and describing what is left.

    [Fact]
    public void Every_listed_KiCad_file_has_its_own_Remove_button()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb", "Open128.kicad_pro", "Pages/vic.kicad_sch");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);
            window.Show();

            try
            {
                List<Button> removeButtons = window.GetControl<ItemsControl>("KiCadFilesItemsControl")
                    .GetVisualDescendants()
                    .OfType<Button>()
                    .Where(button => (button.Content as string) == "Remove")
                    .ToList();

                Assert.Equal(3, removeButtons.Count);

                // Each button carries ITS OWN row, so a click removes the file on that line.
                Assert.Equal(
                    window.KiCadFiles.Select(row => row.RelativePath),
                    removeButtons.Select(button => Assert.IsType<BoardKiCadFileRow>(button.Tag).RelativePath));
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public async Task Removing_a_KiCad_file_deletes_it_and_the_listing_follows()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Pages/vic.kicad_sch");

        // A REAL minimal PCB for the file that stays: the report is rebuilt after the remove, and an
        // unreadable leftover would (correctly) replace the "Removed" message with a read error.
        File.WriteAllText(
            Path.Combine(
                DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey),
                "KiCad data",
                "Open128.kicad_pcb"),
            BoardFilesWindowTests.MinimalPcb);

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            BoardKiCadFileRow page = window.KiCadFiles.Single(row => row.RelativePath == "Pages/vic.kicad_sch");

            await window.RemoveKiCadFileAsync(page);

            Assert.False(File.Exists(page.FullPath));
            Assert.Equal(new[] { "Open128.kicad_pcb" }, window.KiCadFiles.Select(row => row.RelativePath));
            Assert.Equal("1 KiCad file imported:", window.GetControl<TextBlock>("KiCadStateText").Text);
            Assert.Contains("Removed [Pages/vic.kicad_sch]", window.StatusTextForTests, StringComparison.Ordinal);
            Assert.True(window.ChangedTheDraft);
        });
    }

    // ###########################################################################################
    // *** THE KiCad COPY IS NEVER MADE ON THE UI THREAD (code review, 2026-10-25). *** The KiCad data
    // window made the board's KiCad copy - tens of megabytes on a large board - on the UI thread,
    // twice per refresh, so the window froze with its "please wait" unable to paint. Opening, an
    // import and a removal now each read it once, on the pool.
    // ###########################################################################################
    [Fact]
    public async Task The_KiCad_data_window_reads_the_boards_KiCad_files_off_the_UI_thread()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb", "Pages/vic.kicad_sch");

        var readOnUiThread = new List<bool>();

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow { ReadingKiCadForTests = () => readOnUiThread.Add(Dispatcher.UIThread.CheckAccess()) };

            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>(), BoardFilesSection.KiCadData);
            BoardFilesWindowTests.WaitForOpening(window);

            Assert.Equal(2, window.KiCadFiles.Count);

            await window.RemoveKiCadFileAsync(window.KiCadFiles.Single(row => row.RelativePath == "Pages/vic.kicad_sch"));

            Assert.Single(window.KiCadFiles);
        });

        Assert.Equal([false, false], readOnUiThread);
    }

    // ###########################################################################################
    // *** REMOVE ON A PUBLISHED KiCad FILE THE CONTRIBUTOR IMPORTED AGAIN BRINGS THE PUBLISHED ONE
    // BACK (code review, 2026-10-25). *** It took the name off the draft's list, so the board lost
    // the file at the next submission; now the draft lists the published file again, the row leaves
    // the list of imports, and the status line says the published file is used.
    // ###########################################################################################
    [Fact]
    public async Task Removing_a_published_KiCad_file_imported_again_brings_the_published_one_back()
    {
        string dataRoot = Path.Combine(this.thisWorkspace.Root, "Data");
        BoardDocument.StoredFileRow pcb = StoredFiles.Put(dataRoot, "KiCad data/board.kicad_pcb", BoardFilesWindowTests.MinimalPcb);
        StoredFiles.WriteBoard(
            DraftBoardSource.PublishedPathOf(dataRoot, BoardFilesWindowTests.BoardKey),
            new BoardData { Components = [new ComponentEntry { BoardLabel = "U1" }] },
            pcb);
        Assert.True(DraftSeeder.SeedFromPublishedFile(DraftManager.DraftsRoot, dataRoot, BoardFilesWindowTests.BoardKey).Created);

        // The contributor's own copy of the same file, with other bytes, lying in the draft folder.
        this.WriteKiCadFiles("board.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow { DataRootOverrideForTests = dataRoot };
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, ["U1"], BoardFilesSection.KiCadData);
            BoardFilesWindowTests.WaitForOpening(window);

            await window.RemoveKiCadFileAsync(Assert.Single(window.KiCadFiles));

            Assert.Empty(window.KiCadFiles);
            Assert.Contains("the published file is used again", window.StatusTextForTests, StringComparison.Ordinal);
        });

        string draftBoardFile = Path.Combine(DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey), DataTreeLayout.BoardFileName);
        BoardDocument.StoredFileRow listed = Assert.Single(BoardFiles.ReadDocument(draftBoardFile)!.Files, file => file.Name == "KiCad data/board.kicad_pcb");
        Assert.Equal(pcb.Sha256, listed.Sha256);
    }

    // Removing the last file puts the window back where it was before any import.
    [Fact]
    public async Task Removing_the_LAST_KiCad_file_says_there_is_none_and_hides_the_list()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            await window.RemoveKiCadFileAsync(Assert.Single(window.KiCadFiles));

            Assert.Empty(window.KiCadFiles);
            Assert.Equal("No KiCad data imported yet.", window.GetControl<TextBlock>("KiCadStateText").Text);
            Assert.False(window.GetControl<Border>("KiCadFilesPanel").IsVisible);
            Assert.False(window.GetControl<Border>("ReportPanel").IsVisible);
        });
    }

    // ------------------------------------------------------------------ the report, on REOPEN

    // A minimal but genuinely parseable PCB: one footprint, so the matcher has a reference to
    // compare board labels against.
    private const string MinimalPcb = """
    (kicad_pcb (version 20221018) (generator pcbnew)
      (net 0 "")
      (net 1 "GND")
      (footprint "Package_DIP:DIP-40" (layer "F.Cu")
        (at 100 50)
        (property "Reference" "U1")
        (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 1 "GND"))
      )
    )
    """;

    // ###########################################################################################
    // *** THE MATCH REPORT IS REBUILT WHEN THE WINDOW REOPENS (owner report, 2026-09-24). ***
    //
    // It used to be produced only at the end of an import, so a contributor returning to a board
    // that already had KiCad data saw an empty panel. The report is the one thing that says whether
    // the board labels line up with the KiCad references - "the number one cause of 'I did
    // everything and no traces appear'", in the Wiki's own words - and it was available exactly
    // once, to whoever happened to run the import.
    //
    // RunAsync rather than Run: Initialize kicks the parse off as fire-and-forget, so the assertion
    // has to let the dispatcher finish it. Blocking on it inside UiTest.Run would deadlock - see
    // UiTest's own header.
    // ###########################################################################################
    [Fact]
    public async Task Reopening_on_EXISTING_KiCad_data_rebuilds_the_match_report()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string root = Path.Combine(
            DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey),
            "KiCad data");

        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, "board.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow();

            // A board label that DOES match the footprint above, so the report has something to say.
            window.Initialize(
                "C64 - 250407",
                BoardFilesWindowTests.BoardKey,
                new[] { "U1" },
                BoardFilesSection.KiCadData);

            // Let the fire-and-forget parse complete.
            for (int i = 0; i < 100 && !window.GetControl<Border>("ReportPanel").IsVisible; i++)
            {
                await Task.Delay(20);
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
            }

            Assert.True(
                window.GetControl<Border>("ReportPanel").IsVisible,
                "The match report should be rebuilt when the window opens on existing KiCad data.");
        });
    }

    // ###########################################################################################
    // *** DRAGGING THE SCHEMATIC IMAGES INTO A NEW ORDER (owner request, 2026-09-27). ***
    //
    // "The same move functionality which is already present in the worklog modal": each image its
    // own panel, dragged up or down with a dashed placeholder showing where it will land. Driven
    // with REAL pointer input in a shown window - press, move, release - because what went wrong
    // before (rows "rapidly switching position and cannot settle") only exists in that sequence.
    // ###########################################################################################

    private static readonly string[] FiveViews = ["PCB; Top", "PCB; Bottom", "Address Multiplexing", "Clock", "CPU 8502"];

    // A drafted board holding these schematic rows, in this order. No image files: the rows are
    // the point here, and a row with no file still draws a same-sized preview box.
    private void SeedBoardWithSchematics(params string[] names)
    {
        DraftSeeder.CreateNewBoard(
            DraftManager.DraftsRoot,
            new NewBoardRegistration
            {
                HardwareName = "C64",
                BoardName = "250407",
                ExcelDataFile = BoardFilesWindowTests.BoardKey,
            });

        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            BoardFilesWindowTests.BoardKey,
            board =>
            {
                foreach (string name in names)
                {
                    board.Schematics.Add(new BoardSchematicEntry
                    {
                        SchematicName = name,
                        SchematicImageFile = BoardFilesWindowTests.StoredImagePath(name + ".png"),
                    });
                }

                return board;
            });
    }

    // Tall enough by default that five rows are all on screen - a press on a row scrolled out of
    // sight lands on nothing. The auto-scroll test asks for a short window on purpose. (1000 until
    // each row gained its Name and CAD name line, 2026-10-10.)
    private BoardFilesWindow ShowOnSeededBoard(double height = 1400)
    {
        BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);
        window.Height = height;
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    // Where a hand presses to drag a row: its drag handle - the image and its names. The Name box
    // and the CAD name under them are not part of it (2026-10-10), and a row's middle can fall on
    // them, so a test presses here and only AIMS by RowCentres.
    private static Point HandleCentre(BoardFilesWindow window, int index) =>
        CentreOf(window, HandleOf(window, index));

    private static Control RowContainer(BoardFilesWindow window, int index) =>
        window.GetControl<ItemsControl>("SchematicsItemsControl").ContainerFromIndex(index)!;

    private static Grid HandleOf(BoardFilesWindow window, int index) =>
        RowContainer(window, index).GetVisualDescendants().OfType<Grid>()
            .Single(grid => grid.Classes.Contains("SchematicDragHandle"));

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    // The centre of every row as laid out BEFORE the drag - where a hand would aim.
    private static Point[] RowCentres(BoardFilesWindow window) =>
        Enumerable.Range(0, window.Schematics.Count).Select(i => CentreOf(window, RowContainer(window, i))).ToArray();

    private static void PressAt(Window window, Point point)
    {
        window.MouseDown(point, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static void DragTo(Window window, Point point)
    {
        window.MouseMove(point, Avalonia.Input.RawInputModifiers.LeftMouseButton);
        Dispatcher.UIThread.RunJobs();
    }

    private static void ReleaseAt(Window window, Point point)
    {
        window.MouseUp(point, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();

        // The new order is saved as one of the window's queued changes.
        if (window is BoardFilesWindow files)
        {
            WaitForChanges(files);
        }
    }

    private static List<string> Order(BoardFilesWindow window) =>
        window.Schematics.Select(row => row.SchematicName).ToList();

    private static List<string> DraftOrder() =>
        DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!
            .Schematics.Select(entry => entry.SchematicName).ToList();

    private static DateTime DraftWrittenAt() =>
        File.GetLastWriteTimeUtc(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey));

    [Fact]
    public void Every_schematic_image_is_its_own_panel_that_can_be_dragged_but_not_by_its_Remove_button()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            for (int i = 0; i < FiveViews.Length; i++)
            {
                Control container = RowContainer(window, i);

                Border panel = container.GetVisualDescendants().OfType<Border>()
                    .First(border => border.CornerRadius.TopLeft == 3 && border.BorderThickness.Top == 1);
                Assert.True(panel.IsEffectivelyVisible);

                // The whole row but its Remove button is the handle, with the move cursor. Cursor
                // has no value equality, so it is compared by name.
                Grid handle = HandleOf(window, i);
                Assert.Equal(
                    new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.SizeNorthSouth).ToString(),
                    handle.Cursor?.ToString());

                Button remove = container.GetVisualDescendants().OfType<Button>().Single();
                Assert.DoesNotContain(remove, handle.GetVisualDescendants());
            }

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A VISIBLE MOVE ICON ON EVERY PANEL (owner request, 2026-09-27): "I need a move icon to be
    // visible, so it is clear it can be moved". *** The cursor only says so once the pointer is on
    // the row. The grip is Font Awesome's "grip-vertical" (U+F58E, present in the shipped font and
    // inside its ascent, so not clipped), drawn INSIDE the handle - pressing the icon itself is the
    // most natural place to start a drag, so it must start one.
    // ###########################################################################################
    [Fact]
    public void Every_panel_shows_a_move_grip_inside_its_drag_handle()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            for (int i = 0; i < FiveViews.Length; i++)
            {
                TextBlock grip = HandleOf(window, i).GetVisualDescendants().OfType<TextBlock>()
                    .Single(text => text.Classes.Contains("SchematicDragGrip"));

                Assert.True(grip.IsEffectivelyVisible);
                Assert.Equal("", grip.Text);
                Assert.True(grip.Bounds.Width > 0 && grip.Bounds.Height > 0);
            }

            window.Close();
        });
    }

    [Fact]
    public void Pressing_the_grip_itself_starts_the_drag()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);

            TextBlock grip = HandleOf(window, 0).GetVisualDescendants().OfType<TextBlock>()
                .Single(text => text.Classes.Contains("SchematicDragGrip"));

            // Measured before pressing: once the drag starts, this row is the placeholder and its
            // grip is no longer drawn.
            Point onGrip = CentreOf(window, grip);
            Point target = new(onGrip.X, centres[2].Y + 5);

            PressAt(window, onGrip);
            DragTo(window, target);
            ReleaseAt(window, target);

            Assert.Equal(["PCB; Bottom", "Address Multiplexing", "PCB; Top", "Clock", "CPU 8502"], DraftOrder());

            window.Close();
        });
    }

    [Fact]
    public void Dragging_a_row_turns_it_into_a_dashed_placeholder_of_its_own_height_and_the_others_make_room()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);
            double rowHeight = RowContainer(window, 0).Bounds.Height;

            PressAt(window, CentreOf(window, HandleOf(window, 0)));
            DragTo(window, centres[2] + new Point(0, 5));

            BoardSchematicRow dragged = window.Schematics.Single(row => row.SchematicName == "PCB; Top");

            Assert.True(dragged.IsDropPlaceholder);
            Assert.Equal(rowHeight, dragged.PlaceholderHeight, precision: 1);
            Assert.All(window.Schematics.Where(row => row != dragged), row => Assert.False(row.IsDropPlaceholder));

            // Drawn as the dashed slot, in the place the drop will land.
            Control slot = RowContainer(window, window.Schematics.IndexOf(dragged));
            Avalonia.Controls.Shapes.Rectangle dashes = slot.GetVisualDescendants().OfType<Avalonia.Controls.Shapes.Rectangle>().Single();
            Assert.True(dashes.IsEffectivelyVisible);
            Assert.NotNull(dashes.StrokeDashArray);

            Assert.Equal(["PCB; Bottom", "Address Multiplexing", "PCB; Top", "Clock", "CPU 8502"], Order(window));

            ReleaseAt(window, centres[2] + new Point(0, 5));
            window.Close();
        });
    }

    // ###########################################################################################
    // *** WHAT ACTUALLY MADE THE WORKLOG ROWS "SWITCH POSITION AND NOT SETTLE". *** A gap drawn at
    // any height but the row's own shifts every row below it the moment the drag starts, and again
    // with every move - so the rows run from under a still pointer. The worklog's placeholder was
    // drawn at a fixed height for exactly that reason once (its PlaceholderHeight did not notify).
    // Here with a row made TALLER than the rest by a long image file name (it wraps under the
    // row's name), so a fixed default cannot pass by coincidence: the gap is the row's own height,
    // and the row below it does not move. It was a long CAD name until the CAD name got its own
    // drop-down (2026-10-10).
    // ###########################################################################################
    [Fact]
    public void A_tall_rows_placeholder_is_exactly_its_height_so_nothing_below_it_moves()
    {
        this.SeedBoardWithSchematics(FiveViews);

        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            BoardFilesWindowTests.BoardKey,
            board =>
            {
                BoardSchematicEntry plain = board.Schematics[2];
                board.Schematics[2] = new BoardSchematicEntry
                {
                    SchematicName = plain.SchematicName,
                    SchematicImageFile = BoardFilesWindowTests.StoredImagePath(
                        string.Join(" ", Enumerable.Repeat("Open128 - Address Multiplexing, both banks", 14)) + ".png"),
                };

                return board;
            });

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            double tall = RowContainer(window, 2).Bounds.Height;
            Assert.True(tall > RowContainer(window, 0).Bounds.Height + 20, "The long image file name must make the fixture row taller.");

            Point[] centres = RowCentres(window);
            Point rowBelow = centres[3];

            // Pressed and moved within its own slot: it becomes the gap and nothing reorders.
            PressAt(window, HandleCentre(window, 2));
            DragTo(window, centres[2] + new Point(0, 10));

            Assert.True(window.Schematics[2].IsDropPlaceholder);
            Assert.Equal(tall, RowContainer(window, 2).Bounds.Height, precision: 1);
            Assert.Equal(rowBelow.Y, CentreOf(window, RowContainer(window, 3)).Y, precision: 1);

            ReleaseAt(window, centres[2] + new Point(0, 10));
            window.Close();
        });
    }

    [Fact]
    public void Releasing_saves_the_new_order_into_the_draft()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);

            PressAt(window, CentreOf(window, HandleOf(window, 0)));
            DragTo(window, centres[2] + new Point(0, 5));
            ReleaseAt(window, centres[2] + new Point(0, 5));

            Assert.Equal(["PCB; Bottom", "Address Multiplexing", "PCB; Top", "Clock", "CPU 8502"], DraftOrder());
            Assert.Equal(DraftOrder(), Order(window));
            Assert.All(window.Schematics, row => Assert.False(row.IsDropPlaceholder));

            Assert.Equal("Moved [PCB; Top] to position 3 of 5.", window.StatusTextForTests);
            Assert.True(window.ChangedTheDraft);

            window.Close();
        });
    }

    [Fact]
    public void A_row_dragged_UP_lands_above_the_row_whose_middle_the_pointer_passed()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);

            PressAt(window, CentreOf(window, HandleOf(window, 4)));
            DragTo(window, centres[1] - new Point(0, 5));
            ReleaseAt(window, centres[1] - new Point(0, 5));

            Assert.Equal(["PCB; Top", "CPU 8502", "PCB; Bottom", "Address Multiplexing", "Clock"], DraftOrder());

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE OFF-BY-ONE, SEEN IN THE WINDOW. *** Nudged a few pixels down - past its own middle,
    // still over itself - a row must stay where it is. The worklog's first version already put it
    // below the next row here, so a downward drag ran one row ahead of the pointer.
    // ###########################################################################################
    [Fact]
    public void A_row_nudged_down_past_its_own_middle_stays_where_it_is()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);

            PressAt(window, HandleCentre(window, 0));
            DragTo(window, centres[0] + new Point(0, 12));

            Assert.True(window.Schematics[0].IsDropPlaceholder);
            Assert.Equal(FiveViews, Order(window));

            ReleaseAt(window, centres[0] + new Point(0, 12));
            window.Close();
        });
    }

    // A click is not a drag: nothing moves, and the draft is not even written.
    [Fact]
    public void A_click_without_moving_reorders_nothing_and_saves_nothing()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            DateTime before = DraftWrittenAt();
            Point handle = CentreOf(window, HandleOf(window, 1));

            PressAt(window, handle);
            DragTo(window, handle + new Point(1, 2));    // a shaky hand, under the threshold
            ReleaseAt(window, handle + new Point(1, 2));

            Assert.Equal(FiveViews, Order(window));
            Assert.All(window.Schematics, row => Assert.False(row.IsDropPlaceholder));
            Assert.Equal(before, DraftWrittenAt());
            Assert.False(window.ChangedTheDraft);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE REPORTED FLICKER - "it rapidly switches position in the UI and cannot settle". ***
    // Several moves arrive with NO layout pass between them, jiggling by a pixel the way a hand
    // does, and then more after a redraw. Against the live layout the row moved back and forth on
    // each; against the frozen slots it lands once and stays.
    // ###########################################################################################
    [Fact]
    public void Holding_a_dragged_row_still_does_not_make_the_rows_flicker()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);
            Point resting = centres[3] + new Point(0, 5);

            PressAt(window, CentreOf(window, HandleOf(window, 0)));

            for (int i = 0; i < 8; i++)
            {
                window.MouseMove(resting + new Point(0, i % 3), Avalonia.Input.RawInputModifiers.LeftMouseButton);
            }

            Dispatcher.UIThread.RunJobs();
            List<string> settled = Order(window);
            Assert.Equal(["PCB; Bottom", "Address Multiplexing", "Clock", "PCB; Top", "CPU 8502"], settled);

            for (int i = 0; i < 8; i++)
            {
                DragTo(window, resting + new Point(0, i % 3));
                Assert.Equal(settled, Order(window));
            }

            ReleaseAt(window, resting);
            window.Close();
        });
    }

    [Fact]
    public void Dragging_a_row_back_to_where_it_started_saves_nothing()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);
            DateTime before = DraftWrittenAt();

            PressAt(window, HandleCentre(window, 1));
            DragTo(window, centres[3] + new Point(0, 5));
            DragTo(window, centres[1]);
            ReleaseAt(window, centres[1]);

            Assert.Equal(FiveViews, Order(window));
            Assert.Equal(before, DraftWrittenAt());

            window.Close();
        });
    }

    // Flung past either end of the list, a row lands AT that end rather than being thrown away.
    [Theory]
    [InlineData(-400.0, new[] { "Address Multiplexing", "PCB; Top", "PCB; Bottom", "Clock", "CPU 8502" })]
    [InlineData(2000.0, new[] { "PCB; Top", "PCB; Bottom", "Clock", "CPU 8502", "Address Multiplexing" })]
    public void A_row_dragged_past_either_end_lands_at_that_end(double pointerY, string[] expected)
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point start = CentreOf(window, HandleOf(window, 2));

            PressAt(window, start);
            DragTo(window, new Point(start.X, pointerY));
            ReleaseAt(window, new Point(start.X, pointerY));

            Assert.Equal(expected, DraftOrder());

            window.Close();
        });
    }

    // ###########################################################################################
    // A drag that ends WITHOUT a release - the system taking the pointer capture away, as when the
    // window loses focus - saves nothing and puts every row back. What is on screen must never
    // disagree with the draft.
    // ###########################################################################################
    [Fact]
    public void A_drag_that_loses_the_pointer_puts_the_rows_back_and_saves_nothing()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);
            DateTime before = DraftWrittenAt();

            PressAt(window, HandleCentre(window, 0));
            DragTo(window, centres[3] + new Point(0, 5));
            Assert.NotEqual(FiveViews, Order(window));

            window.LoseSchematicDragCaptureForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(FiveViews, Order(window));
            Assert.All(window.Schematics, row => Assert.False(row.IsDropPlaceholder));
            Assert.Equal(before, DraftWrittenAt());

            // The release that follows finds no drag to finish.
            ReleaseAt(window, centres[3]);
            Assert.Equal(FiveViews, Order(window));
            Assert.Equal(before, DraftWrittenAt());

            window.Close();
        });
    }

    // ###########################################################################################
    // A board has a couple of dozen views and the list scrolls. Held near the bottom of the list's
    // viewport, the drag scrolls the list, and the placeholder keeps following the pointer down
    // into rows that were out of sight when the drag began.
    // ###########################################################################################
    [Fact]
    public void Holding_a_row_near_the_bottom_edge_scrolls_the_list_and_the_row_follows_down()
    {
        string[] views = Enumerable.Range(1, 14).Select(i => $"View {i:00}").ToArray();
        this.SeedBoardWithSchematics(views);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard(height: 480);
            ScrollViewer viewer = window.GetControl<ScrollViewer>("ContentScrollViewer");

            Assert.True(viewer.Extent.Height > viewer.Viewport.Height, "The fixture must scroll.");

            Point nearBottom = viewer.TranslatePoint(new Point(viewer.Bounds.Width / 2, viewer.Viewport.Height - 4), window)!.Value;

            PressAt(window, CentreOf(window, HandleOf(window, 0)));
            DragTo(window, nearBottom);

            BoardSchematicRow dragged = window.Schematics.Single(row => row.SchematicName == "View 01");
            int beforeScrolling = window.Schematics.IndexOf(dragged);

            for (int tick = 0; tick < 40; tick++)
            {
                window.AutoScrollSchematicDragForTests();
                Dispatcher.UIThread.RunJobs();
            }

            Assert.True(viewer.Offset.Y > 0, "The list should have scrolled down.");
            Assert.True(window.Schematics.IndexOf(dragged) > beforeScrolling, "The row should have followed the pointer down.");

            ReleaseAt(window, nearBottom);

            Assert.Equal(Order(window), DraftOrder());
            Assert.NotEqual("View 01", DraftOrder()[0]);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE CAPTION SURVIVES EVERY CHANGE IN THIS WINDOW (2026-09-27). *** Its copy of the board
    // dropped HardwareName and BoardName, so removing (or importing, or now reordering) an image
    // wrote the draft back without its "# Hardware:" / "# Board:" lines. The Remove case fails
    // against the hand-written copy that was there before.
    // ###########################################################################################
    [Fact]
    public void Removing_an_image_keeps_the_boards_hardware_and_board_caption()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            Button remove = RowContainer(window, 1).GetVisualDescendants().OfType<Button>().Single();
            remove.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            WaitForChanges(window);

            BoardData draft = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!;
            Assert.Equal(4, draft.Schematics.Count);
            Assert.Equal("C64", draft.HardwareName);
            Assert.Equal("250407", draft.BoardName);
            Assert.True(window.ChangedTheDraft);

            window.Close();
        });
    }

    [Fact]
    public void Reordering_keeps_the_boards_hardware_and_board_caption()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            Point[] centres = RowCentres(window);

            PressAt(window, HandleCentre(window, 0));
            DragTo(window, centres[2] + new Point(0, 5));
            ReleaseAt(window, centres[2] + new Point(0, 5));

            BoardData draft = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!;
            Assert.Equal("C64", draft.HardwareName);
            Assert.Equal("250407", draft.BoardName);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE COMPONENTS LEFT TO LABEL ARE INDIANRED (owner request, 2026-10-10: "3 components are
    // not labelled yet - please have them highlighted with the (standard?) IndianRed color"). ***
    // The to-do list's badges carry the Todo class, the other lists' do not, and the class's
    // colours are CRT's IndianRed in both themes.
    // ###########################################################################################
    [Fact]
    public void Only_the_components_left_to_label_are_drawn_in_IndianRed()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8" },
                unmatched: new[] { "PLA/U17" },
                unused: new[] { "LOGO3", "N1", "NT1" }));

            window.Show();

            try
            {
                List<Border> BadgesOf(string list) => window.GetControl<ItemsControl>(list)
                    .GetVisualDescendants().OfType<Border>()
                    .Where(border => border.Classes.Contains("ReferenceBadge"))
                    .ToList();

                Assert.Equal(3, BadgesOf("UnusedList").Count);
                Assert.All(BadgesOf("UnusedList"), badge => Assert.Contains("Todo", badge.Classes));
                Assert.All(BadgesOf("MatchedList").Concat(BadgesOf("UnmatchedList")), badge => Assert.DoesNotContain("Todo", badge.Classes));

                foreach (Avalonia.Styling.ThemeVariant theme in new[] { Avalonia.Styling.ThemeVariant.Light, Avalonia.Styling.ThemeVariant.Dark })
                {
                    Assert.True(Application.Current!.TryGetResource("Draft_ReferenceBadge_Todo_Bg", theme, out object? brush));
                    Assert.Equal(Avalonia.Media.Colors.IndianRed, ((Avalonia.Media.ISolidColorBrush)brush!).Color);
                }
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The CAD names are picked in the "Schematic images" window now - the list says where.
    [Fact]
    public void The_CAD_names_list_says_they_are_picked_in_the_Schematic_images_window()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8" },
                unmatched: Array.Empty<string>(),
                unused: Array.Empty<string>(),
                views: new[] { "Open128 - PCB Top" }));

            string explanation = window.GetControl<StackPanel>("ViewNamesPanel").Children.OfType<TextBlock>().ElementAt(1).Text!;

            Assert.Contains("\"Schematic images\"", explanation);
            Assert.Contains("CAD name", explanation);
        });
    }

    // ------------------------------------------------------------------ the KiCad NET NAMES
    //
    // Owner request, 2026-10-10: show in the window what the logfile listed as "normalized KiCad
    // net name [...]", each one copyable for the "Important signals" sheet. Which row finds which
    // net is KiCadNetNames' (CRT.Data.Tests); these cover how the window shows it.

    private static KiCadImportantSignalEntry Signal(string displayName, string netName) =>
        new() { DisplayName = displayName, KiCadNetName = netName };

    private static List<KiCadNetNameBadge> NetBadgesOf(BoardFilesWindow window) =>
        (window.GetControl<ItemsControl>("NetNamesList").ItemsSource ?? Array.Empty<KiCadNetNameBadge>())
            .Cast<KiCadNetNameBadge>()
            .ToList();

    [Fact]
    public void Every_net_name_is_listed_counted_and_the_used_ones_are_marked()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowNetNameReport(KiCadNetNames.Build(
                [("/A10", "A10", false), ("/A1", "A1", false), ("+5V", "+5V", false)],
                [Signal("Address 1", "/A1")]));

            Assert.True(VisibleOf(window, "NetNamesPanel"));
            Assert.Equal("3 KiCad net names you can use", TextOf(window, "NetNamesHeaderText"));
            Assert.Equal(["+5V", "A1", "A10"], NetBadgesOf(window).Select(badge => badge.Name));
            Assert.Equal(["A1"], NetBadgesOf(window).Where(badge => badge.IsUsed).Select(badge => badge.Name));
            Assert.Contains("Click a name to copy it", TextOf(window, "NetNamesIntroText"));
            Assert.Contains("\"Important signals\"", TextOf(window, "NetNamesIntroText"));
            Assert.Contains("The one in green is already used", TextOf(window, "NetNamesIntroText"));

            Assert.False(VisibleOf(window, "UnresolvedSignalsPanel"));
        });
    }

    // The same silent failure as a component that will not light up: the panel drops the row.
    [Fact]
    public void An_important_signal_naming_no_net_is_shown_and_counted()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowNetNameReport(KiCadNetNames.Build(
                [("PHI2", "PHI2", false)],
                [Signal("PHI2", "PH12"), Signal("Reset", "RES")]));

            Assert.True(VisibleOf(window, "UnresolvedSignalsPanel"));
            Assert.Equal("2 important signals will NOT light up", TextOf(window, "UnresolvedSignalsHeaderText"));
            Assert.Equal(["PHI2: PH12", "Reset: RES"], ItemsOf(window, "UnresolvedSignalsList"));
        });
    }

    [Fact]
    public void A_project_with_no_nets_shows_no_net_names()
    {
        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();

            window.ShowNetNameReport(KiCadNetNames.Build([], [Signal("Clock", "CLK")]));

            Assert.False(VisibleOf(window, "NetNamesPanel"));
        });
    }

    [Fact]
    public async Task Clicking_a_net_name_copies_it_to_the_clipboard_and_says_so()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow();
            window.Show();

            try
            {
                await window.CopyNetNameAsync("128_64");

                Avalonia.Input.Platform.IClipboard clipboard = TopLevel.GetTopLevel(window)!.Clipboard!;
                Assert.Equal("128_64", await clipboard.TryGetTextAsync());
                Assert.Contains("Copied [128_64]", window.StatusTextForTests);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The real path: the PCB's nets read from the draft's KiCad data, the board's "Important
    // signals" rows handed in by the Drafts tab.
    [Fact]
    public async Task Opening_on_KiCad_data_lists_its_nets_against_the_boards_important_signals()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string root = Path.Combine(DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey), "KiCad data");
        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, "board.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow();

            window.Initialize(
                "C64 - 250407",
                BoardFilesWindowTests.BoardKey,
                new[] { "U1" },
                BoardFilesSection.KiCadData,
                [Signal("Ground", "GND"), Signal("Clock", "CLK")]);

            for (int i = 0; i < 100 && !VisibleOf(window, "NetNamesPanel"); i++)
            {
                await Task.Delay(20);
                Dispatcher.UIThread.RunJobs();
            }

            // The fixture's (net 0 "") - KiCad's "no net" - is not listed (owner decision, 2026-10-10).
            Assert.Equal(["GND"], NetBadgesOf(window).Select(badge => badge.Name));
            Assert.True(Assert.Single(NetBadgesOf(window)).IsUsed);
            Assert.Equal(["Clock: CLK"], ItemsOf(window, "UnresolvedSignalsList"));
        });
    }

    // ------------------------------------------------------------------ picking each image's CAD NAME
    //
    // Owner request, 2026-10-10: "In the "Schematic images" then have that mapping available, so I
    // can select a KiCad name. Once selected, it cannot be selected again." Which names an image may
    // pick is SchematicCadNameChoices' (CRT.Data.Tests); these cover the window - reading the names,
    // keeping each drop-down's selection, saving a pick into the draft.

    // The draft's KiCad data: one PCB (two views named after the file) and a root sheet naming its
    // child "Power supply".
    private void WriteCadNameKiCadData()
    {
        string root = Path.Combine(DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey), "KiCad data");

        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, "Open128.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);
        File.WriteAllText(Path.Combine(root, "root.kicad_sch"), KiCadFixtures.SchematicRootWithChildSheet);
        File.WriteAllText(Path.Combine(root, "child.kicad_sch"), KiCadFixtures.SchematicChildSheet);
    }

    private static readonly string[] CadViews = ["Open128 - PCB Top", "Open128 - PCB Bottom", "root", "Power supply"];

    // Gives one seeded image a CAD name, as the draft would hold it. Seeded images are named after
    // their files (SeedBoardWithSchematics).
    private static void SetDraftCadName(string schematicName, string cadName) =>
        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            BoardFilesWindowTests.BoardKey,
            board => SchematicRows.WithCadName(
                board,
                new SchematicRowRef(schematicName, BoardFilesWindowTests.StoredImagePath(schematicName + ".png")),
                cadName).Board);

    private static string DraftCadName(string schematicName) =>
        DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!
            .Schematics.Single(entry => entry.SchematicName == schematicName).CadName;

    private static ComboBox PickerOf(BoardFilesWindow window, int index) =>
        RowContainer(window, index).GetVisualDescendants().OfType<ComboBox>()
            .Single(combo => combo.Classes.Contains("CadNamePicker"));

    // A name another image has is not offered at all (owner request, 2026-10-10: "remove the CAD
    // name entirely from the drop-down when it has been selected once"). It replaces the first
    // version's greyed-out "(used by ...)" entry, changed on purpose.
    [Fact]
    public async Task Every_image_is_offered_the_KiCad_view_names_but_not_one_another_image_has()
    {
        this.SeedBoardWithSchematics("Top scan", "Bottom scan");
        this.WriteCadNameKiCadData();
        SetDraftCadName("Top scan", "Open128 - PCB Top");

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);
            await window.CadNamesLoading;

            BoardSchematicRow top = window.Schematics[0];
            BoardSchematicRow bottom = window.Schematics[1];

            Assert.Equal(["", .. CadViews], top.CadNameChoices.Select(choice => choice.CadName));
            Assert.Equal("Open128 - PCB Top", top.SelectedCadNameChoice!.CadName);
            Assert.Equal(string.Empty, bottom.SelectedCadNameChoice!.CadName);

            Assert.Equal(["", .. CadViews.Skip(1)], bottom.CadNameChoices.Select(choice => choice.CadName));

            Assert.False(window.GetControl<TextBlock>("CadNamesStateText").IsVisible);
        });
    }

    // ###########################################################################################
    // *** THE SELECTION SURVIVES THE NAMES ARRIVING. *** The rows are on screen before the KiCad
    // data is read, each drop-down holding only "No CAD name" and the image's own name. Had the
    // names arrived as a NEW list, every ComboBox would have let go of its selection - and its
    // binding written "nothing" back, wiping each image's CAD name. Read off the SHOWN drop-down.
    // ###########################################################################################
    [Fact]
    public async Task A_shown_drop_down_keeps_its_images_CAD_name_when_the_KiCad_names_arrive()
    {
        this.SeedBoardWithSchematics("Top scan", "Bottom scan");
        this.WriteCadNameKiCadData();
        SetDraftCadName("Bottom scan", "Power supply");

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            try
            {
                CadNameChoiceItem before = (CadNameChoiceItem)PickerOf(window, 1).SelectedItem!;
                Assert.Equal("Power supply", before.CadName);

                await window.CadNamesLoading;
                Dispatcher.UIThread.RunJobs();

                ComboBox picker = PickerOf(window, 1);
                Assert.Same(before, picker.SelectedItem);
                Assert.Equal(CadViews.Length + 1, picker.ItemCount);

                Assert.Equal("Power supply", DraftCadName("Bottom scan"));
                Assert.Equal(string.Empty, DraftCadName("Top scan"));
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public async Task Picking_a_CAD_name_saves_it_into_the_draft_and_takes_it_from_the_other_images()
    {
        this.SeedBoardWithSchematics("Top scan", "Bottom scan");
        this.WriteCadNameKiCadData();

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            try
            {
                await window.CadNamesLoading;
                Dispatcher.UIThread.RunJobs();

                ComboBox picker = PickerOf(window, 0);
                picker.SelectedItem = window.Schematics[0].CadNameChoices.Single(choice => choice.CadName == "Open128 - PCB Bottom");
                WaitForChanges(window);

                Assert.Equal("Open128 - PCB Bottom", DraftCadName("Top scan"));
                Assert.Equal("Open128 - PCB Bottom", window.Schematics[0].CadName);
                Assert.Contains("[Top scan] now has the CAD name [Open128 - PCB Bottom]", window.StatusTextForTests);
                Assert.True(window.ChangedTheDraft);

                // Gone from the other image's list - the shown drop-down too.
                Assert.DoesNotContain(window.Schematics[1].CadNameChoices, choice => choice.CadName == "Open128 - PCB Bottom");
                Assert.Equal(CadViews.Length, PickerOf(window, 1).ItemCount);

                // Saved in place: the rows were not read again, so the drop-down still holds its pick.
                Assert.Equal("Open128 - PCB Bottom", ((CadNameChoiceItem)PickerOf(window, 0).SelectedItem!).CadName);
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public async Task Picking_No_CAD_name_clears_it_and_frees_the_name_for_the_other_images()
    {
        this.SeedBoardWithSchematics("Top scan", "Bottom scan");
        this.WriteCadNameKiCadData();
        SetDraftCadName("Top scan", "root");

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            try
            {
                await window.CadNamesLoading;
                Dispatcher.UIThread.RunJobs();

                Assert.DoesNotContain(window.Schematics[1].CadNameChoices, choice => choice.CadName == "root");

                PickerOf(window, 0).SelectedItem = window.Schematics[0].CadNameChoices[0];
                WaitForChanges(window);

                Assert.Equal(string.Empty, DraftCadName("Top scan"));
                // Back in the other image's list, in its place in the project's order.
                Assert.Equal(["", .. CadViews], window.Schematics[1].CadNameChoices.Select(choice => choice.CadName));
                Assert.Contains("[Top scan] has no CAD name now.", window.StatusTextForTests);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // *** TAKING A NAME OUT OF A SHOWN DROP-DOWN LEAVES ITS OWN PICK ALONE. *** Another image picks
    // "root", which sits ABOVE this image's "Power supply" in the list: it is removed in place, and
    // this drop-down keeps showing the very entry it showed - nothing is written for this image.
    // ###########################################################################################
    [Fact]
    public async Task A_name_taken_out_of_a_shown_drop_down_leaves_its_own_pick_alone()
    {
        this.SeedBoardWithSchematics("Top scan", "Bottom scan");
        this.WriteCadNameKiCadData();
        SetDraftCadName("Bottom scan", "Power supply");

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            try
            {
                await window.CadNamesLoading;
                Dispatcher.UIThread.RunJobs();

                object before = PickerOf(window, 1).SelectedItem!;

                PickerOf(window, 0).SelectedItem = window.Schematics[0].CadNameChoices.Single(choice => choice.CadName == "root");
                WaitForChanges(window);

                Assert.Equal("root", DraftCadName("Top scan"));
                Assert.Equal(
                    ["", "Open128 - PCB Top", "Open128 - PCB Bottom", "Power supply"],
                    window.Schematics[1].CadNameChoices.Select(choice => choice.CadName));
                Assert.Same(before, PickerOf(window, 1).SelectedItem);
                Assert.Equal("Power supply", DraftCadName("Bottom scan"));
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public async Task A_board_with_no_KiCad_data_says_there_is_no_CAD_name_to_pick()
    {
        this.SeedBoardWithSchematics("Top scan");

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.SchematicImages);
            await window.CadNamesLoading;

            TextBlock state = window.GetControl<TextBlock>("CadNamesStateText");
            Assert.True(state.IsVisible);
            Assert.Contains("no KiCad data yet", state.Text);

            Assert.Equal([""], window.Schematics[0].CadNameChoices.Select(choice => choice.CadName));
        });
    }

    // The drop-down sits OUTSIDE the drag handle, so opening it can never start a row drag.
    [Fact]
    public void The_CAD_name_drop_down_is_outside_the_drag_handle()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            for (int i = 0; i < FiveViews.Length; i++)
            {
                ComboBox picker = PickerOf(window, i);

                Assert.True(picker.IsEffectivelyVisible);
                Assert.DoesNotContain(picker, HandleOf(window, i).GetVisualDescendants());
            }

            window.Close();
        });
    }

    [Fact]
    public void Picking_a_CAD_name_keeps_the_rows_other_columns_the_other_rows_and_the_caption()
    {
        this.SeedBoardWithSchematics("Top scan", "Bottom scan");

        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            BoardFilesWindowTests.BoardKey,
            board =>
            {
                BoardSchematicEntry plain = board.Schematics[0];
                board.Schematics[0] = new BoardSchematicEntry
                {
                    SchematicName = plain.SchematicName,
                    SchematicImageFile = plain.SchematicImageFile,
                    SchematicHighlightColor = "#FF0000",
                    ThumbnailHighlightOpacity = "0.5",
                };

                return board;
            });

        SetDraftCadName("top SCAN", "Open128 - PCB Top");

        BoardData draft = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!;

        Assert.Equal("Open128 - PCB Top", draft.Schematics[0].CadName);
        Assert.Equal("#FF0000", draft.Schematics[0].SchematicHighlightColor);
        Assert.Equal("0.5", draft.Schematics[0].ThumbnailHighlightOpacity);
        Assert.Equal(string.Empty, draft.Schematics[1].CadName);
        Assert.Equal("C64", draft.HardwareName);
        Assert.Equal("250407", draft.BoardName);
    }

    // ###########################################################################################
    // A row's choices are changed IN PLACE (BoardSchematicRow.ApplyCadNameChoices): the entries
    // already there keep their objects, so a ComboBox showing one keeps showing it.
    // ###########################################################################################
    [Fact]
    public void A_rows_choices_are_updated_in_place_and_its_selected_entry_keeps_its_object()
    {
        UiTest.Run(() =>
        {
            var row = new BoardSchematicRow { SchematicName = "Top scan", CadName = "Reset" };

            row.ApplyCadNameChoices(SchematicCadNameChoices.For([("Top scan", "Reset")], 0, []));
            CadNameChoiceItem none = row.CadNameChoices[0];
            CadNameChoiceItem reset = row.SelectedCadNameChoice!;
            Assert.Equal("Reset", reset.CadName);

            // The names arrive: inserted around the selected entry, which keeps its object.
            row.ApplyCadNameChoices(SchematicCadNameChoices.For([("Top scan", "Reset"), ("Other", "")], 0, ["Clock", "Reset", "VIC"]));

            Assert.Equal(["", "Clock", "Reset", "VIC"], row.CadNameChoices.Select(choice => choice.CadName));
            Assert.Same(none, row.CadNameChoices[0]);
            Assert.Same(reset, row.CadNameChoices[2]);
            Assert.Same(reset, row.SelectedCadNameChoice);

            // Another image takes "Clock", ABOVE the selected entry: removed, the rest untouched.
            CadNameChoiceItem vic = row.CadNameChoices[3];
            row.ApplyCadNameChoices(SchematicCadNameChoices.For([("Top scan", "Reset"), ("Other", "Clock")], 0, ["Clock", "Reset", "VIC"]));

            Assert.Equal(["", "Reset", "VIC"], row.CadNameChoices.Select(choice => choice.CadName));
            Assert.Same(reset, row.SelectedCadNameChoice);
            Assert.Same(vic, row.CadNameChoices[2]);

            // The names going away again: only what is still wanted stays, the selection with it.
            row.ApplyCadNameChoices(SchematicCadNameChoices.For([("Top scan", "Reset")], 0, []));

            Assert.Equal(["", "Reset"], row.CadNameChoices.Select(choice => choice.CadName));
            Assert.Same(reset, row.SelectedCadNameChoice);
        });
    }

    // ------------------------------------------------------------------ RENAMING an image
    //
    // Owner request, 2026-10-10: "add management of the name of the schematic image also" - the
    // "Schematic name" column. Typed into the row's Name box, saved on Enter or on leaving the box.
    // What the rename carries (highlights, calibration) is DraftWorkbookStoreTests'.

    private static TextBox NameBoxOf(BoardFilesWindow window, int index) =>
        RowContainer(window, index).GetVisualDescendants().OfType<TextBox>()
            .Single(box => box.Classes.Contains("SchematicNameBox"));

    // The row's bold title - what the list shows the image as.
    private static string TitleOf(BoardFilesWindow window, int index) =>
        HandleOf(window, index).GetVisualDescendants().OfType<TextBlock>()
            .Single(text => text.FontWeight == Avalonia.Media.FontWeight.SemiBold).Text ?? string.Empty;

    // Empties a row's Name box and types `name` into it, as a contributor would.
    private static void TypeName(BoardFilesWindow window, int index, string name)
    {
        TextBox box = NameBoxOf(window, index);
        box.Focus();
        box.Text = string.Empty;

        if (name.Length > 0)
        {
            window.KeyTextInput(name);
        }

        Dispatcher.UIThread.RunJobs();
    }

    private static void Press(Window window, Avalonia.Input.Key key, Avalonia.Input.PhysicalKey physicalKey)
    {
        window.KeyPress(key, Avalonia.Input.RawInputModifiers.None, physicalKey, keySymbol: null);
        Dispatcher.UIThread.RunJobs();

        // Enter saves a name as one of the window's queued changes.
        if (window is BoardFilesWindow files)
        {
            WaitForChanges(files);
        }
    }

    [Fact]
    public void Typing_a_new_name_and_pressing_Enter_renames_the_image_in_the_draft()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 0, "PCB top side");
            Press(window, Avalonia.Input.Key.Enter, Avalonia.Input.PhysicalKey.Enter);

            Assert.Equal(["PCB top side", "PCB; Bottom", "Address Multiplexing", "Clock", "CPU 8502"], DraftOrder());
            Assert.Equal("PCB top side", window.Schematics[0].SchematicName);
            Assert.Equal("PCB top side", TitleOf(window, 0));
            Assert.Equal("Renamed [PCB; Top] to [PCB top side].", window.StatusTextForTests);
            Assert.True(window.ChangedTheDraft);

            // Enter saved the name; it did not close the window.
            Assert.True(window.IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void Leaving_the_Name_box_saves_the_name()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 1, "PCB bottom side");
            NameBoxOf(window, 2).Focus();
            WaitForChanges(window);

            Assert.Equal("PCB bottom side", DraftOrder()[1]);

            window.Close();
        });
    }

    // The name is the sheet's key: a second image of one name is refused, with the reason, and the
    // box shows the name the image still has.
    [Fact]
    public void A_name_another_image_has_is_refused_and_the_box_shows_the_name_again()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            DateTime before = DraftWrittenAt();

            TypeName(window, 0, "clock");
            Press(window, Avalonia.Input.Key.Enter, Avalonia.Input.PhysicalKey.Enter);

            Assert.Equal(SchematicRename.NameTaken("clock"), window.StatusTextForTests);
            Assert.Equal("PCB; Top", NameBoxOf(window, 0).Text);
            Assert.Equal(FiveViews, DraftOrder());
            Assert.Equal(before, DraftWrittenAt());
            Assert.False(window.ChangedTheDraft);

            window.Close();
        });
    }

    [Fact]
    public void A_blank_name_is_refused()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 3, string.Empty);
            Press(window, Avalonia.Input.Key.Enter, Avalonia.Input.PhysicalKey.Enter);

            Assert.Equal(SchematicRename.NeedsAName, window.StatusTextForTests);
            Assert.Equal("Clock", NameBoxOf(window, 3).Text);
            Assert.Equal(FiveViews, DraftOrder());

            window.Close();
        });
    }

    // Escape in a box holding an unsaved name puts the name back - it does not close the window,
    // which Escape anywhere else in it does.
    [Fact]
    public void Escape_puts_the_name_back_without_closing_the_window()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 0, "Not this");
            Press(window, Avalonia.Input.Key.Escape, Avalonia.Input.PhysicalKey.Escape);

            Assert.True(window.IsVisible);
            Assert.Equal("PCB; Top", NameBoxOf(window, 0).Text);
            Assert.Equal(FiveViews, DraftOrder());

            // With nothing left to put back, Escape closes the window as it always did.
            Press(window, Avalonia.Input.Key.Escape, Avalonia.Input.PhysicalKey.Escape);
            Assert.False(window.IsVisible);

            // And the name put back was not saved on the way out.
            Assert.Equal(FiveViews, DraftOrder());
        });
    }

    // Closed by its title bar, which takes no focus from the box: the typed name is still saved,
    // before the Drafts tab looks at ChangedTheDraft.
    [Fact]
    public void Closing_the_window_with_a_name_still_typed_saves_it()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 4, "CPU 8502 and 6510");
            window.Close();
            WaitForChanges(window);

            Assert.Equal("CPU 8502 and 6510", DraftOrder()[4]);
            Assert.True(window.ChangedTheDraft);
            Assert.False(window.IsVisible);
        });
    }

    // Every other save in the window finds its row by its name, so after a rename they must use the
    // new one.
    [Fact]
    public void A_CAD_name_picked_after_a_rename_lands_on_the_renamed_image()
    {
        this.SeedBoardWithSchematics("Top scan", "Bottom scan");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 0, "Top side");
            Press(window, Avalonia.Input.Key.Enter, Avalonia.Input.PhysicalKey.Enter);

            _ = window.SaveCadNameAsync(window.Schematics[0], "Open128 - PCB Top");
            WaitForChanges(window);

            Assert.Equal("Open128 - PCB Top", DraftCadName("Top side"));

            window.Close();
        });
    }

    // The Name box sits OUTSIDE the drag handle, like the CAD name, so pressing in it to type never
    // starts a drag.
    [Fact]
    public void The_Name_box_is_outside_the_drag_handle()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            for (int i = 0; i < FiveViews.Length; i++)
            {
                TextBox box = NameBoxOf(window, i);

                Assert.True(box.IsEffectivelyVisible);
                Assert.Equal(FiveViews[i], box.Text);
                Assert.DoesNotContain(box, HandleOf(window, i).GetVisualDescendants());
            }

            window.Close();
        });
    }

    // ------------------------------------------------------------------ the code review of 2026-10-10
    //
    // Every change the window makes to the draft goes to the row it was made on, is written under
    // the "please wait", one at a time, and closing waits for them (BoardFilesWindow.DraftChanges.cs).

    // A drafted board with these images, each named and filed as given - so two can share a name.
    private void SeedBoardWithImages(params (string Name, string File)[] images)
    {
        DraftSeeder.CreateNewBoard(
            DraftManager.DraftsRoot,
            new NewBoardRegistration
            {
                HardwareName = "C64",
                BoardName = "250407",
                ExcelDataFile = BoardFilesWindowTests.BoardKey,
            });

        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            BoardFilesWindowTests.BoardKey,
            board =>
            {
                foreach ((string name, string file) in images)
                {
                    board.Schematics.Add(new BoardSchematicEntry
                    {
                        SchematicName = name,
                        SchematicImageFile = BoardFilesWindowTests.StoredImagePath(file),
                    });
                }

                return board;
            });
    }

    private static List<string> DraftFiles() =>
        DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!
            .Schematics.Select(entry => entry.SchematicImageFile).ToList();

    // Highlight rectangles drawn on one image of the seeded draft, one per label.
    private static void AddHighlights(string schematicName, params string[] labels) =>
        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            BoardFilesWindowTests.BoardKey,
            board =>
            {
                board.ComponentHighlights.AddRange(labels.Select(label => new ComponentHighlightEntry
                {
                    SchematicName = schematicName,
                    BoardLabel = label,
                    X = "1",
                    Y = "2",
                    Width = "3",
                    Height = "4",
                }));

                return board;
            });

    // A one-pixel image outside the draft, to add to it.
    private string WriteScan(string fileName)
    {
        string scan = Path.Combine(this.thisWorkspace.Root, "Scans", fileName);

        Directory.CreateDirectory(Path.GetDirectoryName(scan)!);
        File.WriteAllBytes(scan, BoardFilesWindowTests.OnePixelPng);

        return scan;
    }

    // ###########################################################################################
    // *** OF TWO IMAGES SHARING A NAME, A CHANGE GOES TO THE ONE IT WAS MADE ON. *** Every save
    // found its row by the name alone: renaming the second "Top" renamed the first, and a CAD name
    // picked for one could land on the other. The row is found by its image file too now.
    // ###########################################################################################
    [Fact]
    public void A_change_to_one_of_two_images_sharing_a_name_lands_on_that_image()
    {
        this.SeedBoardWithImages(("Top", "top-a.png"), ("Top", "top-b.png"), ("Clock", "clock.png"));

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 1, "Top (second scan)");
            Press(window, Avalonia.Input.Key.Enter, Avalonia.Input.PhysicalKey.Enter);

            Assert.Equal(["Top", "Top (second scan)", "Clock"], DraftOrder());
            Assert.Equal(BoardFilesWindowTests.StoredImagePath("top-b.png"), DraftFiles()[1]);
            Assert.Equal("Top (second scan)", window.Schematics[1].SchematicName);

            _ = window.SaveCadNameAsync(window.Schematics[1], "Open128 - PCB Bottom");
            WaitForChanges(window);

            Assert.Equal("Open128 - PCB Bottom", DraftCadName("Top (second scan)"));
            Assert.Equal(string.Empty, DraftCadName("Top"));

            window.Close();
        });
    }

    // Remove took EVERY row of the name.
    [Fact]
    public void Removing_one_of_two_images_sharing_a_name_removes_only_that_one()
    {
        this.SeedBoardWithImages(("Top", "top-a.png"), ("Top", "top-b.png"));

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            _ = window.RemoveSchematicAsync(window.Schematics[1]);
            WaitForChanges(window);

            Assert.Equal([BoardFilesWindowTests.StoredImagePath("top-a.png")], DraftFiles());
            Assert.Equal("Removed [Top]. The image file is kept in your draft folder.", window.StatusTextForTests);
            Assert.True(window.ChangedTheDraft);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** REMOVE ASKS FIRST WHEN THE IMAGE HAS SOMETHING DRAWN ON IT. *** Its highlights (and its
    // KiCad calibration) go with the row, and adding the image again brings none of them back - it
    // used to go without a word. A "no" leaves the draft as it was.
    // ###########################################################################################
    [Fact]
    public void Removing_an_image_with_highlights_on_it_asks_first_and_a_no_keeps_everything()
    {
        this.SeedBoardWithSchematics(FiveViews);
        AddHighlights("PCB; Top", "U1", "U2");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            DateTime before = DraftWrittenAt();
            var asked = new List<(string Name, SchematicRemovalLoss Loss)>();

            window.ConfirmRemovalOverrideForTests = (name, loss) =>
            {
                asked.Add((name, loss));
                return Task.FromResult(false);
            };

            _ = window.RemoveSchematicAsync(window.Schematics[0]);
            WaitForChanges(window);

            Assert.Equal([("PCB; Top", new SchematicRemovalLoss(2, false))], asked);
            Assert.Equal(FiveViews, DraftOrder());
            Assert.Equal(before, DraftWrittenAt());
            Assert.False(window.ChangedTheDraft);

            window.ConfirmRemovalOverrideForTests = (_, _) => Task.FromResult(true);

            _ = window.RemoveSchematicAsync(window.Schematics[0]);
            WaitForChanges(window);

            Assert.Equal(FiveViews.Skip(1), DraftOrder());
            Assert.DoesNotContain(
                DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!.ComponentHighlights,
                highlight => highlight.SchematicName == "PCB; Top");
            Assert.True(window.ChangedTheDraft);

            window.Close();
        });
    }

    // Nothing drawn on it, nothing lost: removed at once, as it always was.
    [Fact]
    public void Removing_an_image_with_nothing_drawn_on_it_does_not_ask()
    {
        this.SeedBoardWithSchematics(FiveViews);
        AddHighlights("PCB; Top", "U1");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            bool asked = false;

            window.ConfirmRemovalOverrideForTests = (_, _) =>
            {
                asked = true;
                return Task.FromResult(true);
            };

            _ = window.RemoveSchematicAsync(window.Schematics[3]);
            WaitForChanges(window);

            Assert.False(asked);
            Assert.Equal(["PCB; Top", "PCB; Bottom", "Address Multiplexing", "CPU 8502"], DraftOrder());

            window.Close();
        });
    }

    // A disposable that does nothing - a reveal the test never runs.
    private sealed class NoReveal : IDisposable
    {
        public void Dispose()
        {
        }
    }

    // ###########################################################################################
    // *** A RENAME IS WRITTEN UNDER THE "PLEASE WAIT". *** It rewrote the whole workbook on the UI
    // thread - up to a couple of seconds on the largest board - with nothing on screen, against the
    // rule that every wait the user watches runs under BusyOverlay. Every wait that starts is seen
    // here by the sentence it starts with.
    // ###########################################################################################
    [Fact]
    public void A_rename_is_written_under_the_windows_please_wait()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            BusyOverlay overlay = BusyOverlay.For(window)!;
            var started = new List<string>();

            overlay.ScheduleRevealOverrideForTests = (_, _) =>
            {
                started.Add(overlay.Message);
                return new NoReveal();
            };

            TypeName(window, 0, "PCB top side");
            Press(window, Avalonia.Input.Key.Enter, Avalonia.Input.PhysicalKey.Enter);

            Assert.Equal([CrtWaitWording.SavingSchematicName], started);
            Assert.False(overlay.IsBusy);
            Assert.Equal("PCB top side", DraftOrder()[0]);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THROUGH THE DRAFT HELD IN MEMORY (CRT 4.0 step 1, 2026-10-25). *** Each change the window
    // makes is one step of the draft's history, named for what it did - so the Drafts tab's table,
    // open behind the window, shows it at once, and its Ctrl+Z takes it back.
    // ###########################################################################################
    [Fact]
    public void Each_change_the_window_makes_is_one_step_of_the_drafts_history()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 0, "PCB top side");
            Press(window, Avalonia.Input.Key.Enter, Avalonia.Input.PhysicalKey.Enter);
            WaitForChanges(window);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey)!;
            Assert.Equal(DraftStepWording.SchematicRenamed("PCB; Top", "PCB top side"), session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);

            Assert.True(session.Undo());
            Assert.Equal("PCB; Top", session.Content.Board.Schematics[0].SchematicName);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A NAME THAT CANNOT BE SAVED KEEPS THE WINDOW OPEN, WITH THE REASON. *** Closed with a name
    // another image has, the window used to close and drop the name - its reason shown on a status
    // line nobody saw. Closed again, with nothing left to save, it goes.
    // ###########################################################################################
    [Fact]
    public void Closing_with_a_name_that_cannot_be_saved_keeps_the_window_open_and_says_why()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 0, "clock");
            window.Close();
            WaitForChanges(window);

            Assert.True(window.IsVisible);
            Assert.Equal(SchematicRename.NameTaken("clock"), window.StatusTextForTests);
            Assert.Equal("PCB; Top", NameBoxOf(window, 0).Text);
            Assert.Equal(FiveViews, DraftOrder());

            window.Close();
            WaitForChanges(window);

            Assert.False(window.IsVisible);
        });
    }

    // ###########################################################################################
    // *** A NAME THE DRAFT COULD NOT BE WRITTEN WITH STAYS IN ITS BOX (code review, 2026-10-11). ***
    // The name was fine - the file could not be written, as when Excel holds it - yet the box showed
    // the old name again: after closing Excel the name had to be typed again, and closing the window
    // a second time went through without it. Now it stays, and the next close saves it. The write is
    // stopped by holding the draft's board file open (CRT 4.0's one file), which Windows enforces.
    // ###########################################################################################
    [Fact]
    public void A_name_that_could_not_be_written_stays_in_its_box_and_is_saved_once_the_file_is_free()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey);

            // Readable, but not replaceable - as when another program holds the file.
            using (new FileStream(boardFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                TypeName(window, 0, "PCB top side");
                window.Close();
                WaitForChanges(window);
            }

            Assert.True(window.IsVisible);
            Assert.Equal(BoardFilesWindow.NameNotWritten, window.StatusTextForTests);
            Assert.Equal("PCB top side", NameBoxOf(window, 0).Text);
            Assert.Equal(FiveViews, DraftOrder());

            window.Close();
            WaitForChanges(window);

            Assert.False(window.IsVisible);
            Assert.Equal("PCB top side", DraftOrder()[0]);
        });
    }

    // ###########################################################################################
    // *** A NAME STILL BEING TYPED SURVIVES THE LIST BEING READ AGAIN. *** Images dropped in while a
    // Name box held a name replaced every row, and the typed name went with its old row. It stays in
    // the box of the same image, unsaved - and is saved the way a typed name always is.
    // ###########################################################################################
    [Fact]
    public void A_name_still_being_typed_when_images_are_added_stays_in_its_box()
    {
        this.SeedBoardWithSchematics(FiveViews);
        string scan = this.WriteScan("Keyboard.png");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            TypeName(window, 1, "PCB bottom side");
            _ = window.ImportSchematicsAsync([scan]);
            WaitForChanges(window);

            Assert.Equal(6, window.Schematics.Count);
            Assert.Equal("PCB bottom side", window.Schematics[1].NameDraft);
            Assert.Equal("PCB; Bottom", DraftOrder()[1]);

            window.Close();
            WaitForChanges(window);

            Assert.Equal("PCB bottom side", DraftOrder()[1]);
            Assert.False(window.IsVisible);
        });
    }

    // ###########################################################################################
    // *** IMAGES THE DRAFT CANNOT TAKE ARE NOT "ADDED". *** The draft's answer was ignored: a
    // workbook Excel held open still read "Added 1 schematic image." with nothing new in the list,
    // and the copy stayed in the draft folder for no row. A workbook that cannot be read cannot be
    // written either - on every system, unlike a file another program holds.
    // ###########################################################################################
    [Fact]
    public void Images_the_draft_cannot_take_are_not_reported_as_added_and_their_copies_are_taken_back()
    {
        this.SeedBoardWithSchematics(FiveViews);
        string scan = this.WriteScan("Keyboard.png");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();

            File.WriteAllText(
                DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardFilesWindowTests.BoardKey),
                "not a workbook");

            _ = window.ImportSchematicsAsync([scan]);
            WaitForChanges(window);

            Assert.Contains("could not be added to your draft", window.StatusTextForTests);
            Assert.False(window.ChangedTheDraft);
            Assert.False(File.Exists(Path.Combine(DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey), "Keyboard.png")));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** ESCAPE ON AN OPEN CAD NAME LIST CLOSES THE LIST, NOT THE WINDOW. *** The window's Escape
    // runs on the tunnel route, before the drop-down sees the key, and closed the whole window.
    // ###########################################################################################
    [Fact]
    public void Escape_on_an_open_CAD_name_list_closes_the_list_and_not_the_window()
    {
        this.SeedBoardWithSchematics(FiveViews);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.ShowOnSeededBoard();
            ComboBox picker = PickerOf(window, 0);

            picker.Focus();
            picker.IsDropDownOpen = true;
            Dispatcher.UIThread.RunJobs();

            Press(window, Avalonia.Input.Key.Escape, Avalonia.Input.PhysicalKey.Escape);

            Assert.True(window.IsVisible);
            Assert.False(picker.IsDropDownOpen);

            // With no list open, Escape closes the window as it always did.
            Press(window, Avalonia.Input.Key.Escape, Avalonia.Input.PhysicalKey.Escape);

            Assert.False(window.IsVisible);
        });
    }

    // ###########################################################################################
    // *** A KiCad IMPORT THAT STOPS PART-WAY STILL CHANGED THE DRAFT. *** The files before the one
    // that failed are overwritten, so the board on screen has to be read again when the window
    // closes - it went on drawing the old project. A folder where the second file goes makes that
    // copy fail on every system.
    // ###########################################################################################
    [Fact]
    public async Task A_KiCad_import_that_stops_part_way_still_counts_as_a_change()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string source = Path.Combine(this.thisWorkspace.Root, "My KiCad project");
        Directory.CreateDirectory(source);
        File.WriteAllText(Path.Combine(source, "a.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);
        File.WriteAllText(Path.Combine(source, "b.kicad_sch"), "(kicad_sch (version 20230121))");

        Directory.CreateDirectory(Path.Combine(DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey), "KiCad data", "b.kicad_sch"));

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            await window.ImportKiCadFolderAsync(source);

            Assert.True(window.ChangedTheDraft);
            Assert.Contains("Only 1 of the 2 KiCad files could be copied", window.StatusTextForTests);
            Assert.Equal(["a.kicad_pcb"], window.KiCadFiles.Select(row => row.RelativePath));
        });
    }

    // ###########################################################################################
    // *** AN IMPORTED KiCad PROJECT IS LISTED IN THE DRAFT AT ONCE (code review, 2026-10-25). ***
    // Submit sends the KiCad files the draft's board file lists; left lying in "KiCad data" until
    // some other edit wrote the draft, a project imported last was drawn and counted but not sent.
    // ###########################################################################################
    [Fact]
    public async Task An_imported_KiCad_project_is_listed_in_the_drafts_board_file_at_once()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string source = Path.Combine(this.thisWorkspace.Root, "My KiCad project");
        Directory.CreateDirectory(Path.Combine(source, "Pages"));
        File.WriteAllText(Path.Combine(source, "a.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);
        File.WriteAllText(Path.Combine(source, "Pages", "b.kicad_sch"), "(kicad_sch (version 20230121))");

        await UiTest.RunAsync(async () =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard(BoardFilesSection.KiCadData);

            await window.ImportKiCadFolderAsync(source);
        });

        string draftFolder = DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey);
        BoardDocument? draft = BoardFiles.ReadDocument(Path.Combine(draftFolder, DataTreeLayout.BoardFileName));

        Assert.NotNull(draft);
        Assert.Equal(
            ["KiCad data/Pages/b.kicad_sch", "KiCad data/a.kicad_pcb"],
            draft.Files.Where(Handlers.DataHandling.DraftFiles.IsKiCad).Select(file => file.Name).Order(StringComparer.Ordinal));
        Assert.Empty(Handlers.DataHandling.DraftFiles.ArrivedKiCadNames(draftFolder));
    }

    // ###########################################################################################
    // *** THE KiCad REPORT READS THE PUBLISHED KiCad DATA TOO. *** Like the Schematics tab, which
    // draws the traces from the published board's files and the draft's. Read from the draft's
    // alone, a component or net only the published files have was reported as "will NOT light up".
    // The draft's own list still lists only the draft's files - those are what Remove takes.
    // ###########################################################################################
    // In CRT 4.0 (2026-10-24) the draft's board file names the published KiCad files it was copied
    // with - so the report reads them through the draft, and they are not the draft's imports.
    [Fact]
    public void The_KiCad_report_reads_the_published_KiCad_data_too()
    {
        string dataRoot = Path.Combine(this.thisWorkspace.Root, "Data");
        BoardDocument.StoredFileRow pcb = StoredFiles.Put(dataRoot, "KiCad data/board.kicad_pcb", BoardFilesWindowTests.MinimalPcb);
        StoredFiles.WriteBoard(
            DraftBoardSource.PublishedPathOf(dataRoot, BoardFilesWindowTests.BoardKey),
            new BoardData { Components = [new ComponentEntry { BoardLabel = "U1" }] },
            pcb);
        Assert.True(DraftSeeder.SeedFromPublishedFile(DraftManager.DraftsRoot, dataRoot, BoardFilesWindowTests.BoardKey).Created);

        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow { DataRootOverrideForTests = dataRoot };

            window.Initialize(
                "C64 - 250407",
                BoardFilesWindowTests.BoardKey,
                new[] { "U1" },
                BoardFilesSection.KiCadData,
                [Signal("Ground", "GND")]);

            BoardFilesWindowTests.WaitForOpening(window);

            Assert.True(window.GetControl<Border>("ReportPanel").IsVisible);
            Assert.Equal(["U1"], ItemsOf(window, "MatchedList"));
            Assert.Equal(["GND"], NetBadgesOf(window).Select(badge => badge.Name));
            Assert.False(VisibleOf(window, "UnresolvedSignalsPanel"));

            TextBlock source = window.GetControl<TextBlock>("ReportSourceText");
            Assert.True(source.IsVisible);
            Assert.Contains("published KiCad data", source.Text);

            Assert.Empty(window.KiCadFiles);
        });
    }
}
