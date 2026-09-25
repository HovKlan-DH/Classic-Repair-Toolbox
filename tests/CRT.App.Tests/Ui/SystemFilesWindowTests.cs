using Avalonia.Controls;
using Avalonia.LogicalTree;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// SystemFilesWindow - importing board images and KiCad data into a drafted system, and the report
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
public sealed class SystemFilesWindowTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SystemFilesWindowTests()
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

    private static string TextOf(SystemFilesWindow window, string name) =>
        window.GetControl<TextBlock>(name).Text ?? string.Empty;

    private static bool VisibleOf(SystemFilesWindow window, string name) =>
        window.GetControl<StackPanel>(name).IsVisible;

    // The references one of the report's badge lists is showing.
    private static List<string> ItemsOf(SystemFilesWindow window, string name) =>
        (window.GetControl<ItemsControl>(name).ItemsSource ?? Array.Empty<string>())
            .Cast<string>()
            .ToList();

    [Fact]
    public void The_window_constructs_without_throwing()
    {
        UiTest.Run(() => Assert.NotNull(new SystemFilesWindow()));
    }

    [Fact]
    public void The_report_panel_is_hidden_until_there_is_a_report()
    {
        UiTest.Run(() =>
        {
            var window = new SystemFilesWindow();

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
            var window = new SystemFilesWindow();

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
            var window = new SystemFilesWindow();

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
            var window = new SystemFilesWindow();

            window.ShowReport(Report(
                matched: new[] { "U8" },
                unmatched: Array.Empty<string>(),
                unused: Array.Empty<string>()));

            Assert.Equal("1 component matches", TextOf(window, "MatchedHeaderText"));
        });
    }

    // Not an error - this is the to-do list for a board being built up, which is exactly what a
    // brand-new system needs.
    [Fact]
    public void Unlabelled_components_are_shown_as_what_is_left_to_do()
    {
        UiTest.Run(() =>
        {
            var window = new SystemFilesWindow();

            window.ShowReport(Report(
                matched: Array.Empty<string>(),
                unmatched: Array.Empty<string>(),
                unused: new[] { "U8", "C15" }));

            Assert.True(VisibleOf(window, "UnusedPanel"));
            Assert.Equal("2 components are not labelled yet", TextOf(window, "UnusedHeaderText"));
        });
    }

    // ###########################################################################################
    // *** EVERY UNLABELLED COMPONENT IS LISTED - NO CAP (maintainer request, 2026-09-24). ***
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
            var window = new SystemFilesWindow();
            string[] unused = Enumerable.Range(1, 365).Select(i => $"C{i}").ToArray();

            window.ShowReport(Report(
                matched: Array.Empty<string>(),
                unmatched: Array.Empty<string>(),
                unused: unused));

            Assert.Equal(unused, ItemsOf(window, "UnusedList"));
        });
    }

    // ###########################################################################################
    // *** EACH COMPONENT IS ITS OWN BADGE (maintainer request, 2026-09-24). ***
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
            var window = new SystemFilesWindow();

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
            var window = new SystemFilesWindow();

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
            var window = new SystemFilesWindow();

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
            var window = new SystemFilesWindow();
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

        Assert.Equal("Main board", SystemFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    [Fact]
    public void A_taken_schematic_name_gets_a_numbered_suffix()
    {
        var taken = new HashSet<string>(new[] { "Main board" }, StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board (2)", SystemFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    [Fact]
    public void The_suffix_keeps_climbing_past_names_already_taken()
    {
        var taken = new HashSet<string>(
            new[] { "Main board", "Main board (2)", "Main board (3)" },
            StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board (4)", SystemFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    [Fact]
    public void The_collision_check_ignores_casing()
    {
        var taken = new HashSet<string>(new[] { "MAIN BOARD" }, StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board (2)", SystemFilesWindow.ResolveFreeSchematicName("Main board", taken));
    }

    // A file called ".png" has no name to take - it still needs one, or it would be written as a
    // row with a blank natural key that could never be found or removed again.
    [Fact]
    public void A_blank_desired_name_falls_back_to_a_usable_one()
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Board image", SystemFilesWindow.ResolveFreeSchematicName("   ", taken));
    }

    [Fact]
    public void A_desired_name_is_trimmed_before_it_is_used()
    {
        var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        Assert.Equal("Main board", SystemFilesWindow.ResolveFreeSchematicName("  Main board  ", taken));
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

    private const string SystemKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    // Creates a real drafted system holding one schematic row, and optionally the image file that
    // row names. Goes through DraftSeeder and DraftWorkbookStore rather than hand-writing files, so
    // the window reads exactly the shape the application really produces.
    private void SeedSystemWithSchematic(string imageFileName, bool writeTheImageFile)
    {
        DraftSeeder.CreateNewSystem(
            DraftManager.DraftsRoot,
            new NewSystemRegistration
            {
                HardwareName = "C64",
                BoardName = "250407",
                ExcelDataFile = SystemFilesWindowTests.SystemKey,
            });

        DraftWorkbookStore.Edit(
            DraftManager.DraftsRoot,
            SystemFilesWindowTests.SystemKey,
            board =>
            {
                // ###########################################################################
                // *** STORED DATA-ROOT RELATIVE, exactly as the application writes it. ***
                //
                // These tests originally stored a BARE FILE NAME, and passed against a
                // LoadThumbnail that combined the draft folder with the stored value by hand -
                // which on real data produces a doubled path and showed "No preview" for every
                // image. An unrealistic fixture hid a real defect, so the prefix here is
                // load-bearing, not decoration.
                // ###########################################################################
                board.Schematics.Add(new BoardSchematicEntry
                {
                    SchematicName = "Top side",
                    SchematicImageFile = SystemFilesWindowTests.StoredImagePath(imageFileName),
                });

                return board;
            });

        if (writeTheImageFile)
        {
            // The bytes live directly in the draft's system folder, under the bare name.
            File.WriteAllBytes(
                Path.Combine(
                    DraftManager.GetSystemFolder(SystemFilesWindowTests.SystemKey),
                    imageFileName),
                SystemFilesWindowTests.OnePixelPng);
        }
    }

    // The value a board workbook really holds: relative to the DATA ROOT, forward slashes, and
    // therefore already carrying the system's own three folder segments.
    private static string StoredImagePath(string fileName) =>
        "Commodore/C64/250407/" + fileName;

    // The window shows ONE section per opening (the Drafts tab has a button for each), and loads
    // only that section's data - so every test names the section it is about.
    private SystemFilesWindow OpenOnSeededSystem(SystemFilesSection section)
    {
        var window = new SystemFilesWindow();

        window.Initialize("C64 - 250407", SystemFilesWindowTests.SystemKey, Array.Empty<string>(), section);

        return window;
    }

    // ------------------------------------------------------------------ ONE section per opening
    //
    // *** THE DRAFTS TAB HAS A BUTTON EACH FOR "Schematic images" AND "KiCad data" (maintainer
    // request, 2026-09-24), where it used to have one for both. *** Each opens this window on its
    // own section: that section is the only one shown, the title names it, and - the part a
    // screenshot cannot show - only that section's data is loaded. Opening "Schematic images" must
    // not parse the KiCad project, and opening "KiCad data" must not decode the board images.

    [Fact]
    public void Opening_on_SCHEMATIC_IMAGES_shows_and_loads_only_the_images()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.SchematicImages);

            Assert.Equal("Schematic images", window.Title);
            Assert.True(window.GetControl<StackPanel>("SchematicsSection").IsVisible);
            Assert.False(window.GetControl<StackPanel>("KiCadSection").IsVisible);

            Assert.Single(window.Schematics);

            // The KiCad file is on disk, and deliberately not read.
            Assert.Empty(window.KiCadFiles);
            Assert.False(window.GetControl<Border>("ReportPanel").IsVisible);
        });
    }

    [Fact]
    public void Opening_on_KICAD_DATA_shows_and_loads_only_the_KiCad_files()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

            Assert.Equal("KiCad data", window.Title);
            Assert.False(window.GetControl<StackPanel>("SchematicsSection").IsVisible);
            Assert.True(window.GetControl<StackPanel>("KiCadSection").IsVisible);

            Assert.Single(window.KiCadFiles);

            // The schematic row exists in the draft, and its preview is deliberately not decoded.
            Assert.Empty(window.Schematics);
        });
    }

    // The line under the system name describes the section that is open, not both.
    [Fact]
    public void The_introduction_describes_only_the_section_that_is_open()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            string images = TextOf(this.OpenOnSeededSystem(SystemFilesSection.SchematicImages), "IntroText");
            string kiCad = TextOf(this.OpenOnSeededSystem(SystemFilesSection.KiCadData), "IntroText");

            Assert.Contains("schematic images", images, StringComparison.Ordinal);
            Assert.DoesNotContain("KiCad", images, StringComparison.Ordinal);

            Assert.Contains("KiCad project", kiCad, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // *** EACH SCHEMATIC ROW CARRIES A PREVIEW OF ITS OWN IMAGE (maintainer request, 2026-09-24).
    // ***
    //
    // The list used to name a file and nothing more, so the only way to find out whether the right
    // image had been imported - or whether it was the right way up, or the right side of the board
    // - was to go and open it. This is the "does that look right?" glance.
    // ###########################################################################################
    [Fact]
    public void A_schematic_row_carries_a_preview_of_its_image()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.SchematicImages);

            SystemSchematicRow row = Assert.Single(window.Schematics);

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
        this.SeedSystemWithSchematic("gone.png", writeTheImageFile: false);

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.SchematicImages);

            SystemSchematicRow row = Assert.Single(window.Schematics);

            // The row is still listed - the schematic exists in the draft even if its file does not.
            Assert.Equal("Top side", row.SchematicName);
            Assert.Equal(SystemFilesWindowTests.StoredImagePath("gone.png"), row.ImageFile);

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
        this.SeedSystemWithSchematic("broken.png", writeTheImageFile: false);

        File.WriteAllText(
            Path.Combine(
                DraftManager.GetSystemFolder(SystemFilesWindowTests.SystemKey),
                "broken.png"),
            "this is not a png");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.SchematicImages);

            SystemSchematicRow row = Assert.Single(window.Schematics);

            // Still listed, still named, still removable - the row is not lost with its bytes.
            Assert.Equal("Top side", row.SchematicName);
            Assert.Equal(SystemFilesWindowTests.StoredImagePath("broken.png"), row.ImageFile);

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
        this.SeedSystemWithSchematic("maybe.png", writeTheImageFile: imageExists);

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.SchematicImages);

            SystemSchematicRow row = Assert.Single(window.Schematics);

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
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.SchematicImages);

            Avalonia.Media.Imaging.Bitmap? first = Assert.Single(window.Schematics).Thumbnail;
            Assert.NotNull(first);

            // Re-initialising is what an import does once it has written its files.
            window.Initialize("C64 - 250407", SystemFilesWindowTests.SystemKey, Array.Empty<string>(), SystemFilesSection.SchematicImages);

            Avalonia.Media.Imaging.Bitmap? second = Assert.Single(window.Schematics).Thumbnail;
            Assert.NotNull(second);

            // A fresh decode, and the old one released.
            Assert.NotSame(first, second);
            Assert.Throws<ObjectDisposedException>(() => first!.PixelSize);
        });
    }

    // ------------------------------------------------------------------ the imported KiCad FILES

    // Writes KiCad files into the seeded system's "KiCad data" folder, as an import would.
    private void WriteKiCadFiles(params string[] relativePaths)
    {
        string root = Path.Combine(
            DraftManager.GetSystemFolder(SystemFilesWindowTests.SystemKey),
            "KiCad data");

        foreach (string relativePath in relativePaths)
        {
            string full = Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar));

            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            File.WriteAllText(full, "(kicad_sch (version 20230121))");
        }
    }

    // ###########################################################################################
    // *** THE WINDOW NAMES THE IMPORTED FILES, not just how many there are (maintainer report,
    // 2026-09-24). ***
    //
    // Reopening this window after an import said "25 KiCad files imported" and nothing else, so a
    // contributor coming back could not tell which project was in the system - an import of the
    // right folder and an import of the wrong one looked identical.
    // ###########################################################################################
    [Fact]
    public void The_imported_KiCad_files_are_LISTED_not_just_counted()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb", "Pages/vic.kicad_sch");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

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
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Pages/cpu-8500.kicad_sch");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

            SystemKiCadFileRow row = Assert.Single(window.KiCadFiles);

            Assert.Equal("Pages/cpu-8500.kicad_sch", row.RelativePath);

            // Forward slashes whatever the platform separator is - this is how KiCad and the Wiki
            // both write it, and a backslash here would read as a Windows-only detail leaking out.
            Assert.DoesNotContain('\\', row.RelativePath);
        });
    }

    // ###########################################################################################
    // *** THE LISTING HAS NO SCROLLER OF ITS OWN (maintainer report, 2026-09-24). ***
    //
    // It used to sit in its own ScrollViewer capped at 150px, inside the window's ScrollViewer - a
    // scroll inside a scroll, where the wheel moved whichever one the pointer was over. Every file
    // is now shown at full height and the WINDOW's scroller is the only one. Asserted on the
    // logical tree: exactly one ScrollViewer above the list, and no height cap on its panel.
    // ###########################################################################################
    [Fact]
    public void The_KiCad_file_listing_shows_every_file_with_no_scroller_of_its_own()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles(Enumerable.Range(1, 25).Select(i => $"Pages/sheet{i}.kicad_sch").ToArray());

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

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
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

            Assert.NotEmpty(Assert.Single(window.KiCadFiles).SizeText);
        });
    }

    [Fact]
    public void A_system_with_NO_KiCad_data_lists_nothing_and_says_so()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

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
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

            Assert.Single(window.KiCadFiles);

            window.Initialize("C64 - 250407", SystemFilesWindowTests.SystemKey, Array.Empty<string>(), SystemFilesSection.KiCadData);

            Assert.Single(window.KiCadFiles);
        });
    }

    // ------------------------------------------------------------------ REMOVING a KiCad file
    //
    // *** EACH IMPORTED KiCad FILE CAN BE REMOVED ON ITS OWN (maintainer request, 2026-09-24), the
    // way a schematic image can. *** The delete itself - and what it refuses - is covered by
    // KiCadImportedFilesTests; these cover the window offering it and describing what is left.

    [Fact]
    public void Every_listed_KiCad_file_has_its_own_Remove_button()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb", "Open128.kicad_pro", "Pages/vic.kicad_sch");

        UiTest.Run(() =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);
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
                    removeButtons.Select(button => Assert.IsType<SystemKiCadFileRow>(button.Tag).RelativePath));
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
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Pages/vic.kicad_sch");

        // A REAL minimal PCB for the file that stays: the report is rebuilt after the remove, and an
        // unreadable leftover would (correctly) replace the "Removed" message with a read error.
        File.WriteAllText(
            Path.Combine(
                DraftManager.GetSystemFolder(SystemFilesWindowTests.SystemKey),
                "KiCad data",
                "Open128.kicad_pcb"),
            SystemFilesWindowTests.MinimalPcb);

        await UiTest.RunAsync(async () =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

            SystemKiCadFileRow page = window.KiCadFiles.Single(row => row.RelativePath == "Pages/vic.kicad_sch");

            await window.RemoveKiCadFileAsync(page);

            Assert.False(File.Exists(page.FullPath));
            Assert.Equal(new[] { "Open128.kicad_pcb" }, window.KiCadFiles.Select(row => row.RelativePath));
            Assert.Equal("1 KiCad file imported:", window.GetControl<TextBlock>("KiCadStateText").Text);
            Assert.Contains("Removed [Pages/vic.kicad_sch]", window.StatusTextForTests, StringComparison.Ordinal);
        });
    }

    // Removing the last file puts the window back where it was before any import.
    [Fact]
    public async Task Removing_the_LAST_KiCad_file_says_there_is_none_and_hides_the_list()
    {
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        await UiTest.RunAsync(async () =>
        {
            SystemFilesWindow window = this.OpenOnSeededSystem(SystemFilesSection.KiCadData);

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
    // *** THE MATCH REPORT IS REBUILT WHEN THE WINDOW REOPENS (maintainer report, 2026-09-24). ***
    //
    // It used to be produced only at the end of an import, so a contributor returning to a system
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
        this.SeedSystemWithSchematic("top.png", writeTheImageFile: true);

        string root = Path.Combine(
            DraftManager.GetSystemFolder(SystemFilesWindowTests.SystemKey),
            "KiCad data");

        Directory.CreateDirectory(root);
        File.WriteAllText(Path.Combine(root, "board.kicad_pcb"), SystemFilesWindowTests.MinimalPcb);

        await UiTest.RunAsync(async () =>
        {
            var window = new SystemFilesWindow();

            // A board label that DOES match the footprint above, so the report has something to say.
            window.Initialize(
                "C64 - 250407",
                SystemFilesWindowTests.SystemKey,
                new[] { "U1" },
                SystemFilesSection.KiCadData);

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
}
