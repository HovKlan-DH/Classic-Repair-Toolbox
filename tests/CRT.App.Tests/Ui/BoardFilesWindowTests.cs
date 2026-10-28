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

    // The window on the seeded board, its opening read.
    private BoardFilesWindow OpenOnSeededBoard()
    {
        var window = new BoardFilesWindow();

        window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>());
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

    // The window is "KiCad data" - its "Schematic images" half moved to the Contribute tab's
    // Pictures part (CRT 4.0 step 3) - and opens on the board's KiCad files.
    [Fact]
    public void The_window_is_KiCad_data_and_opens_on_the_boards_KiCad_files()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard();

            Assert.Equal("KiCad data", window.Title);
            Assert.True(window.GetControl<StackPanel>("KiCadSection").IsVisible);
            Assert.Single(window.KiCadFiles);
        });
    }

    // ###########################################################################################
    // *** OPENING IS UNDER THE WINDOW'S "PLEASE WAIT" (owner request, 2026-10-10: "it can take a
    // few secs to display, so please do show that new standard wait method"). *** From the first
    // moment - before anything is read - until all of it is in, and then it lets go.
    // ###########################################################################################
    [Fact]
    public void Opening_reads_under_the_windows_please_wait_and_lets_go_when_done()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>());

            BusyOverlay overlay = BusyOverlay.For(window)!;
            Assert.True(overlay.IsBusy);
            Assert.Equal(CrtWaitWording.CheckingKiCadMatches, overlay.Message);

            BoardFilesWindowTests.WaitForOpening(window);

            Assert.False(overlay.IsBusy);
        });
    }

    // ###########################################################################################
    // *** OPENING AND CLOSING CHANGES NOTHING (owner report, 2026-10-10: closing the window
    // refreshed the detached schematic thumbnails although nothing had changed). *** The Drafts tab
    // reloads the board on screen only when ChangedTheDraft says so - and looking is no change,
    // however the KiCad report was read.
    // ###########################################################################################
    [Fact]
    public void Opening_and_closing_without_a_change_says_nothing_changed()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);
        this.WriteKiCadFiles("Open128.kicad_pcb");

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard();
            window.Show();
            Dispatcher.UIThread.RunJobs();
            window.Close();

            Assert.False(window.ChangedTheDraft);
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
            BoardFilesWindow window = this.OpenOnSeededBoard();

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
            BoardFilesWindow window = this.OpenOnSeededBoard();

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
            BoardFilesWindow window = this.OpenOnSeededBoard();

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
            BoardFilesWindow window = this.OpenOnSeededBoard();

            Assert.NotEmpty(Assert.Single(window.KiCadFiles).SizeText);
        });
    }

    [Fact]
    public void A_board_with_NO_KiCad_data_lists_nothing_and_says_so()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        UiTest.Run(() =>
        {
            BoardFilesWindow window = this.OpenOnSeededBoard();

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
            BoardFilesWindow window = this.OpenOnSeededBoard();

            Assert.Single(window.KiCadFiles);

            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>());

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
            BoardFilesWindow window = this.OpenOnSeededBoard();
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
            BoardFilesWindow window = this.OpenOnSeededBoard();

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

            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, Array.Empty<string>());
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
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, ["U1"]);
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
            BoardFilesWindow window = this.OpenOnSeededBoard();

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
                new[] { "U1" });

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

    // ---- dropped on the Contribute tab, and the KiCad summary (step 3, slice 3e) ----------------------

    private static async Task UntilImported(BoardFilesWindow window)
    {
        for (int i = 0; i < 500 && !(window.DroppedImport.IsCompleted && window.DraftChanging.IsCompleted); i++)
        {
            await Task.Delay(10);
            Dispatcher.UIThread.RunJobs();
        }

        Dispatcher.UIThread.RunJobs();
        Assert.True(window.DroppedImport.IsCompleted, "The files dropped should have been imported.");
    }

    // ###########################################################################################
    // The plan's "KiCad: drop the project folder and get a summary of what was found, what was
    // ignored, and what is left to do": a folder dropped on the Contribute tab is imported as the
    // window opens, and the summary says what the KiCad data holds, what was left out - its schematic
    // still in an older KiCad's form, said with what to do - and what is left to do.
    // ###########################################################################################
    [Fact]
    public async Task A_KiCad_folder_dropped_is_imported_as_the_window_opens_and_summed_up()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string source = this.thisWorkspace.Path_("Project");
        File.WriteAllText(this.thisWorkspace.Path_("Project", "board.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);
        File.WriteAllText(this.thisWorkspace.Path_("Project", "board.sch"), "EESchema Schematic File Version 4");
        File.WriteAllText(this.thisWorkspace.Path_("Project", "lib.pretty", "DIP-40.kicad_mod"), "(footprint)");
        File.WriteAllText(this.thisWorkspace.Path_("Project", "3d", "socket.step"), "ISO-10303-21;");

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, ["U1"]);
            window.ImportOnceOpen([source]);

            await UntilImported(window);
            for (int i = 0; i < 200 && !window.GetControl<Border>("KiCadSummaryPanel").IsVisible; i++)
            {
                await Task.Delay(10);
                Dispatcher.UIThread.RunJobs();
            }

            Assert.Equal(["board.kicad_pcb"], window.KiCadFiles.Select(row => row.RelativePath));
            Assert.True(window.GetControl<Border>("KiCadSummaryPanel").IsVisible);
            Assert.Equal("Found 1 PCB with 1 component and 1 net, and no schematic sheet.", TextOf(window, "KiCadFoundText"));
            Assert.Equal(
                "Left out, as CRT does not read them: 1 footprint, 1 3D model and 1 file from an older KiCad.",
                TextOf(window, "KiCadLeftOutText"));
            Assert.Equal(string.Join(" ", KiCadImportSummary.CannotRead(["board.sch"], ["board.sch"], ["board.kicad_pcb"])), TextOf(window, "KiCadOlderText"));
            Assert.True(window.GetControl<TextBlock>("KiCadOlderText").IsVisible);

            // What is left to do, by the checklist's rules on the draft as it now is.
            Assert.True(VisibleOf(window, "KiCadLeftToDoPanel"));
            Assert.Contains("- Match each picture to the KiCad view it shows.", TextOf(window, "KiCadLeftToDoText"), StringComparison.Ordinal);

            // The check done, the status line no longer says it is checking (render, 2026-10-28: it
            // read "Checking what it lines up with..." under the finished summary).
            Assert.Equal("Imported 1 KiCad file - the summary above says what it holds.", window.StatusTextForTests);
        });
    }

    // ###########################################################################################
    // The folder's walk is ONE OF THE WINDOW'S CHANGES (code review, 2026-10-28): moved off the UI
    // thread but outside the queue, the window could be closed while it ran - closing waits only for
    // what is queued - and the import went into the draft after, the board on screen never read again.
    // ###########################################################################################
    [Fact]
    public async Task Closing_waits_for_a_folder_still_being_read()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string source = this.thisWorkspace.Path_("Project");
        File.WriteAllText(this.thisWorkspace.Path_("Project", "board.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);

        using var release = new ManualResetEventSlim(false);

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow { ScanningKiCadForTests = () => release.Wait(TimeSpan.FromSeconds(10)) };
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, ["U1"]);
            BoardFilesWindowTests.WaitForOpening(window);

            Task importing = window.ImportKiCadFolderAsync(source);

            for (int i = 0; i < 10; i++)
            {
                await Task.Delay(5);
                Dispatcher.UIThread.RunJobs();
            }

            Assert.False(window.DraftChanging.IsCompleted, "The folder being read is a change closing has to wait for.");

            release.Set();

            for (int i = 0; i < 500 && !(importing.IsCompleted && window.DraftChanging.IsCompleted); i++)
            {
                await Task.Delay(10);
                Dispatcher.UIThread.RunJobs();
            }

            Assert.True(window.ChangedTheDraft);
            Assert.Equal(["board.kicad_pcb"], window.KiCadFiles.Select(row => row.RelativePath));
        });
    }

    // ###########################################################################################
    // .sch and .brd are EAGLE's too (code review, 2026-10-28): an EAGLE project was told it came from
    // an older KiCad and to open it in KiCad and save it - which does nothing for EAGLE. Only a file
    // that SAYS it is KiCad's (its first line) is an older KiCad's; EAGLE's is told how KiCad imports it.
    // ###########################################################################################
    [Fact]
    public async Task An_EAGLE_project_is_not_called_an_older_KiCad()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string source = this.thisWorkspace.Path_("Eagle");
        File.WriteAllText(this.thisWorkspace.Path_("Eagle", "board.sch"), "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<!DOCTYPE eagle SYSTEM \"eagle.dtd\">\n<eagle version=\"9.6.2\">");
        File.WriteAllText(this.thisWorkspace.Path_("Eagle", "board.brd"), "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<!DOCTYPE eagle SYSTEM \"eagle.dtd\">\n<eagle version=\"9.6.2\">");

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, ["U1"]);
            BoardFilesWindowTests.WaitForOpening(window);

            await window.ImportKiCadFolderAsync(source);

            Assert.Empty(window.KiCadFiles);
            Assert.DoesNotContain("older KiCad", window.StatusTextForTests, StringComparison.Ordinal);
            Assert.Contains("EAGLE", window.StatusTextForTests, StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // An import whose copy failed imported nothing, so the summary says nothing of it (code review,
    // 2026-10-28: it was recorded before the copy, and a later report summed up what it "left out").
    // ###########################################################################################
    [Fact]
    public async Task A_copy_that_failed_leaves_nothing_for_the_summary()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        // A FILE where the draft's "KiCad data" folder goes: nothing can be copied there.
        File.WriteAllText(Path.Combine(DraftManager.GetBoardFolder(BoardFilesWindowTests.BoardKey), "KiCad data"), "in the way");

        string source = this.thisWorkspace.Path_("Project");
        File.WriteAllText(this.thisWorkspace.Path_("Project", "board.kicad_pcb"), BoardFilesWindowTests.MinimalPcb);
        File.WriteAllText(this.thisWorkspace.Path_("Project", "lib.pretty", "DIP-40.kicad_mod"), "(footprint)");

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, ["U1"]);
            BoardFilesWindowTests.WaitForOpening(window);

            await window.ImportKiCadFolderAsync(source);

            Assert.Equal("Could not copy those KiCad files - see the log for details.", window.StatusTextForTests);
            Assert.False(window.HasLastImportForTests);
        });
    }

    // A folder from an older KiCad only: nothing imported, and the status line says why and what to do.
    [Fact]
    public async Task An_older_KiCads_project_dropped_says_what_to_do()
    {
        this.SeedBoardWithSchematic("top.png", writeTheImageFile: true);

        string source = this.thisWorkspace.Path_("Old");
        File.WriteAllText(this.thisWorkspace.Path_("Old", "c64.sch"), "EESchema Schematic File Version 4");
        File.WriteAllText(this.thisWorkspace.Path_("Old", "c64.pro"), "update=");
        Directory.CreateDirectory(source);

        await UiTest.RunAsync(async () =>
        {
            var window = new BoardFilesWindow();
            window.Initialize("C64 - 250407", BoardFilesWindowTests.BoardKey, ["U1"]);
            window.ImportOnceOpen([source]);

            await UntilImported(window);

            Assert.Empty(window.KiCadFiles);
            Assert.Equal(KiCadImportSummary.NothingToImport(["c64.pro", "c64.sch"], ["c64.sch"]), window.StatusTextForTests);
            Assert.StartsWith("No KiCad files CRT can read in that folder. Its schematic files (.sch)", window.StatusTextForTests, StringComparison.Ordinal);
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

    // The views are picked on the picture cards of the Contribute tab's Pictures part since step 3 -
    // the list says where, in the words the card uses.
    [Fact]
    public void The_view_names_list_says_they_are_picked_on_the_picture_cards()
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

            Assert.Contains("\"Shows KiCad view\"", explanation);
            Assert.Contains("\"Pictures\"", explanation);
            Assert.DoesNotContain("Schematic images", explanation);
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
            BoardFilesWindow window = this.OpenOnSeededBoard();

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
            BoardFilesWindow window = this.OpenOnSeededBoard();

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
