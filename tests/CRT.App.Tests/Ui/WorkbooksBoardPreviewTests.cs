using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// The Workbooks tab's board pane (marker 3 in the mockup): every schematic image with one or more
// entries in the SELECTED workbook, each drawn with its real entries.
//
// COLLECTION NOTE: "HeadlessUi", for the same reason WorkbooksListTests is - constructing a control
// needs the shared dispatcher thread. This file does NOT touch DataManager's static state (the
// "DataManager" collection a class can only join one of), because the schematic image path is
// built as an ABSOLUTE path and handed to BoardSchematicEntry.SchematicImageFile directly.
// Path.Combine(anything, anAbsolutePath) always returns just the absolute path - a documented BCL
// contract, not implementation-specific behaviour - so the code under test's
// Path.Combine(DataManager.DataRoot, schematic.SchematicImageFile) resolves correctly no matter
// what DataManager.DataRoot currently holds, including a value another parallel test left behind.
//
// Like WorkbooksListTests, EVERY assertion runs inside UiTest.Run - reading a control's property
// from the test thread throws even for a plain GetControl, because the name-scope lookup itself
// reads a styled property.
[Collection("HeadlessUi")]
public sealed class WorkbooksBoardPreviewTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    // A board key unique to THIS test instance (xunit constructs the class per test), so the
    // UserSettings.ActiveWorkbookIdByBoard entry that BuildTab's activation now writes cannot leak
    // into the next test and pre-select a workbook it never activated. Was a shared literal back
    // when selection was set on an in-memory field and nothing was persisted.
    private readonly string thisBoardKey = "Commodore 64|250469 " + Guid.NewGuid().ToString("N");


    public void Dispose()
    {
        this.LoadWorklog();
        this.thisWorkspace.Dispose();
    }

    private string LoadWorklog()
    {
        string root = this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(root);
        return root;
    }

    // A minimal but genuinely valid 1x1 PNG, written as bytes rather than through Avalonia's own
    // bitmap encoder - the code under test does a plain File-based `new Bitmap(path)` load, so the
    // fixture only needs to be a file libpng can decode, not anything Avalonia produced.
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

    private string WriteSchematicImage(string fileName)
    {
        string path = this.thisWorkspace.Path_(fileName);
        File.WriteAllBytes(path, OnePixelPng);
        return path;
    }

    // A real N x N PNG, unlike OnePixelPng - needed for the positioning test below, where an
    // anchored badge and a parked one must land at genuinely different, distinguishable points
    // inside the image. Rendered and encoded by Avalonia itself rather than hand-built bytes, since
    // a specific pixel size (not just "a valid PNG") is what the test needs. Must run inside
    // UiTest.Run - RenderTargetBitmap needs the headless session's compositor.
    private string WriteSizedSchematicImage(string fileName, int pixelSize)
    {
        string path = this.thisWorkspace.Path_(fileName);

        using var renderTarget = new RenderTargetBitmap(new PixelSize(pixelSize, pixelSize), new Vector(96, 96));
        using (var context = renderTarget.CreateDrawingContext())
        {
            context.DrawRectangle(Brushes.White, null, new Rect(0, 0, pixelSize, pixelSize));
        }

        using (var stream = File.Create(path))
        {
            renderTarget.Save(stream, PngBitmapEncoderOptions.Default);
        }

        return path;
    }

    private static BoardData BuildBoardData(params BoardSchematicEntry[] schematics) =>
        new() { Schematics = schematics.ToList() };

    // Stands in for Main.ActivateWorkbook, doing exactly what it does: persist the choice to
    // UserSettings.ActiveWorkbookIdByBoard, then refresh - Main goes via RefreshWorklogBar, which
    // calls straight back into RefreshWorkbooks, so RefreshWorkbooks IS the refresh here.
    //
    // Injecting this rather than letting the tab branch on "MainWindow == null" is the whole point:
    // with a branch, every test below ran a test-only fallback that set the selected id directly,
    // while the shipped path - persist, then re-derive the selection from the saved id - was pinned
    // by nothing. Now there is one path and these tests are on it.
    private static void InstallActivation(TabWorkbooks tab)
    {
        tab.ActivateWorkbookOverrideForTests = (boardKey, workbookId) =>
        {
            UserSettings.SetActiveWorkbookId(boardKey, workbookId);
            tab.RefreshWorkbooks();
        };
    }

    private static TabWorkbooks BuildTab(string boardKey, BoardData boardData, int selectedWorkbookId)
    {
        var tab = new TabWorkbooks
        {
            BoardKeyOverrideForTests = boardKey,
            CurrentBoardDataOverrideForTests = boardData
        };
        InstallActivation(tab);
        tab.RefreshWorkbooks();
        tab.SelectWorkbookForTests(selectedWorkbookId);
        return tab;
    }

    // Hosts the tab in a real, shown Window and pumps the dispatcher so a real layout pass runs -
    // plain construction leaves every control's Bounds at 0x0, which is enough for the other tests
    // in this file (they only check WHAT was built, not WHERE it landed) but not for a positioning
    // assertion, which needs Image.SizeChanged to have actually fired. Mirrors
    // SchematicsZoomTests.CreateShownTabWithImage's own pattern for the same reason.
    private static TabWorkbooks BuildShownTab(string boardKey, BoardData boardData, int selectedWorkbookId)
    {
        var tab = BuildTab(boardKey, boardData, selectedWorkbookId);

        var window = new Window { Width = 800, Height = 600, Content = tab };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return tab;
    }

    private static StackPanel PreviewPanel(TabWorkbooks tab) =>
        tab.GetControl<StackPanel>("BoardPreviewPanel");

    [Fact]
    public void A_workbook_with_no_entries_shows_the_empty_hint_and_no_previews()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No entries yet", "");
        Assert.NotNull(workbook);

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, BuildBoardData(), workbook!.Id);

            Assert.Empty(PreviewPanel(tab).Children);
            Assert.True(tab.GetControl<TextBlock>("NoBoardPreviewsText").IsVisible);
        });
    }

    [Fact]
    public void A_schematic_with_one_entry_gets_one_preview_with_the_images_bitmap()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Recap", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(10, 10, 50, 50), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry
        {
            SchematicName = "Sheet 1",
            CadName = "board.kicad_sch",
            SchematicImageFile = imagePath
        });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var previews = PreviewPanel(tab).Children;
            Assert.Single(previews);
            Assert.False(tab.GetControl<TextBlock>("NoBoardPreviewsText").IsVisible);

            var image = ((Control)previews[0]).GetSelfAndVisualDescendants().OfType<Image>().Single();
            Assert.NotNull(image.Source);
        });
    }

    // The board pane decodes full-resolution schematic PNGs (a 4220x2941 sheet is ~47 MB of BGRA),
    // and RefreshBoardPreviews clears and rebuilds the whole pane on every board change, entry save
    // and workbook create/close. Decoding per pass stranded one such surface each time; disposing on
    // clear instead would be worse, because an editor opened from a pill outlives the refresh a save
    // triggers and renders the bitmap this tab handed it.
    //
    // So the bitmaps are shared, one per image path, for the life of the tab. This asserts the
    // sharing directly - the SAME instance across rebuilds - because that is the property both the
    // leak fix and the dispose-order safety rest on.
    [Fact]
    public void A_schematics_bitmap_is_decoded_once_and_reused_across_rebuilds()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Recap", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("shared.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(1, 1, 2, 2), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry
        {
            SchematicName = "Sheet 1",
            SchematicImageFile = imagePath
        });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var first = ((Control)PreviewPanel(tab).Children[0]).GetSelfAndVisualDescendants().OfType<Image>().Single().Source;
            Assert.NotNull(first);

            // A full rebuild, exactly what Main.RefreshWorklogBar drives.
            tab.RefreshBoardPreviewsForCurrentSelection();

            var second = ((Control)PreviewPanel(tab).Children[0]).GetSelfAndVisualDescendants().OfType<Image>().Single().Source;

            Assert.Same(first, second);
        });
    }

    // The highlight-rect override is a TEST SEAM, and a seam that can certify the opposite of
    // production is worse than none. The real cache is always OrdinalIgnoreCase-keyed - Main builds
    // it that way at all four of its write sites, TabSchematics declares it that way - so a test
    // handing over a plain `new Dictionary<...>()` would pin ORDINAL lookup as the contract and
    // pass, while the app behaved the other way round. The setter rejects it instead.
    [Fact]
    public void The_highlight_rect_test_seam_refuses_a_dictionary_with_the_wrong_comparer()
    {
        UiTest.Run(() =>
        {
            var tab = new TabWorkbooks();

            Assert.Throws<ArgumentException>(() =>
                tab.HighlightRectsBySchematicAndLabelOverrideForTests =
                    new Dictionary<string, Dictionary<string, List<Rect>>>());

            // The comparer the app actually uses is accepted, as is clearing the override.
            tab.HighlightRectsBySchematicAndLabelOverrideForTests =
                new Dictionary<string, Dictionary<string, List<Rect>>>(StringComparer.OrdinalIgnoreCase);
            tab.HighlightRectsBySchematicAndLabelOverrideForTests = null;
        });
    }

    // Malformed board data: two schematics whose names differ only in case.
    //
    // Board Excel files sync from classic-repair-toolbox.dk independently of app releases, and
    // BoardDataReader.MapSchematics does no dedup and no uniqueness validation - whatever the
    // Schematics sheet holds arrives verbatim. The pane keys schematics by name with
    // OrdinalIgnoreCase, so "Sheet 1" and "sheet 1" collide here while being two distinct rows
    // everywhere else, and a bare ToDictionary threw ArgumentException on the second one. Nothing in
    // this tab catches, so it propagated out through RefreshWorkbooks and Main.RefreshWorklogBar and
    // took down BOARD SELECTION - not just this tab - on any board carrying such a pair.
    [Fact]
    public void Two_schematics_whose_names_differ_only_in_case_do_not_throw()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Duplicate names", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("dup.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(1, 1, 2, 2), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(
            new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath },
            new BoardSchematicEntry { SchematicName = "sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            // The assertion IS that this returns: before the fix it threw out of BuildTab.
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            // And the surviving row still renders - first wins, the whole pane is not lost.
            Assert.Single(PreviewPanel(tab).Children);
        });
    }

    // Only schematics with an entry in the SELECTED workbook appear - a schematic with entries
    // belonging to a DIFFERENT workbook on the same board must not leak into this one's board pane.
    [Fact]
    public void Only_schematics_with_an_entry_in_the_selected_workbook_are_shown()
    {
        this.LoadWorklog();
        var shown = WorklogManager.CreateWorkbook(this.thisBoardKey, "Shown", "");
        var other = WorklogManager.CreateWorkbook(this.thisBoardKey, "Other", "");
        Assert.NotNull(shown);
        Assert.NotNull(other);

        string sheet1 = this.WriteSchematicImage("sheet1.png");
        string sheet2 = this.WriteSchematicImage("sheet2.png");

        WorklogManager.AddEntry(shown!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "In the selected workbook", "", "Note", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(other!.Id, "Sheet 2", new Avalonia.Rect(0, 0, 10, 10), "In the OTHER workbook", "", "Note", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(
            new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = sheet1 },
            new BoardSchematicEntry { SchematicName = "Sheet 2", SchematicImageFile = sheet2 });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, shown.Id);

            Assert.Single(PreviewPanel(tab).Children);

            var caption = ((Control)PreviewPanel(tab).Children[0])
                .GetSelfAndVisualDescendants().OfType<TextBlock>().First();
            Assert.Equal("Sheet 1", caption.Text);
        });
    }

    // Two entries on the same schematic must both appear on the one preview for it, not spawn a
    // second preview of the same image.
    [Fact]
    public void Two_entries_on_the_same_schematic_share_one_preview()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Two faults", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "First", "", "Note", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sheet 1", new Avalonia.Rect(20, 20, 10, 10), "Second", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);
            Assert.Single(PreviewPanel(tab).Children);
        });
    }

    // An entry that references a schematic no longer present in the board data (renamed or
    // removed) must be skipped rather than crash the pane or show an imageless preview.
    [Fact]
    public void An_entry_for_an_unknown_schematic_is_skipped()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Orphan entry", "");
        Assert.NotNull(workbook);

        WorklogManager.AddEntry(workbook!.Id, "Deleted Sheet", new Avalonia.Rect(0, 0, 10, 10), "Orphaned", "", "Note", "Open", Array.Empty<string>());

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, BuildBoardData(), workbook.Id);

            Assert.Empty(PreviewPanel(tab).Children);
            Assert.True(tab.GetControl<TextBlock>("NoBoardPreviewsText").IsVisible);
        });
    }

    // An entry whose image file is missing from disk must be skipped too, rather than throwing out
    // of the whole refresh and leaving every OTHER schematic's preview unbuilt as well.
    [Fact]
    public void A_missing_image_file_drops_only_its_own_schematic()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Missing file", "");
        Assert.NotNull(workbook);

        string missingPath = this.thisWorkspace.Path_("does-not-exist.png");
        string realPath = this.WriteSchematicImage("sheet-ok.png");

        WorklogManager.AddEntry(workbook!.Id, "Missing Sheet", new Avalonia.Rect(0, 0, 10, 10), "No image", "", "Note", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "OK Sheet", new Avalonia.Rect(0, 0, 10, 10), "Has an image", "", "Note", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(
            new BoardSchematicEntry { SchematicName = "Missing Sheet", SchematicImageFile = missingPath },
            new BoardSchematicEntry { SchematicName = "OK Sheet", SchematicImageFile = realPath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var previews = PreviewPanel(tab).Children;
            Assert.Single(previews);

            var caption = ((Control)previews[0]).GetSelfAndVisualDescendants().OfType<TextBlock>().First();
            Assert.Equal("OK Sheet", caption.Text);
        });
    }

    // Switching the selected workbook must rebuild the board pane for the NEWLY selected one - the
    // whole point of tying marker 3 to selection rather than to the board as a whole.
    [Fact]
    public void Switching_the_selected_workbook_rebuilds_the_board_pane()
    {
        this.LoadWorklog();
        var first = WorklogManager.CreateWorkbook(this.thisBoardKey, "First", "");
        var second = WorklogManager.CreateWorkbook(this.thisBoardKey, "Second", "");
        Assert.NotNull(first);
        Assert.NotNull(second);

        string sheet1 = this.WriteSchematicImage("sheet1.png");
        string sheet2 = this.WriteSchematicImage("sheet2.png");

        WorklogManager.AddEntry(first!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "First's fault", "", "Note", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(second!.Id, "Sheet 2", new Avalonia.Rect(0, 0, 10, 10), "Second's fault", "", "Note", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(
            new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = sheet1 },
            new BoardSchematicEntry { SchematicName = "Sheet 2", SchematicImageFile = sheet2 });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, first.Id);

            var firstCaption = ((Control)PreviewPanel(tab).Children[0])
                .GetSelfAndVisualDescendants().OfType<TextBlock>().First();
            Assert.Equal("Sheet 1", firstCaption.Text);

            tab.SelectWorkbookForTests(second.Id);

            Assert.Single(PreviewPanel(tab).Children);
            var secondCaption = ((Control)PreviewPanel(tab).Children[0])
                .GetSelfAndVisualDescendants().OfType<TextBlock>().First();
            Assert.Equal("Sheet 2", secondCaption.Text);
        });
    }

    // An entry with "show marked area" OFF still needs to be visible - just as a pill, with no
    // bounds rectangle. This is the difference the user's own request called out explicitly: "a
    // worklog can either be a pill in top-right corner, or a marked coloured bounds".
    [Fact]
    public void An_entry_with_show_marked_area_off_still_gets_a_badge()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Parked entry", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "No area shown", "", "Note", "Open",
            Array.Empty<string>(), showMarkedArea: false);

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var preview = (Control)PreviewPanel(tab).Children[0];
            var badgeText = preview.GetSelfAndVisualDescendants()
                .OfType<TextBlock>()
                .FirstOrDefault(t => t.Text == "#1");

            Assert.NotNull(badgeText);
        });
    }

    // Finds the badge whose id text reads "#{id}" and returns its own Canvas position - the bug
    // this guards against (reported: a "show marked area" OFF entry's pill appeared pinned to
    // where the marker was drawn, instead of parked in the corner) is entirely about THIS number,
    // not about whether a badge exists at all, which is why the test above was not enough.
    private static Point BadgeCanvasPosition(Control preview, int entryId)
    {
        var idLabel = preview.GetSelfAndVisualDescendants()
            .OfType<TextBlock>()
            .First(t => t.Text == $"#{entryId}");

        // The badge is the outer Border built in BuildPreviewBadge: idLabel -> StackPanel -> Border.
        var badge = (Control)idLabel.Parent!.Parent!;

        return new Point(Canvas.GetLeft(badge), Canvas.GetTop(badge));
    }

    // The bug being fixed: a "show marked area" OFF entry's badge was anchored to the marker's
    // drawn position exactly like an ON entry's, instead of being parked in the image's top-right
    // corner. This pins the two cases apart using a marker deliberately placed far from the corner
    // (bottom-left-ish, at 10,10 in a 200x200 image) - if the parked badge only fixed itself up
    // marginally, or landed at the marker by coincidence, this still catches it, because it asserts
    // WHERE the parked badge is (near the top-right), not merely that it differs from the anchored
    // one.
    [Fact]
    public void A_parked_badge_sits_in_the_top_right_corner_not_at_its_markers_position()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Mixed entries", "");
        Assert.NotNull(workbook);

        UiTest.Run(() =>
        {
            string imagePath = this.WriteSizedSchematicImage("sheet1.png", 200);

            // Entry 1: area shown, marker near the bottom-left - this one SHOULD anchor there.
            WorklogManager.AddEntry(
                workbook!.Id, "Sheet 1", new Avalonia.Rect(10, 150, 20, 20), "Anchored", "", "Note", "Open",
                Array.Empty<string>(), showMarkedArea: true);

            // Entry 2: area hidden, SAME corner of the image - this one should NOT anchor there;
            // it must be parked in the top-right corner instead.
            WorklogManager.AddEntry(
                workbook.Id, "Sheet 1", new Avalonia.Rect(10, 150, 20, 20), "Parked", "", "Issue", "Open",
                Array.Empty<string>(), showMarkedArea: false);

            var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });
            var tab = BuildShownTab(this.thisBoardKey, boardData, workbook.Id);

            var preview = (Control)PreviewPanel(tab).Children[0];
            var anchoredPosition = BadgeCanvasPosition(preview, 1);
            var parkedPosition = BadgeCanvasPosition(preview, 2);

            var image = preview.GetSelfAndVisualDescendants().OfType<Image>().Single();
            double imageWidth = image.Bounds.Width;

            // Sanity check that a real layout pass actually ran - a 0-width image would make every
            // assertion below vacuously true regardless of whether the fix works.
            Assert.True(imageWidth > 0, "Image was never arranged - BuildShownTab's layout pass did not run.");

            // The anchored badge sits near its marker: bottom-left of a 200x200 image, so a low X
            // and a high Y.
            Assert.True(anchoredPosition.Y > 100, $"Anchored badge should be low in the image, was at Y={anchoredPosition.Y}");

            // The parked badge must NOT be at the marker's position...
            Assert.NotEqual(anchoredPosition, parkedPosition);

            // ...and must actually be in the top-right corner: high X (right side of the image),
            // low Y (top).
            Assert.True(parkedPosition.X > imageWidth / 2, $"Parked badge should be on the right, was at X={parkedPosition.X} (image width {imageWidth})");
            Assert.True(parkedPosition.Y < 100, $"Parked badge should be near the top, was at Y={parkedPosition.Y}");
        });
    }

    // Two entries both marked "show marked area" OFF must both park, stacked rather than
    // overlapping - the same ArrangeInTopRightBlock geometry the real Schematics tab uses for its
    // own parked pills, reused here rather than reinvented.
    [Fact]
    public void Two_parked_badges_stack_without_overlapping()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Two parked", "");
        Assert.NotNull(workbook);

        UiTest.Run(() =>
        {
            string imagePath = this.WriteSizedSchematicImage("sheet1.png", 200);

            WorklogManager.AddEntry(
                workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 5, 5), "First parked", "", "Note", "Open",
                Array.Empty<string>(), showMarkedArea: false);
            WorklogManager.AddEntry(
                workbook.Id, "Sheet 1", new Avalonia.Rect(0, 0, 5, 5), "Second parked", "", "Issue", "Open",
                Array.Empty<string>(), showMarkedArea: false);

            var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });
            var tab = BuildShownTab(this.thisBoardKey, boardData, workbook.Id);

            var preview = (Control)PreviewPanel(tab).Children[0];
            var first = BadgeCanvasPosition(preview, 1);
            var second = BadgeCanvasPosition(preview, 2);

            Assert.NotEqual(first, second);
        });
    }

    // A pill's Border is the click target (see OnPreviewBadgePointerPressed, wired up in
    // BuildSchematicPreview) - a Hand cursor is the visible sign that it is meant to be clicked,
    // matching every other clickable pill in the app (the Schematics tab's own badges, workbook
    // cards in the left panel). Actually driving a PointerPressed through to the modal editor is
    // not exercised here: WorklogEntryEditorWindow.ShowDialog blocks on a real window with nothing
    // to dismiss it in a headless run, which is why the Schematics tab's own equivalent handler
    // (OnWorklogEntryPillPointerPressed) has no test coverage either - see CLAUDE.md's "Deliberately
    // not covered" section for pointer interaction in Tabs/.
    [Fact]
    public void A_pill_has_a_hand_cursor_and_the_canvas_it_sits_on_is_clickable()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Clickable pill", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Click me", "", "Note", "Open",
            Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var preview = (Control)PreviewPanel(tab).Children[0];
            var idLabel = preview.GetSelfAndVisualDescendants().OfType<TextBlock>().First(t => t.Text == "#1");
            var badge = (Control)idLabel.Parent!.Parent!;

            Assert.NotNull(badge.Cursor);
            Assert.NotEqual(Cursor.Default, badge.Cursor);

            var badgeCanvas = (Canvas)badge.Parent!;
            Assert.True(badgeCanvas.IsHitTestVisible, "The badge canvas must accept pointer input for its pills to be clickable.");
        });
    }

    // The bug report this guards against: switching to the Workbooks tab (or the board on it)
    // showed the workbook list correctly but never the schematic images/pills, even though the
    // active board genuinely had worklog data. Root cause was a sequencing gap in
    // Main.OnBoardSelectionChanged - RefreshWorklogBar (which rebuilds this pane) ran BEFORE the
    // await that loads the new board's data, so the pane's very first build used the OLD board's
    // data (or none, on the session's first board) and nothing ever asked it to try again once the
    // real data arrived. RefreshBoardPreviewsForCurrentSelection is the fix: Main calls it a SECOND
    // time, right after board data finishes loading. This pins down that a later call, supplying
    // the board data that "arrived late" in the real bug, actually populates the pane - the same
    // shape as the fix, without needing to construct Main itself (which no test in this suite does).
    [Fact]
    public void A_later_call_with_the_boards_real_data_populates_the_pane_that_an_earlier_call_could_not()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Late-arriving board data", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Needs board data", "", "Note", "Open",
            Array.Empty<string>());

        UiTest.Run(() =>
        {
            // First build: no board data yet, exactly like Main's early RefreshWorklogBar call
            // running before DataManager.LoadBoardDataAsync completes.
            var tab = new TabWorkbooks
            {
                BoardKeyOverrideForTests = this.thisBoardKey,
                CurrentBoardDataOverrideForTests = null
            };
            tab.RefreshWorkbooks();
            tab.SelectWorkbookForTests(workbook.Id);

            Assert.Empty(PreviewPanel(tab).Children);
            Assert.True(tab.GetControl<TextBlock>("NoBoardPreviewsText").IsVisible);

            // The board data "arrives" - the second, later call the real fix adds.
            tab.CurrentBoardDataOverrideForTests = BuildBoardData(
                new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });
            tab.RefreshBoardPreviewsForCurrentSelection();

            Assert.Single(PreviewPanel(tab).Children);
            Assert.False(tab.GetControl<TextBlock>("NoBoardPreviewsText").IsVisible);
        });
    }

    // RefreshBoardPreviewsForCurrentSelection must NOT rebuild the workbook list or the top-line -
    // it exists specifically to be a narrower, cheaper alternative to a full RefreshWorklogBar-style
    // refresh for the one thing that needed a second pass (see the test above). If it silently grew
    // to touch the list too, calling it after every board load would duplicate work RefreshWorkbooks
    // already did in the same round-trip.
    [Fact]
    public void The_second_board_data_refresh_does_not_touch_the_workbook_list()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "List must stay put", "");
        Assert.NotNull(workbook);

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, BuildBoardData(), workbook!.Id);
            string countBefore = tab.GetControl<TextBlock>("WorkbookCountText").Text ?? string.Empty;

            tab.RefreshBoardPreviewsForCurrentSelection();

            Assert.Equal(countBefore, tab.GetControl<TextBlock>("WorkbookCountText").Text);
        });
    }

    // Schematic names are NOT unique across boards - "Motherboard" and "Sheet 1" recur across
    // Commodore revisions. The selection used to be cleared only when its name was absent from the
    // new board's grouped entries, so a name present on both boards carried a selection made against
    // the previous one straight over, and the entry list showed the other board's schematic as the
    // chosen one.
    [Fact]
    public void A_board_switch_drops_the_selected_schematic_even_when_the_new_board_has_that_name()
    {
        this.LoadWorklog();

        const string otherBoardKey = "Commodore 64|250425 (other board)";
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Board A job", "");
        var otherWorkbook = WorklogManager.CreateWorkbook(otherBoardKey, "Board B job", "");
        Assert.NotNull(workbook);
        Assert.NotNull(otherWorkbook);

        string sheetA = this.WriteSchematicImage("shared-a.png");
        string sheetB = this.WriteSchematicImage("shared-b.png");

        // Board A carries both schematics; board B has only the shared name.
        WorklogManager.AddEntry(workbook!.Id, "Motherboard", new Avalonia.Rect(1, 1, 2, 2), "A1", "", "Issue", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sheet 2", new Avalonia.Rect(1, 1, 2, 2), "A2", "", "Issue", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(otherWorkbook!.Id, "Motherboard", new Avalonia.Rect(1, 1, 2, 2), "B1", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(
            new BoardSchematicEntry { SchematicName = "Motherboard", SchematicImageFile = sheetA },
            new BoardSchematicEntry { SchematicName = "Sheet 2", SchematicImageFile = sheetB });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            // Deliberately NOT the alphabetically-first one, so the reset below is observable.
            tab.SelectSchematicForTests("Sheet 2");
            Assert.Contains("Sheet 2", EntriesHeaderText(tab));

            // The board switch, as Main drives it: new board key, new board data, new workbook.
            tab.BoardKeyOverrideForTests = otherBoardKey;
            tab.CurrentBoardDataOverrideForTests = BuildBoardData(
                new BoardSchematicEntry { SchematicName = "Motherboard", SchematicImageFile = sheetA });
            tab.RefreshWorkbooks();

            // "Motherboard" exists on this board too, so a name-only validity test would have kept
            // whatever was selected. The selection must have been re-derived for THIS board.
            Assert.Contains("Motherboard", EntriesHeaderText(tab));
            Assert.DoesNotContain("Sheet 2", EntriesHeaderText(tab));
        });
    }

    // ---------------------------------------------------------------------------------------
    // Schematic selection and the entry list on the right (marker 4): clicking a schematic
    // preview (SelectSchematicForTests stands in for the click - see its comment) selects it and
    // switches the entry list to its entries. Each entry is ONE 1px-bordered card holding three
    // stacked rows - title + "#N" badge, description, category chip + status pill - see
    // BuildEntryDetailCard.
    // ---------------------------------------------------------------------------------------

    private static string EntriesHeaderText(TabWorkbooks tab) =>
        tab.GetControl<TextBlock>("SelectedSchematicEntriesHeaderText").Text ?? string.Empty;

    private static bool NoEntriesTextVisible(TabWorkbooks tab) =>
        tab.GetControl<TextBlock>("NoSelectedSchematicEntriesText").IsVisible;

    private static StackPanel EntriesPanel(TabWorkbooks tab) =>
        tab.GetControl<StackPanel>("SelectedSchematicEntriesPanel");

    // Reads back one entry detail card's four stacked rows, in order: title (the title text, the
    // "#N" badge AND the "Delete worklog" button's label, concatenated - the button shares that
    // row's Grid so it can sit in the card's top-right corner), description, category+status (the
    // two outlined pills' concatenated text), and the stats row (hours/cost/comments/links/photos/
    // files, concatenated).
    // Asserts there are exactly four rows and that the whole card is ONE bordered panel (a single
    // outer Border, not one per field), since that shape - "one border around the worklog, not each
    // element inside it" - is the point being pinned down here.
    private static (string Title, string Description, string CategoryStatus, string Stats) ReadEntryCard(Control card)
    {
        var outerBorder = (Border)card;
        var rows = ((StackPanel)outerBorder.Child!).Children.ToList();
        Assert.Equal(4, rows.Count);

        static string AllText(Control root) => string.Join(
            " ",
            root.GetSelfAndVisualDescendants().OfType<TextBlock>().Select(t => t.Text).Where(t => !string.IsNullOrEmpty(t)));

        string title = AllText(rows[0]);
        string description = AllText(rows[1]);
        string categoryStatus = AllText(rows[2]);
        string stats = AllText(rows[3]);

        // The card's own outline is the ONLY border directly on the card - the title/description/
        // category-status/stats rows are plain panels, not each wrapped in their own Border, which
        // is what distinguishes this shape from the earlier three-separately-bordered-panels layout.
        Assert.DoesNotContain(rows, r => r is Border);

        return (title, description, categoryStatus, stats);
    }

    // ###########################################################################################
    // THE EMPTY STATES on a board with no workbooks at all - reported as reading wrong: both
    // messages sat in the vertical MIDDLE of their (full-height, otherwise empty) panels, far from
    // the headings they belong to, and the entry list's told the user to click a schematic image
    // when there were no schematic images to click.
    // ###########################################################################################

    // Vertical alignment is the whole point here: a TextBlock in a Grid cell defaults to Stretch,
    // which lays a single line of text out centred down the panel. Both must be Top.
    [Fact]
    public void The_empty_state_messages_are_top_aligned_rather_than_floating_mid_panel()
    {
        this.LoadWorklog();

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, BuildBoardData(), selectedWorkbookId: 0);

            Assert.Equal(VerticalAlignment.Top, tab.GetControl<TextBlock>("NoBoardPreviewsText").VerticalAlignment);
            Assert.Equal(VerticalAlignment.Top, tab.GetControl<TextBlock>("NoSelectedSchematicEntriesText").VerticalAlignment);
        });
    }

    // Both panels say the same thing, and neither mentions clicking a schematic image: on a board
    // with no workbooks there is nothing on the left to click, so the old wording described an
    // action the user could not take.
    [Fact]
    public void Both_empty_state_messages_say_no_worklogs_are_recorded_for_the_board()
    {
        this.LoadWorklog();

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, BuildBoardData(), selectedWorkbookId: 0);

            const string expected = "No worklogs recorded yet for any schematics in this board.";

            var boardPane = tab.GetControl<TextBlock>("NoBoardPreviewsText");
            var entryList = tab.GetControl<TextBlock>("NoSelectedSchematicEntriesText");

            Assert.Equal(expected, boardPane.Text);
            Assert.Equal(expected, entryList.Text);

            Assert.True(boardPane.IsVisible);
            Assert.True(entryList.IsVisible);

            // The message the entry list used to carry named an action that does not exist on an
            // empty board - there is no schematic image on the left to click.
            Assert.DoesNotContain("Click a schematic", entryList.Text!, StringComparison.OrdinalIgnoreCase);
        });
    }

    [Fact]
    public void The_first_schematic_is_selected_by_default_and_highlighted()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Default selection", "");
        Assert.NotNull(workbook);

        string sheet1 = this.WriteSchematicImage("sheet1.png");
        string sheet2 = this.WriteSchematicImage("sheet2.png");

        WorklogManager.AddEntry(workbook!.Id, "Sheet A", new Avalonia.Rect(0, 0, 10, 10), "First sheet's entry", "", "Note", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sheet B", new Avalonia.Rect(0, 0, 10, 10), "Second sheet's entry", "", "Note", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(
            new BoardSchematicEntry { SchematicName = "Sheet A", SchematicImageFile = sheet1 },
            new BoardSchematicEntry { SchematicName = "Sheet B", SchematicImageFile = sheet2 });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            // "Sheet A" sorts first (RefreshBoardPreviews orders schematics alphabetically), so it
            // is the default selection.
            Assert.Contains("Sheet A", EntriesHeaderText(tab));
            Assert.Single(EntriesPanel(tab).Children);

            var previews = PreviewPanel(tab).Children.Cast<Border>().ToList();
            var sheetA = previews.Single(p => (string)p.Tag! == "Sheet A");
            var sheetB = previews.Single(p => (string)p.Tag! == "Sheet B");

            Assert.NotEqual(sheetA.BorderBrush, sheetB.BorderBrush);
        });
    }

    [Fact]
    public void Selecting_a_different_schematic_switches_the_entry_list_and_the_highlight()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Switch selection", "");
        Assert.NotNull(workbook);

        string sheet1 = this.WriteSchematicImage("sheet1.png");
        string sheet2 = this.WriteSchematicImage("sheet2.png");

        WorklogManager.AddEntry(workbook!.Id, "Sheet A", new Avalonia.Rect(0, 0, 10, 10), "On A", "", "Note", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sheet B", new Avalonia.Rect(0, 0, 10, 10), "On B", "", "Note", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(
            new BoardSchematicEntry { SchematicName = "Sheet A", SchematicImageFile = sheet1 },
            new BoardSchematicEntry { SchematicName = "Sheet B", SchematicImageFile = sheet2 });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);
            Assert.Contains("Sheet A", EntriesHeaderText(tab));

            var previews = PreviewPanel(tab).Children.Cast<Border>().ToList();
            var sheetA = previews.Single(p => (string)p.Tag! == "Sheet A");
            var sheetB = previews.Single(p => (string)p.Tag! == "Sheet B");
            var sheetABorderWhenSelected = sheetA.BorderBrush;

            tab.SelectSchematicForTests("Sheet B");

            Assert.Contains("Sheet B", EntriesHeaderText(tab));

            var (title, _, _, _) = ReadEntryCard((Control)EntriesPanel(tab).Children[0]);
            Assert.Contains("On B", title);

            // The highlight must have MOVED, not just appeared on B: A's border reverts to the
            // unselected colour it had before it was ever selected, and B's now matches what A's
            // used to be.
            Assert.NotEqual(sheetABorderWhenSelected, sheetA.BorderBrush);
            Assert.Equal(sheetABorderWhenSelected, sheetB.BorderBrush);
        });
    }

    // ###########################################################################################
    // "Delete worklog" on an entry card - the per-worklog twin of the header's "Delete workbook".
    //
    // What these pin down is the button's PLACEMENT and its wiring, not the delete itself: the
    // click opens a modal a headless test cannot dismiss (the same reason the card's own
    // click-to-open-the-editor is only pinned by its Hand cursor), and WorklogManager.DeleteEntry
    // is covered directly in WorklogManagerTests.
    // ###########################################################################################

    private static Button DeleteWorklogButton(Control card) =>
        card.GetSelfAndVisualDescendants()
            .OfType<Button>()
            .Single(b => (b.Content as string) == "Delete worklog");

    [Fact]
    public void Every_entry_card_carries_its_own_delete_worklog_button()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Delete buttons", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "First", "", "Issue", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Second", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var cards = EntriesPanel(tab).Children.Cast<Control>().ToList();
            Assert.Equal(2, cards.Count);

            // One PER CARD, not one for the list: the button acts on the worklog it sits on, so a
            // shared one at the top of the panel would have nothing to name.
            foreach (var card in cards)
            {
                Assert.NotNull(DeleteWorklogButton(card));
            }
        });
    }

    // Top-RIGHT of the card, which is what was asked for and mirrors where "Delete workbook" sits
    // relative to the workbook it acts on. Asserted structurally rather than by pixels: it is the
    // last column of the title row's Grid (so it is right of the title) and Top-aligned (so a title
    // that wraps to two lines leaves it level with the FIRST line rather than dragging it down the
    // card).
    [Fact]
    public void The_delete_worklog_button_sits_in_the_cards_top_right_corner()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Button placement", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "A title long enough that it would wrap onto a second line in a narrow entry list",
            "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];

            // The FIRST of the card's four rows - top of the card, above the description.
            var titleRow = ((StackPanel)((Border)card).Child!).Children[0];
            var grid = Assert.IsType<Grid>(titleRow);

            var button = DeleteWorklogButton(card);

            // Right: the button owns the Grid's second (Auto) column while the title text sits in
            // the first (star) one, so the title takes the slack and the button hugs the edge.
            Assert.Equal(1, Grid.GetColumn(button));
            Assert.Equal(2, grid.ColumnDefinitions.Count);
            Assert.Equal(GridUnitType.Star, grid.ColumnDefinitions[0].Width.GridUnitType);
            Assert.Equal(GridUnitType.Auto, grid.ColumnDefinitions[1].Width.GridUnitType);

            // Top: level with the first line of a wrapping title.
            Assert.Equal(VerticalAlignment.Top, button.VerticalAlignment);
            Assert.Equal(HorizontalAlignment.Right, button.HorizontalAlignment);
        });
    }

    // The same destructive styling "Delete workbook" carries in the header above: these are the
    // same kind of permanent delete one level apart, and a differently-coloured one here would
    // read as a different kind of action.
    [Fact]
    public void The_delete_worklog_button_uses_the_same_destructive_brushes_as_delete_workbook()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Button styling", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];
            var button = DeleteWorklogButton(card);

            // Asserted against the Button_Cancel_* THEME KEYS, which is what "the same as Delete
            // workbook" actually means - that button names those same three keys in the markup.
            //
            // Deliberately NOT compared against the header button's resolved brushes: its values
            // come from DynamicResource bindings, which have not resolved on a tab that is built
            // but never attached to a window, so it reads Black here and the comparison would be
            // testing Avalonia's binding state rather than this styling. Resolving the keys the
            // same way the code under test does keeps the assertion correct in both themes without
            // hardcoding a colour.
            Assert.Equal(ThemeResources.Resolve<IBrush?>("Button_Cancel_Fg", null), button.Foreground);
            Assert.Equal(ThemeResources.Resolve<IBrush?>("Button_Cancel_Bg", null), button.Background);
            Assert.Equal(ThemeResources.Resolve<IBrush?>("Button_Cancel_Border", null), button.BorderBrush);
            Assert.NotNull(button.Foreground);
        });
    }

    // The card behind the button is clickable as a whole (it opens the editor) and carries a Hand
    // cursor to say so. The button must NOT inherit it: a Hand here would say this does the same
    // benign thing the rest of the card does, when it is the one control on the card that destroys
    // something.
    [Fact]
    public void The_delete_worklog_button_does_not_inherit_the_cards_hand_cursor()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Button cursor", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];

            // Compared by the cursor's own ToString, not by instance: Cursor has no value equality,
            // so two Cursors built from the same StandardCursorType are NOT Equal and an
            // Assert.Equal against a fresh `new Cursor(...)` fails with the baffling
            // "Expected: Hand / Actual: Hand".
            Assert.Equal("Hand", ((Border)card).Cursor?.ToString());
            Assert.NotEqual("Hand", DeleteWorklogButton(card).Cursor?.ToString());
        });
    }

    [Fact]
    public void An_entry_card_shows_id_title_description_category_and_status_in_one_bordered_panel()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Card contents", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Bad cap", "Leaked electrolytic, pads cleaned.", "Issue", "Closed", Array.Empty<string>());
        Assert.NotNull(entry);

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];
            var (title, description, categoryStatus, _) = ReadEntryCard(card);

            Assert.Contains("Bad cap", title);
            Assert.Contains($"#{entry!.Id}", title);
            Assert.Equal("Leaked electrolytic, pads cleaned.", description);
            Assert.Contains("Issue", categoryStatus);
            Assert.Contains("Closed", categoryStatus);
        });
    }

    // The fourth row: total hours/cost (summed across Work done rows) plus how many comments,
    // links, photos and files the entry carries. Populates one of each via WorklogManager.UpdateEntry
    // (the same real persistence path BuildEntryDetailCard's data comes through, not a bare
    // in-memory record) so this proves the card reads real saved data, not just that the row exists.
    [Fact]
    public void An_entry_card_shows_hours_cost_and_item_counts_in_its_stats_row()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Stats row", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Bad cap", "Leaked electrolytic, pads cleaned.", "Issue", "Closed", Array.Empty<string>());
        Assert.NotNull(entry);

        entry!.WorkDoneItems.Add(new WorklogWorkDoneRecord { Id = 1, Text = "Replaced cap", Date = DateTime.Now, HoursSpent = 1.5, Cost = 8 });
        entry.WorkDoneItems.Add(new WorklogWorkDoneRecord { Id = 2, Text = "Tested", Date = DateTime.Now, HoursSpent = 0.5, Cost = 0 });
        // AddEntry already seeded Comments with one automatic "created" entry, so adding ONE more
        // here means the real total is 2 - the count this test asserts against.
        entry.Comments.Add(new WorklogCommentRecord { Id = 2, Text = "Looks fixed", Date = DateTime.Now });
        entry.Links.Add(new WorklogLinkRecord { Id = 1, Headline = "Datasheet", Url = "https://example.invalid" });
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "before.png", DisplayOrder = 0 });
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 2, FileName = "after.png", DisplayOrder = 1 });
        entry.Files.Add(new WorklogAttachmentRecord { Id = 1, FileName = "notes.txt", DisplayOrder = 0 });
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        // The currency is SET here rather than assumed: UserSettings is static and this class does
        // not point it at a temp file, so the code left behind by whichever test ran before this one
        // is what the card would otherwise print. Restored in the finally, for the same reason.
        string savedCurrency = UserSettings.WorklogCurrencyCode;
        try
        {
            UserSettings.WorklogCurrencyCode = "SEK";

            UiTest.Run(() =>
            {
                var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

                var card = (Control)EntriesPanel(tab).Children[0];
                var (_, _, _, stats) = ReadEntryCard(card);

                // "2 hours", not "2 h" - the time reads as hours and minutes everywhere the user
                // sees one (see WorklogDurationFormatter). Asserted with the WORD, since the old
                // "2 h" is a substring of "2 hours" and would pass either way.
                Assert.Contains("2 hours", stats);
                Assert.DoesNotContain("2 h ", stats);

                // The cost carries the configured currency code - not a bare "8" a reader has to
                // guess at. Asserted WITH the code, since a bare "8" is a substring of "8 SEK" and
                // so would pass against a card printing no currency at all.
                Assert.Contains("8 SEK", stats);
                Assert.Contains("2 comments", stats);
                Assert.Contains("1 link", stats);
                Assert.Contains("2 photos", stats);
                Assert.Contains("1 file", stats);

                // This entry HAS both a time and a cost, so both appear - the counterpart to
                // An_entry_card_omits_every_stat_it_has_nothing_to_report_for, which pins the
                // other side of the same rule.
                Assert.Contains("2 hours", stats);
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = savedCurrency;
        }
    }

    // ###########################################################################################
    // EVERY item on the stats row is omitted when it has nothing to report - reported directly,
    // first for the time and the cost and then for the counts beside them. A brand-new worklog used
    // to read "0 h . 0 CHF . 1 comment . 0 links . 0 photos . 0 files": six items, one of which
    // carried information.
    //
    // AddEntry seeds one automatic "created" comment, so that single "1 comment" is the ONLY thing
    // this card can honestly say - which makes it the assertion that matters. Asserting merely that
    // the zeroes are gone would pass against a row that had stopped rendering altogether.
    // ###########################################################################################
    [Fact]
    public void An_entry_card_omits_every_stat_it_has_nothing_to_report_for()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Empty stats", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Just a note", "Nothing attached to this one.", "Note", "Open", Array.Empty<string>());
        Assert.NotNull(entry);

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];
            var (_, _, _, stats) = ReadEntryCard(card);

            // The one true thing, still there.
            Assert.Equal("1 comment", stats.Trim());

            // And nothing reporting an absence. The cost is checked by its CODE rather than by a
            // bare "0", which is a substring of any figure containing a zero (a "10 USD" cost would
            // fail a naive Assert.DoesNotContain("0")).
            Assert.DoesNotContain("USD", stats);
            Assert.DoesNotContain("minute", stats);
            Assert.DoesNotContain("hour", stats);
            Assert.DoesNotContain("0 links", stats);
            Assert.DoesNotContain("0 photos", stats);
            Assert.DoesNotContain("0 files", stats);
        });
    }

    // ###########################################################################################
    // The stats row separates its items with the " · " dot every other multi-part line in this app
    // uses - the summary strip's lines, the workbook card's "6 worklogs · started ...", the
    // header's "#1 · Title". This row was the one exception: it relied on WrapPanel spacing alone,
    // so "175 DKK 3 comments" read as one run of words with a gap in it. Reported directly.
    //
    // The interesting half is the interaction with the omit-empty rule: a separator decided up
    // front leaves a LEADING dot when the first stat is dropped and a DOUBLED one when a middle
    // stat is, which is why the row is asserted here in full rather than merely for containing a
    // dot somewhere. This entry has no time and no cost (the two that lead the row) and no links,
    // so the version that gets it wrong produces "· 1 comment · · 2 photos".
    // ###########################################################################################
    [Fact]
    public void An_entry_cards_stats_are_separated_by_dots_with_none_leading_or_doubled()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Dot separators", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Gaps", "No time, no cost, no links.", "Note", "Open", Array.Empty<string>());
        Assert.NotNull(entry);

        // Deliberately NO WorkDoneItems (so no hours and no cost) and NO links: that leaves the
        // row's first two items and one of its middle items dropped, which is exactly what a
        // naive separator gets wrong. AddEntry seeds the one automatic "created" comment.
        entry!.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "before.png", DisplayOrder = 0 });
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 2, FileName = "after.png", DisplayOrder = 1 });
        entry.Files.Add(new WorklogAttachmentRecord { Id = 1, FileName = "notes.txt", DisplayOrder = 0 });
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];
            var (_, _, _, stats) = ReadEntryCard(card);

            // AllText joins the row's TextBlocks with a single space, and each dot is its own block
            // (so it can wrap down with the stat it introduces rather than dangling at a line end),
            // which is what makes this the whole rendered row.
            Assert.Equal("1 comment · 2 photos · 1 file", stats.Trim());
        });
    }

    // ###########################################################################################
    // The other half of that rule: a row with only ONE item carries no separator at all. A
    // separator appended AFTER each item, rather than before all but the first, passes the test
    // above and fails this one.
    // ###########################################################################################
    [Fact]
    public void An_entry_card_with_a_single_stat_shows_no_dot_at_all()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "One stat", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Lonely", "Only the automatic comment.", "Note", "Open", Array.Empty<string>());
        Assert.NotNull(entry);

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];
            var (_, _, _, stats) = ReadEntryCard(card);

            Assert.Equal("1 comment", stats.Trim());
            Assert.DoesNotContain("·", stats);
        });
    }

    // ###########################################################################################
    // The gap either side of the "·" is EQUAL, and it comes from the separator's own padding
    // rather than from the WrapPanel.
    //
    // Reported: the dots on this row read wider than the summary strip's directly above it. The
    // cause was WrapPanel.ItemSpacing, which falls between EVERY pair of children - so the panel
    // added its gap between a stat and the following dot AND between that dot and its stat, on top
    // of anything the dot carried itself. With ItemSpacing at 0 the padding IS the whole gap, which
    // is what lets one number match the single " · " string every other line in the app uses.
    //
    // Asserted structurally, on the panel and the separator, because the thing that went wrong is
    // invisible to any assertion about text: the row read correctly as "175 DKK · 3 comments" the
    // whole time it was spaced wrongly.
    // ###########################################################################################
    [Fact]
    public void An_entry_cards_stat_separators_carry_their_own_equal_side_spacing()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Dot spacing", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Spacing", "Two stats, so exactly one dot.", "Note", "Open", Array.Empty<string>());
        Assert.NotNull(entry);

        entry!.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "before.png", DisplayOrder = 0 });
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];

            // The card's FOURTH stacked row, indexed the way ReadEntryCard reads it - not searched
            // for by type, since the category/status row above it is a WrapPanel too.
            var statsRow = (WrapPanel)((StackPanel)((Border)card).Child!).Children[3];

            // Zero, or the panel contributes a second gap this row cannot account for.
            Assert.Equal(0, statsRow.ItemSpacing);

            var separator = statsRow.Children
                .OfType<TextBlock>()
                .Single(t => t.Text == "·");

            // Equal left and right: the dot sits BETWEEN two facts, so any asymmetry visibly
            // attaches it to one of them.
            Assert.True(separator.Padding.Left > 0);
            Assert.Equal(separator.Padding.Left, separator.Padding.Right);

            // The stats themselves carry no padding of their own, or it would stack with the
            // separator's and put the gap back where it started.
            foreach (var stat in statsRow.Children.OfType<TextBlock>().Where(t => t.Text != "·"))
                Assert.Equal(new Thickness(0), stat.Padding);
        });
    }

    // ###########################################################################################
    // "The intention with the border, was to have one border around the worklog - not for each
    // element inside it" - a reported regression from the previous three-separately-bordered-panel
    // shape, where the title, description and category/status EACH sat in their own 1px outline.
    // Pins the whole card down to exactly ONE outer Border, with none of its four direct rows
    // (title, description, category+status, stats) wrapped in a border of their own - the id
    // badge, category chip and status pill ARE still their own (outlined, not filled - see
    // WorklogInfoPillBuilder, which now builds both) Border "pills" nested further
    // inside the category/status row, by design, so this does not assert "no Border anywhere in
    // the card" - only that the ROW-level wrapping is gone.
    // ###########################################################################################
    [Fact]
    public void An_entry_card_has_one_outer_border_with_no_border_wrapping_each_row()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Single border", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Bad cap", "Leaked electrolytic, pads cleaned.", "Issue", "Closed", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var card = (Control)EntriesPanel(tab).Children[0];
            var outerBorder = Assert.IsType<Border>(card);

            var rows = ((StackPanel)outerBorder.Child!).Children.ToList();
            Assert.Equal(4, rows.Count);
            Assert.All(rows, row => Assert.False(row is Border, $"Row {row.GetType().Name} should not be individually bordered."));
        });
    }

    [Fact]
    public void A_schematic_with_no_entries_selected_shows_the_empty_state()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No entries at all", "");
        Assert.NotNull(workbook);

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, BuildBoardData(), workbook!.Id);

            Assert.Empty(EntriesPanel(tab).Children);
            Assert.True(NoEntriesTextVisible(tab));
            Assert.Equal("Select a schematic", EntriesHeaderText(tab));
        });
    }

    // ---------------------------------------------------------------------------------------
    // Component scope for the pill's editor (OnPreviewBadgePointerPressed): clicking a pill must
    // open the EXACT same modal TabSchematics.Worklog.cs's own OnWorklogEntryPillPointerPressed
    // opens, including the "Mark components in scope"/"Mark components completed" checklist - a
    // reported gap where this tab's editor was missing that section entirely. Driving a real
    // pointer press through to WorklogEntryEditorWindow.ShowDialog is not possible headlessly (see
    // A_pill_has_a_hand_cursor_and_the_canvas_it_sits_on_is_clickable's own comment), so these
    // tests call BuildWorklogEntryComponentScopeForTests directly - the exact computation
    // OnPreviewBadgePointerPressed hands to WorklogEntryEditorWindow.InitializeComponentScope.
    // ---------------------------------------------------------------------------------------

    // Mirrors TabSchematics.Worklog.cs's own BuildWorklogEntryComponentScope test coverage: an
    // entry's area intersecting a component's highlight rect puts that component in scope.
    [Fact]
    public void The_component_scope_includes_components_whose_highlight_rect_the_entrys_area_touches()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Component scope", "");
        Assert.NotNull(workbook);

        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Bad cap", "", "Issue", "Open",
            Array.Empty<string>());
        Assert.NotNull(entry);

        var boardData = new BoardData
        {
            Components = new List<ComponentEntry>
            {
                new() { BoardLabel = "U1", FriendlyName = "Voltage regulator" },
                new() { BoardLabel = "C12", FriendlyName = "Filter cap" }
            }
        };

        var highlightRects = new Dictionary<string, Dictionary<string, List<Avalonia.Rect>>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet 1"] = new(StringComparer.OrdinalIgnoreCase)
            {
                // Overlaps the entry's (0,0,10,10) area - U1 is touched.
                ["U1"] = new List<Avalonia.Rect> { new(5, 5, 10, 10) },
                // Nowhere near the entry's area - C12 is not touched.
                ["C12"] = new List<Avalonia.Rect> { new(500, 500, 10, 10) }
            }
        };

        UiTest.Run(() =>
        {
            var tab = new TabWorkbooks
            {
                CurrentBoardDataOverrideForTests = boardData,
                HighlightRectsBySchematicAndLabelOverrideForTests = highlightRects
            };

            var scope = tab.BuildWorklogEntryComponentScopeForTests(entry!);

            Assert.NotNull(scope);
            var label = Assert.Single(scope!);
            Assert.Equal("U1", label.BoardLabel);
            Assert.Equal("Voltage regulator", label.DisplayName);
        });
    }

    [Fact]
    public void The_component_scope_is_null_with_no_highlight_rect_cache()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No cache", "");
        Assert.NotNull(workbook);

        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Bad cap", "", "Issue", "Open",
            Array.Empty<string>());
        Assert.NotNull(entry);

        UiTest.Run(() =>
        {
            // No HighlightRectsBySchematicAndLabelOverrideForTests set, and no MainWindow either -
            // the same state a headless test (or a real run before Main finishes wiring up) is in.
            var tab = new TabWorkbooks { CurrentBoardDataOverrideForTests = BuildBoardData() };

            Assert.Null(tab.BuildWorklogEntryComponentScopeForTests(entry!));
        });
    }

    [Fact]
    public void The_component_scope_is_null_when_the_schematic_has_no_cached_highlight_rects()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Unknown schematic", "");
        Assert.NotNull(workbook);

        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Bad cap", "", "Issue", "Open",
            Array.Empty<string>());
        Assert.NotNull(entry);

        var highlightRects = new Dictionary<string, Dictionary<string, List<Avalonia.Rect>>>(StringComparer.OrdinalIgnoreCase)
        {
            // Keyed by a DIFFERENT schematic than the entry's own "Sheet 1" - nothing for this
            // entry to match against.
            ["Sheet 2"] = new(StringComparer.OrdinalIgnoreCase)
            {
                ["U1"] = new List<Avalonia.Rect> { new(0, 0, 10, 10) }
            }
        };

        UiTest.Run(() =>
        {
            var tab = new TabWorkbooks
            {
                CurrentBoardDataOverrideForTests = new BoardData { Components = new List<ComponentEntry>() },
                HighlightRectsBySchematicAndLabelOverrideForTests = highlightRects
            };

            Assert.Null(tab.BuildWorklogEntryComponentScopeForTests(entry!));
        });
    }

    // The Legend panel was removed from the board pane entirely - nothing on this tab should
    // still say "Legend" or name a category outside of an actual entry pill/card.
    [Fact]
    public void The_legend_panel_is_gone()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No legend", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Entry", "", "Note", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var legendHeading = tab.GetSelfAndVisualDescendants()
                .OfType<TextBlock>()
                .FirstOrDefault(t => t.Text == "Legend");

            Assert.Null(legendHeading);
        });
    }

    // ------------------------------------------------------- detach / re-attach (tab switching)

    // THE CRASH THIS EXISTS TO PREVENT. A TabControl detaches the previous tab's content on every
    // tab SWITCH, and OnDetachedFromVisualTree disposes this tab's decoded schematic bitmaps. The
    // preview Image controls keep their Source across a detach, so disposing without clearing the
    // pane first left every one of them holding a freed Skia surface - and the next render pass
    // over them threw ObjectDisposedException on the RENDER thread, which is fatal in Avalonia.
    // Reported as the app crashing on switching away from the Workbooks tab.
    //
    // Asserted as "no Image is left holding a Source", which is the property that actually makes it
    // safe: a disposed bitmap that nothing references can never be drawn. Reaching into the Skia
    // surface to prove it is disposed is neither possible from here nor the point.
    [Fact]
    public void Detaching_the_tab_leaves_no_preview_image_holding_a_disposed_bitmap()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Recap", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var window = new Window { Width = 800, Height = 600, Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            // The pane really did build a preview, so the assertion below is not passing merely
            // because there was never anything to strand.
            Assert.NotEmpty(PreviewPanel(tab).Children);
            Assert.Contains(tab.GetSelfAndVisualDescendants().OfType<Image>(), i => i.Source != null);

            // Exactly what a tab switch does: the content is detached from the visual tree.
            window.Content = null;
            Dispatcher.UIThread.RunJobs();

            Assert.DoesNotContain(
                tab.GetSelfAndVisualDescendants().OfType<Image>(),
                image => image.Source != null);

            window.Close();
        });
    }

    // ...and switching BACK rebuilds it, so the fix above does not simply leave the pane
    // permanently empty after the first tab switch.
    [Fact]
    public void Re_attaching_the_tab_rebuilds_the_board_pane()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Recap", "");
        Assert.NotNull(workbook);

        string imagePath = this.WriteSchematicImage("sheet1.png");
        WorklogManager.AddEntry(workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var boardData = BuildBoardData(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = imagePath });

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbook.Id);

            var window = new Window { Width = 800, Height = 600, Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            window.Content = null;
            Dispatcher.UIThread.RunJobs();
            Assert.Empty(PreviewPanel(tab).Children);

            // Back to the Workbooks tab.
            window.Content = tab;
            Dispatcher.UIThread.RunJobs();

            Assert.NotEmpty(PreviewPanel(tab).Children);

            // And the rebuilt preview draws a LIVE bitmap - a re-attach that handed the Image the
            // old disposed instance back would satisfy the count assertion above and still crash.
            Assert.Contains(tab.GetSelfAndVisualDescendants().OfType<Image>(), i => i.Source != null);

            window.Close();
        });
    }

    // A detach with nothing built (the tab was never shown, or its board has no entries) must not
    // throw on the way out - the teardown runs unconditionally.
    [Fact]
    public void Detaching_a_tab_with_no_previews_is_harmless()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No entries yet", "");
        Assert.NotNull(workbook);

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, BuildBoardData(), workbook!.Id);

            var window = new Window { Width = 800, Height = 600, Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            window.Content = null;
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(PreviewPanel(tab).Children);

            window.Close();
        });
    }
}
