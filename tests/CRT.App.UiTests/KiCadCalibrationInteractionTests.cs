using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Align KiCad" as the keys and the pointer drive it, in CRT's real window (owner report,
// 2026-11-06: "incredible slow. It is slow when dragging one of the markers, and it seems not always
// to react when I use the keys for fine-tuning").
//
// Measured on the shipped KiCad data before the change: every arrow key rebuilt the whole overlay
// (75 ms on a fast PC on 250469, 265 ms on the Open128 - and Windows folds a held key's repeats
// that are not taken in time into one, so presses were lost), every drag began and ended with one,
// and every frame of a drag drew all 5,351 primitives. And the keys reached the box only while the
// keyboard was in the Schematics tab. These pin what replaced that:
//   - a key or a drag moves the traces by a matrix and builds nothing - the very same primitives
//     before and after - and only a box resized or flipped is rebuilt, once the input goes quiet;
//   - the frames of a drag, and of a held key, come from a picture of the traces;
//   - the arrow keys reach the box wherever the keyboard is in the window, but in a box typed in.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class KiCadCalibrationInteractionTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public KiCadCalibrationInteractionTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    // A board's top side: one net, two traces on top with a pad at each start.
    private static KiCadProjectBundle Board()
    {
        KiCadNetRef net = new() { Id = "1", Name = "+5V", NormalizedName = "+5V" };

        return new KiCadProjectBundle
        {
            Root = new KiCadProjectRoot
            {
                Project = new KiCadProjectInfo
                {
                    Views = [new KiCadProjectView { DisplayName = "Board - PCB Top", Type = "pcb_top", SourceKind = "pcb", SourceIndex = 0 }],
                },
                Pcb =
                [
                    new KiCadPcb
                    {
                        Nets = new KiCadPcbNets { List = [net] },
                        Footprints =
                        [
                            new KiCadPcbFootprint
                            {
                                Reference = "U1",
                                Pads =
                                [
                                    new KiCadPcbPad { Number = "1", AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 0 }, Layers = ["F.Cu"], Net = net },
                                    new KiCadPcbPad { Number = "2", AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 40 }, Layers = ["F.Cu"], Net = net },
                                ],
                            },
                        ],
                        Routing = new KiCadPcbRouting
                        {
                            Segments =
                            [
                                new KiCadPcbSegment { Start = new KiCadPoint2D { X = 0, Y = 0 }, End = new KiCadPoint2D { X = 80, Y = 0 }, Width = 1, Layer = "F.Cu", Net = net },
                                new KiCadPcbSegment { Start = new KiCadPoint2D { X = 0, Y = 40 }, End = new KiCadPoint2D { X = 80, Y = 40 }, Width = 1, Layer = "F.Cu", Net = net },
                            ],
                        },
                        HighlightIndex = new Dictionary<string, KiCadPcbHighlightBucket>
                        {
                            ["1"] = new KiCadPcbHighlightBucket
                            {
                                Segments = [0, 1],
                                Pads =
                                [
                                    new KiCadPcbHighlightPadRef { FootprintIndex = 0, PadIndex = 0, Reference = "U1", PadNumber = "1" },
                                    new KiCadPcbHighlightPadRef { FootprintIndex = 0, PadIndex = 1, Reference = "U1", PadNumber = "2" },
                                ],
                            },
                        },
                    },
                ],
            },
        };
    }

    // ###########################################################################################
    // CRT's window aligning the board's top: shown, the Schematics tab on it, the keyboard on the
    // picture, and the net's traces drawn (their graph is built off the UI thread - a wait that gives
    // up says so).
    // ###########################################################################################
    private static (CRT.Main Window, TabSchematics Tab) Aligning()
    {
        var window = new CRT.Main { Width = 1300, Height = 800 };
        TabSchematics tab = window.TabSchematicsControl;

        var top = new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png", CadName = "Board - PCB Top" };
        tab.schematicByName["Top"] = top;
        tab.CurrentBoardDataOverrideForTests = new BoardData { Schematics = [top] };

        var thumbnail = new SchematicThumbnail { Name = "Top" };
        tab.currentThumbnails.Add(thumbnail);
        tab.SelectedThumbnail = thumbnail;

        window.Show();
        Dispatcher.UIThread.RunJobs();

        // Its few primitives drawn from a picture as a real board's thousands are.
        tab.SchematicsKiCadOverlayCanvas.PrimitivesWorthAPicture = 1;

        tab.ImageSettledForTests("Top");
        tab.currentFullResBitmap = new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        tab.UseKiCadProjectForTests(Board());
        tab.KiCadCalibrationSaveOverrideForTests = _ => Task.FromResult<string?>(null);

        window.StartPictureTool(PictureTool.AlignKiCad, "Top", readBoardFirst: false);
        tab.KiCadSettled(hasProject: true);
        Assert.True(tab.IsCalibrationModeForTests);

        for (int tries = 0; tries < 400 && Traces(tab).Count < 2; tries++)
        {
            Thread.Sleep(5);
            Dispatcher.UIThread.RunJobs();
        }

        Assert.True(Traces(tab).Count >= 2, $"Only {Traces(tab).Count} of the 2 traces were drawn in time");

        tab.SchematicsContainer.Focus();
        Assert.True(tab.SchematicsContainer.IsFocused);
        return (window, tab);
    }

    private static KiCadOverlayRenderControl Overlay(TabSchematics tab) => tab.SchematicsKiCadOverlayCanvas;

    private static List<KiCadOverlayPrimitive> Traces(TabSchematics tab) =>
        [.. Overlay(tab).Primitives.Where(primitive => primitive.Kind == KiCadOverlayPrimitiveKind.Polyline)];

    // Everything the overlay draws but the calibration box - the primitives a rebuild makes anew.
    private static List<KiCadOverlayPrimitive> Built(TabSchematics tab) =>
        [.. Overlay(tab).Primitives.Where(primitive => !primitive.DrawnUntransformed)];

    private static void AssertTheSameInstances(IReadOnlyList<KiCadOverlayPrimitive> expected, IReadOnlyList<KiCadOverlayPrimitive> actual)
    {
        Assert.Equal(expected.Count, actual.Count);

        for (int i = 0; i < expected.Count; i++)
        {
            Assert.Same(expected[i], actual[i]);
        }
    }

    // The tool ended (no settle left to run once the window has gone), then the window closed.
    private static void End(CRT.Main window, TabSchematics tab)
    {
        tab.TakeKiCadCalibrationKey(Key.Escape, KeyModifiers.None);
        Assert.False(tab.IsKiCadCalibrationSettlePendingForTests);
        window.Close();
    }

    private static void Press(CRT.Main window, Key key, PhysicalKey physical, RawInputModifiers modifiers = RawInputModifiers.None)
    {
        window.KeyPress(key, modifiers, physical, keySymbol: null);
        window.KeyRelease(key, modifiers, physical, keySymbol: null);
        Dispatcher.UIThread.RunJobs();
    }

    // The middle of the calibration box, in the window.
    private static Point MiddleOfTheBox(CRT.Main window, TabSchematics tab)
    {
        var box = tab.KiCadCalibrationBoxForTests;
        Rect content = tab.GetImageContentRect();
        var bitmap = tab.currentFullResBitmap!.PixelSize;

        var local = new Point(
            content.X + ((((box.Left + box.Right) / 2.0) / bitmap.Width) * content.Width),
            content.Y + ((((box.Top + box.Bottom) / 2.0) / bitmap.Height) * content.Height));

        return tab.SchematicsContainer.TranslatePoint(local.Transform(tab.schematicsMatrix), window)!.Value;
    }

    // ###########################################################################################
    // An arrow key moves the box one pixel and its traces by a matrix - the very same primitives, a
    // translation on the overlay - and they are not built again once the keys have gone quiet: a box
    // only moved is drawn exactly as a rebuild would draw it.
    // ###########################################################################################
    [Fact]
    public void A_key_moves_the_box_and_its_traces_without_building_them_again()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();
            var before = tab.KiCadCalibrationBoxForTests;
            List<KiCadOverlayPrimitive> built = Built(tab);

            Press(window, Key.Right, PhysicalKey.ArrowRight);

            var after = tab.KiCadCalibrationBoxForTests;
            Assert.Equal(before.Left + 1.0, after.Left);
            Assert.Equal(before.Right + 1.0, after.Right);
            AssertTheSameInstances(built, Built(tab));

            Matrix moved = Overlay(tab).OverlayTransform;
            Assert.NotEqual(Matrix.Identity, moved);
            Assert.True(KiCadCalibrationTransform.IsTranslationOnly(moved));

            Assert.True(tab.IsKiCadCalibrationSettlePendingForTests);
            tab.SettleKiCadCalibrationOverlayForTests();

            AssertTheSameInstances(built, Built(tab));
            Assert.Equal(moved, Overlay(tab).OverlayTransform);

            End(window, tab);
        });
    }

    // ###########################################################################################
    // NO WORKLOG OVER THE PICTURE WHILE ALIGNING (owner request, 2026-11-07: "When I am aligning KiCad data, then I do
    // not want to see any worklogs on top"): "Show worklogs" ticked, the overlay goes as the alignment starts and stays
    // gone however it is asked for meanwhile, and comes back as the alignment ends.
    // ###########################################################################################
    [Fact]
    public void The_worklogs_are_not_drawn_over_the_picture_while_KiCad_is_aligned_and_come_back_after()
    {
        UiTest.Run(() =>
        {
            WorkbookRecord workbook = WorklogManager.CreateWorkbook("Commodore 64|250407", "No video", "")!;
            WorklogManager.AddEntry(workbook.Id, "Top", new Rect(10, 10, 40, 30), "Check the VIC", "", "Issue", "Open", Array.Empty<string>());

            (CRT.Main window, TabSchematics tab) = Aligning();

            tab.SetShowWorklogEntriesList(true, workbook.Id);
            Assert.False(tab.WorklogOverlayShownForTests);

            tab.RefreshWorklogEntriesListForCurrentWorkbook();
            Assert.False(tab.WorklogOverlayShownForTests);

            tab.TakeKiCadCalibrationKey(Key.Escape, KeyModifiers.None);

            Assert.False(tab.IsAligningKiCad);
            Assert.True(tab.WorklogOverlayShownForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // A key held down: every repeat moves the box, and from the second one the frames come from a
    // picture of the traces - the primitives again once the keys have gone quiet.
    // ###########################################################################################
    [Fact]
    public void A_key_held_down_moves_the_box_on_every_repeat()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();
            var before = tab.KiCadCalibrationBoxForTests;

            for (int repeat = 0; repeat < 10; repeat++)
            {
                window.KeyPress(Key.Down, RawInputModifiers.None, PhysicalKey.ArrowDown, keySymbol: null);
                Dispatcher.UIThread.RunJobs();
            }

            Assert.True(Overlay(tab).IsDrawingFromSnapshot);

            window.KeyRelease(Key.Down, RawInputModifiers.None, PhysicalKey.ArrowDown, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            var after = tab.KiCadCalibrationBoxForTests;
            Assert.Equal(before.Top + 10.0, after.Top);
            Assert.Equal(before.Bottom + 10.0, after.Bottom);

            tab.SettleKiCadCalibrationOverlayForTests();
            Assert.False(Overlay(tab).IsDrawingFromSnapshot);

            End(window, tab);
        });
    }

    // ###########################################################################################
    // Shift+arrow widens the box: the traces scaled by the matrix at once, and BUILT AGAIN once the
    // keys have gone quiet, so their widths are a rebuild's - drawn untransformed again.
    // ###########################################################################################
    [Fact]
    public void Widening_the_box_by_a_key_builds_the_traces_again_once_the_keys_have_gone_quiet()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();
            var before = tab.KiCadCalibrationBoxForTests;
            List<KiCadOverlayPrimitive> built = Built(tab);

            window.KeyPress(Key.LeftShift, RawInputModifiers.None, PhysicalKey.ShiftLeft, keySymbol: null);
            Press(window, Key.Right, PhysicalKey.ArrowRight, RawInputModifiers.Shift);
            window.KeyRelease(Key.LeftShift, RawInputModifiers.None, PhysicalKey.ShiftLeft, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(before.Left, tab.KiCadCalibrationBoxForTests.Left);
            Assert.Equal(before.Right + 1.0, tab.KiCadCalibrationBoxForTests.Right);
            AssertTheSameInstances(built, Built(tab));
            Assert.False(KiCadCalibrationTransform.IsTranslationOnly(Overlay(tab).OverlayTransform));

            tab.SettleKiCadCalibrationOverlayForTests();

            Assert.Equal(Matrix.Identity, Overlay(tab).OverlayTransform);
            Assert.NotSame(built[0], Built(tab)[0]);
            Assert.Equal(built.Count, Built(tab).Count);

            End(window, tab);
        });
    }

    // ###########################################################################################
    // Shift pressed and let go of while aligning rebuilds nothing (nothing is hovered while aligning,
    // so it changes nothing drawn) - it rebuilt the whole overlay at the box's new place each time.
    // ###########################################################################################
    [Fact]
    public void Shift_pressed_and_let_go_of_while_aligning_builds_nothing()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();

            Press(window, Key.Up, PhysicalKey.ArrowUp);
            List<KiCadOverlayPrimitive> built = Built(tab);

            window.KeyPress(Key.LeftShift, RawInputModifiers.Shift, PhysicalKey.ShiftLeft, keySymbol: null);
            Dispatcher.UIThread.RunJobs();
            window.KeyRelease(Key.LeftShift, RawInputModifiers.None, PhysicalKey.ShiftLeft, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            AssertTheSameInstances(built, Built(tab));

            End(window, tab);
        });
    }

    // ###########################################################################################
    // A drag of the box: taking hold of it builds nothing (the box moved by a key first, so a rebuild
    // here would be one at a new calibration), its frames come from a picture, and letting go of a box
    // only moved builds nothing either - the traces sharp again under the matrix the drag ended at.
    // ###########################################################################################
    [Fact]
    public void Dragging_the_box_builds_nothing_and_draws_its_frames_from_a_picture()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();

            Press(window, Key.Left, PhysicalKey.ArrowLeft);
            List<KiCadOverlayPrimitive> built = Built(tab);
            var before = tab.KiCadCalibrationBoxForTests;

            Point start = MiddleOfTheBox(window, tab);
            window.MouseDown(start, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();
            AssertTheSameInstances(built, Built(tab));

            window.MouseMove(new Point(start.X + 30, start.Y + 12));
            Dispatcher.UIThread.RunJobs();

            Assert.True(Overlay(tab).IsDrawingFromSnapshot);
            AssertTheSameInstances(built, Built(tab));

            window.MouseUp(new Point(start.X + 30, start.Y + 12), MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.False(Overlay(tab).IsDrawingFromSnapshot);
            AssertTheSameInstances(built, Built(tab));
            Assert.True(KiCadCalibrationTransform.IsTranslationOnly(Overlay(tab).OverlayTransform));

            var after = tab.KiCadCalibrationBoxForTests;
            Assert.True(after.Left > before.Left);
            Assert.True(after.Top > before.Top);
            Assert.Equal(before.Right - before.Left, after.Right - after.Left, 9);

            tab.SettleKiCadCalibrationOverlayForTests();
            AssertTheSameInstances(built, Built(tab));

            End(window, tab);
        });
    }

    // ###########################################################################################
    // *** THE ARROW KEYS REACH THE BOX WHEREVER THE KEYBOARD IS IN THE WINDOW. *** On a tab header
    // the Right key went to the tab strip: the next tab was shown, and the box stayed where it was.
    // ###########################################################################################
    [Fact]
    public void An_arrow_key_with_the_keyboard_on_the_tab_headers_moves_the_box_and_the_tab_stays()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();
            var before = tab.KiCadCalibrationBoxForTests;

            window.SchematicsTabItem.Focus();
            Assert.False(tab.IsKeyboardFocusWithin);

            Press(window, Key.Right, PhysicalKey.ArrowRight);

            Assert.Equal(before.Left + 1.0, tab.KiCadCalibrationBoxForTests.Left);
            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);

            // Escape there ends the alignment, as on the picture.
            Press(window, Key.Escape, PhysicalKey.Escape);
            Assert.False(tab.IsCalibrationModeForTests);

            End(window, tab);
        });
    }

    // With the keyboard on the picture the tab's own handler takes the key: once, not twice.
    [Fact]
    public void An_arrow_key_with_the_keyboard_on_the_picture_moves_the_box_once()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();
            var before = tab.KiCadCalibrationBoxForTests;

            Press(window, Key.Up, PhysicalKey.ArrowUp);

            Assert.Equal(before.Top - 1.0, tab.KiCadCalibrationBoxForTests.Top);
            Assert.Equal(before.Bottom - 1.0, tab.KiCadCalibrationBoxForTests.Bottom);

            End(window, tab);
        });
    }

    // ###########################################################################################
    // *** AN OPEN DROP-DOWN KEEPS ITS KEYS (code review, 2026-11-08). *** The arrow keys reach the box from anywhere in
    // the window, but in an open drop-down list they move through it and Escape closes it: with the "Board" drop-down
    // opened while aligning, Down moved the box and the list stayed where it was, and Escape ended the alignment - the
    // box's place thrown away - with the list still open.
    // ###########################################################################################
    [Fact]
    public void In_an_open_drop_down_the_keys_are_the_lists_and_the_alignment_goes_on()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();
            var before = tab.KiCadCalibrationBoxForTests;

            ComboBox boards = window.BoardComboBox;
            boards.Focus();
            boards.IsDropDownOpen = true;
            Dispatcher.UIThread.RunJobs();
            Assert.True(boards.IsDropDownOpen);

            Press(window, Key.Down, PhysicalKey.ArrowDown);
            Assert.Equal(before, tab.KiCadCalibrationBoxForTests);

            Press(window, Key.Escape, PhysicalKey.Escape);
            Assert.True(tab.IsCalibrationModeForTests);
            Assert.False(boards.IsDropDownOpen);

            End(window, tab);
        });
    }

    // ###########################################################################################
    // In a box typed in - the component filter - an arrow moves the caret, and the box stays: as the
    // Delete key's rule has it (MainPictureToolsTests), the filter's keys are its own.
    // ###########################################################################################
    [Fact]
    public void An_arrow_key_typed_in_the_component_filter_stays_in_the_filter()
    {
        UiTest.Run(() =>
        {
            (CRT.Main window, TabSchematics tab) = Aligning();
            var before = tab.KiCadCalibrationBoxForTests;

            TextBox filter = window.ComponentSearchTextBox;
            filter.Text = "U12";
            filter.Focus();
            filter.CaretIndex = 1;
            Assert.True(filter.IsFocused);

            Press(window, Key.Right, PhysicalKey.ArrowRight);

            Assert.Equal(2, filter.CaretIndex);
            Assert.Equal(before, tab.KiCadCalibrationBoxForTests);
            Assert.True(tab.IsCalibrationModeForTests);

            End(window, tab);
        });
    }
}
