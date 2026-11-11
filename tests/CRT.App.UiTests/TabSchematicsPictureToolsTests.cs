using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The picture tools on the Schematics tab (TabSchematics.PictureTools.cs -
// CRT 4.0 step 2, slice 6: "Contributor mode goes. Mark components and Align KiCad open the
// Schematics tab in that tool, with a bar and a Done button. Right-clicking a component offers Edit
// this component to everyone.").
//
// A tool asked for waits for its picture's image (and, aligning, the board's KiCad data), and is
// refused with the reason when what it needs can never come; the bar follows the tool; its end is
// told to whoever asked. Images are not decoded headlessly - ImageSettledForTests stands in for the
// end of a picture's load, as the label editor's tests enter past the pointer conversion.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabSchematicsPictureToolsTests
{
    private const string Picture = "Board top";

    private static (TabSchematics Tab, List<string> Ended) CreateTab()
    {
        var tab = new TabSchematics
        {
            CurrentBoardDataOverrideForTests = new BoardData
            {
                Schematics = [new BoardSchematicEntry { SchematicName = Picture }],
                Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" }],
                ComponentHighlights = [new ComponentHighlightEntry { SchematicName = Picture, BoardLabel = "U1", X = "10", Y = "10", Width = "40", Height = "20" }],
            },
        };

        var thumbnail = new SchematicThumbnail { Name = Picture };
        tab.currentThumbnails.Add(thumbnail);
        tab.SelectedThumbnail = thumbnail;

        var ended = new List<string>();
        tab.PictureToolEnded += (_, _) => ended.Add("ended");
        return (tab, ended);
    }

    // Bold pieces or not (BoldRuns: a line with a bold piece has no Text).
    private static string Text(TabSchematics tab, string name) => BoldRuns.TextOf(tab.GetControl<TextBlock>(name));

    private static bool BarShown(TabSchematics tab) => tab.GetControl<ReturnBar>("PictureToolBarBorder").IsVisible;

    private static void Click(TabSchematics tab, string name) =>
        tab.GetControl<Button>(name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    // The keys in the alignment's bar are bold (owner request, 2026-11-03: "e.g. <b>Shift</b>").
    [Fact]
    public void The_alignments_help_says_its_keys_in_bold()
    {
        Assert.Equal(["Arrow keys", "Shift", "Alt", "Esc"], BoldRuns.Marked(PictureToolWording.Help(PictureTool.AlignKiCad)).Where(run => run.IsBold).Select(run => run.Text));
    }

    // ###########################################################################################
    // "Mark components": waiting while the picture is read, then the label editor on it with the bar -
    // and Cancel ends both, the end told once.
    // ###########################################################################################
    [Fact]
    public void Mark_components_waits_for_its_picture_then_starts_the_label_editor_with_the_bar()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateTab();

            tab.StartPictureTool(PictureTool.MarkComponents, Picture);

            Assert.True(BarShown(tab));
            Assert.Equal("Opening [Board top]...", Text(tab, "PictureToolTitleText"));
            Assert.False(tab.GetControl<Button>("PictureToolDoneButton").IsVisible);
            Assert.False(tab.IsLabelEditorModeForTests);

            tab.ImageSettledForTests(Picture);

            Assert.True(tab.IsLabelEditorModeForTests);
            Assert.Equal(Picture, tab.LabelEditorSchematicNameForTests);
            Assert.Equal("Mark components manually on [Board top]", Text(tab, "PictureToolTitleText"));
            Assert.Equal(BoldRuns.Plain(PictureToolWording.Help(PictureTool.MarkComponents)), Text(tab, "PictureToolHelpText"));
            Assert.Equal(["Delete", "D", "Arrow keys", "Ctrl+Z"], BoldRuns.BoldPiecesOf(tab.GetControl<TextBlock>("PictureToolHelpText")));
            Assert.True(tab.GetControl<Button>("PictureToolDoneButton").IsVisible);
            Assert.Empty(ended);

            Click(tab, "PictureToolCancelButton");

            Assert.False(tab.IsLabelEditorModeForTests);
            Assert.False(BarShown(tab));
            Assert.Equal(["ended"], ended);
        });
    }

    // ###########################################################################################
    // Done in "Mark components" goes back too (code review, 2026-10-26): the save reads the board
    // again, and that reload reset the tab - forgetting who asked for the tool - before the bar
    // followed, so its end was never told and CRT stayed on the Schematics tab.
    // ###########################################################################################
    [Fact]
    public async Task Done_in_Mark_components_reports_its_end_though_the_save_reads_the_board_again()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabSchematics tab, List<string> ended) = CreateTab();
            var saved = new List<string>();
            tab.LabelEditorSaveOverrideForTests = schematic =>
            {
                saved.Add(schematic);
                return Task.FromResult<string?>(null);
            };

            // What Main's reload does at once: the viewer reset for the board read again.
            tab.ReloadBoardOverrideForTests = _ => tab.ResetSchematicsViewer();

            tab.StartPictureTool(PictureTool.MarkComponents, Picture);
            tab.ImageSettledForTests(Picture);
            Assert.True(tab.IsLabelEditorModeForTests);

            Click(tab, "PictureToolDoneButton");

            for (int wait = 0; wait < 100 && tab.IsLabelEditorModeForTests; wait++)
            {
                await Task.Delay(10);
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
            }

            Assert.Equal([Picture], saved);
            Assert.False(tab.IsLabelEditorModeForTests);
            Assert.False(BarShown(tab));
            Assert.Equal(["ended"], ended);
        });
    }

    // A picture the board does not have is refused at once - the reason on the bar, Close ends it.
    [Fact]
    public void A_picture_the_board_does_not_have_is_refused_with_the_reason()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateTab();

            tab.StartPictureTool(PictureTool.MarkComponents, "Board bottom");

            Assert.Equal("Components cannot be marked on [Board bottom]", Text(tab, "PictureToolTitleText"));
            Assert.Equal(PictureToolWording.PictureNotFound("Board bottom"), Text(tab, "PictureToolHelpText"));
            Assert.Equal(PictureToolWording.Close, tab.GetControl<Button>("PictureToolCancelButton").Content);
            Assert.False(tab.IsLabelEditorModeForTests);

            Click(tab, "PictureToolCancelButton");

            Assert.False(BarShown(tab));
            Assert.Equal(["ended"], ended);
        });
    }

    // "Align KiCad" waits for the board's KiCad data as well, and is refused once there is none.
    [Fact]
    public void Align_KiCad_waits_for_the_KiCad_data_and_is_refused_on_a_board_without_it()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateTab();

            tab.StartPictureTool(PictureTool.AlignKiCad, Picture);
            tab.ImageSettledForTests(Picture);

            Assert.Equal("Opening [Board top]...", Text(tab, "PictureToolTitleText"));

            tab.KiCadSettled(hasProject: false);

            Assert.False(tab.IsCalibrationModeForTests);
            Assert.Equal("KiCad data cannot be aligned on [Board top]", Text(tab, "PictureToolTitleText"));
            Assert.Equal(PictureToolWording.NoKiCadData, Text(tab, "PictureToolHelpText"));
            Assert.Empty(ended);
        });
    }

    // Waiting can be called off: Cancel lets the request go, and its end is told.
    [Fact]
    public void A_tool_waiting_for_its_picture_can_be_cancelled()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateTab();

            tab.StartPictureTool(PictureTool.MarkComponents, Picture);
            Click(tab, "PictureToolCancelButton");
            tab.ImageSettledForTests(Picture);

            Assert.False(tab.IsLabelEditorModeForTests);
            Assert.False(BarShown(tab));
            Assert.Equal(["ended"], ended);
        });
    }

    // ###########################################################################################
    // The label editor started any other way shows the bar too - but its end is told only for a tool
    // somebody asked for (the Contribute tab goes back only then).
    // ###########################################################################################
    [Fact]
    public void The_bar_follows_the_label_editor_and_only_an_asked_for_tool_reports_its_end()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateTab();

            tab.BeginLabelEditorModeForTests();
            Assert.True(BarShown(tab));

            tab.CancelLabelEditorChangesForTests();
            Assert.False(BarShown(tab));
            Assert.Empty(ended);
        });
    }

    // ---- "Done - next picture" (step 3, slice 3d) -----------------------------------------------------

    // ###########################################################################################
    // A board of three pictures - its top and bottom with their KiCad views, a layout drawing with
    // none - and KiCad data with the two views: calibration can start on either side, a stand-in
    // bitmap for the picture shown (images are not decoded headlessly).
    // ###########################################################################################
    private static (TabSchematics Tab, List<string> Ended) CreateAligningTab()
    {
        BoardSchematicEntry[] schematics =
        [
            new() { SchematicName = "Top", SchematicImageFile = "top.png", CadName = "Board - PCB Top" },
            new() { SchematicName = "Layout", SchematicImageFile = "layout.png" },
            new() { SchematicName = "Bottom", SchematicImageFile = "bottom.png", CadName = "Board - PCB Bottom" },
        ];

        var tab = new TabSchematics { CurrentBoardDataOverrideForTests = new BoardData { Schematics = [.. schematics] } };

        foreach (BoardSchematicEntry schematic in schematics)
        {
            tab.schematicByName[schematic.SchematicName] = schematic;
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = schematic.SchematicName });
        }

        tab.currentFullResBitmap = new Avalonia.Media.Imaging.WriteableBitmap(
            new Avalonia.PixelSize(200, 100),
            new Avalonia.Vector(96, 96),
            Avalonia.Platform.PixelFormat.Bgra8888,
            Avalonia.Platform.AlphaFormat.Premul);

        tab.UseKiCadProjectForTests(new KiCadProjectBundle
        {
            Root = new KiCadProjectRoot
            {
                Project = new KiCadProjectInfo
                {
                    Views =
                    [
                        new KiCadProjectView { DisplayName = "Board - PCB Top", Type = "pcb_top", SourceKind = "pcb", SourceIndex = 0 },
                        new KiCadProjectView { DisplayName = "Board - PCB Bottom", Type = "pcb_bottom", SourceKind = "pcb", SourceIndex = 0 },
                    ],
                },
                Pcb = [new KiCadPcb()],
            },
        });

        var ended = new List<string>();
        tab.PictureToolEnded += (_, _) => ended.Add("ended");
        return (tab, ended);
    }

    // "Align KiCad" on a picture: its image read and the KiCad data settled - calibration on.
    private static void Align(TabSchematics tab, string picture)
    {
        tab.StartPictureTool(PictureTool.AlignKiCad, picture);
        tab.KiCadSettled(hasProject: true);
        tab.ImageSettledForTests(picture);

        Assert.True(tab.IsCalibrationModeForTests);
        Assert.Equal(picture, tab.GetCurrentSchematicName());
    }

    // ###########################################################################################
    // *** A SHEET BEING ALIGNED SHOWS EVERY WIRE, NAMED OR NOT (owner report, 2026-11-04: "it has no traces
    // in the alignment that I can align to" - of a sheet of 221 wires). *** Going by the nets drew only the
    // wires a name could be found for: here, none. The bounds are the wires', so a primitive per wire.
    // ###########################################################################################
    [Fact]
    public void Aligning_a_sheet_draws_every_wire_whether_or_not_a_net_is_known_for_it()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, _) = CreateAligningTab();

            var sheet = new BoardSchematicEntry { SchematicName = "Power", SchematicImageFile = "power.png", CadName = "Board - Power" };
            tab.CurrentBoardDataOverrideForTests = new BoardData { Schematics = [sheet] };
            tab.schematicByName[sheet.SchematicName] = sheet;
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = sheet.SchematicName });

            static KiCadSchematicPathItem Wire(double fromX, double fromY, double toX, double toY) => new()
            {
                Points = [new KiCadPoint2D { X = fromX, Y = fromY }, new KiCadPoint2D { X = toX, Y = toY }],
            };

            tab.UseKiCadProjectForTests(new KiCadProjectBundle
            {
                Root = new KiCadProjectRoot
                {
                    Project = new KiCadProjectInfo
                    {
                        Views = [new KiCadProjectView { DisplayName = "Board - Power", Type = "schematic", SourceKind = "schematic", SourceIndex = 0 }],
                    },
                    Schematics =
                    [
                        new KiCadSchematic
                        {
                            Wires = [Wire(0, 0, 100, 0), Wire(0, 50, 100, 50)],
                            Polylines = [Wire(10, 10, 20, 20)],
                        },
                    ],
                },

                // The sheet's nets: none, as the Open128's "Power" sheet had before its power symbols were read.
                SchematicNetPathIndexBySchematicIndex = { [0] = new Dictionary<string, List<KiCadResolvedPath>>(StringComparer.OrdinalIgnoreCase) },
            });

            var window = new Window { Width = 1000, Height = 700, Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            KiCadOverlayRenderControl overlay = tab.GetControl<KiCadOverlayRenderControl>("SchematicsKiCadOverlayCanvas");

            Align(tab, "Power");
            Dispatcher.UIThread.RunJobs();
            tab.RefreshKiCadOverlayForTests();

            Assert.Equal(3, overlay.Primitives.Count(primitive => primitive.Kind == KiCadOverlayPrimitiveKind.Polyline));

            window.Close();
        });
    }

    // ###########################################################################################
    // ALIGNING SHOWS ONLY THE SIDE BEING ALIGNED (owner request, 2026-11-03: "if I am aligning the top PCB,
    // then I only want to see traces from the top"), whatever "Show opposite side traces" says - it is on
    // here (the default), so the other side's copper would be drawn too. Two traces on top, one on the
    // bottom, so the count says which side is drawn. The net's graph is built off the UI thread, so the
    // lines arrive a moment later - a wait that gives up says so.
    // ###########################################################################################
    [Fact]
    public void Aligning_draws_only_the_traces_of_the_side_being_aligned()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, _) = CreateAligningTab();

            KiCadNetRef net = new() { Id = "1", Name = "+5V", NormalizedName = "+5V" };

            tab.UseKiCadProjectForTests(new KiCadProjectBundle
            {
                Root = new KiCadProjectRoot
                {
                    Project = new KiCadProjectInfo
                    {
                        Views =
                        [
                            new KiCadProjectView { DisplayName = "Board - PCB Top", Type = "pcb_top", SourceKind = "pcb", SourceIndex = 0 },
                            new KiCadProjectView { DisplayName = "Board - PCB Bottom", Type = "pcb_bottom", SourceKind = "pcb", SourceIndex = 0 },
                        ],
                    },
                    Pcb =
                    [
                        new KiCadPcb
                        {
                            Nets = new KiCadPcbNets { List = [net] },

                            // A net is drawn from its pads outward (BuildKiCadPcbActiveDrawIds), so a pad at the
                            // start of each trace.
                            Footprints =
                            [
                                new KiCadPcbFootprint
                                {
                                    Reference = "U1",
                                    Pads =
                                    [
                                        new KiCadPcbPad { Number = "1", AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 0 }, Layers = ["F.Cu"], Net = net },
                                        new KiCadPcbPad { Number = "2", AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 25 }, Layers = ["F.Cu"], Net = net },
                                        new KiCadPcbPad { Number = "3", AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 50 }, Layers = ["B.Cu"], Net = net },
                                    ],
                                },
                            ],
                            Routing = new KiCadPcbRouting
                            {
                                Segments =
                                [
                                    new KiCadPcbSegment { Start = new KiCadPoint2D { X = 0, Y = 0 }, End = new KiCadPoint2D { X = 50, Y = 0 }, Width = 1, Layer = "F.Cu", Net = net },
                                    new KiCadPcbSegment { Start = new KiCadPoint2D { X = 0, Y = 25 }, End = new KiCadPoint2D { X = 50, Y = 25 }, Width = 1, Layer = "F.Cu", Net = net },
                                    new KiCadPcbSegment { Start = new KiCadPoint2D { X = 0, Y = 50 }, End = new KiCadPoint2D { X = 50, Y = 50 }, Width = 1, Layer = "B.Cu", Net = net },
                                ],
                            },
                            HighlightIndex = new Dictionary<string, KiCadPcbHighlightBucket>
                            {
                                ["1"] = new KiCadPcbHighlightBucket
                                {
                                    Segments = [0, 1, 2],
                                    Pads =
                                    [
                                        new KiCadPcbHighlightPadRef { FootprintIndex = 0, PadIndex = 0, Reference = "U1", PadNumber = "1" },
                                        new KiCadPcbHighlightPadRef { FootprintIndex = 0, PadIndex = 1, Reference = "U1", PadNumber = "2" },
                                        new KiCadPcbHighlightPadRef { FootprintIndex = 0, PadIndex = 2, Reference = "U1", PadNumber = "3" },
                                    ],
                                },
                            },
                        },
                    ],
                },
            });

            var window = new Window { Width = 1000, Height = 700, Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            KiCadOverlayRenderControl overlay = tab.GetControl<KiCadOverlayRenderControl>("SchematicsKiCadOverlayCanvas");

            int Traces() => overlay.Primitives.Count(primitive => primitive.Kind is KiCadOverlayPrimitiveKind.Line or KiCadOverlayPrimitiveKind.Polyline);

            int TracesDrawnOnceSettled(int expected)
            {
                for (int tries = 0; tries < 400 && Traces() < expected; tries++)
                {
                    Thread.Sleep(5);
                    Dispatcher.UIThread.RunJobs();
                    tab.RefreshKiCadOverlayForTests();
                }

                Assert.True(Traces() >= expected, $"Only {Traces()} of the {expected} traces were drawn in time");

                // The other side's graph, were it asked for, lands in the next look.
                Thread.Sleep(200);
                Dispatcher.UIThread.RunJobs();
                tab.RefreshKiCadOverlayForTests();

                return Traces();
            }

            Align(tab, "Top");
            Assert.Equal(2, TracesDrawnOnceSettled(2));

            Click(tab, "PictureToolCancelButton");

            Align(tab, "Bottom");
            Assert.Equal(1, TracesDrawnOnceSettled(1));

            window.Close();
        });
    }

    // ###########################################################################################
    // ALIGNING DRAWS NO COPPER POUR (owner request, 2026-11-05: "NOT include any ground plane, as it is hard
    // to align those nets"): a filled zone hides the traces under it and has no edge worth lining up, whatever
    // "Show zones" says - it is on here (the default). Measured on the shipped KiCad data, the pours are on
    // "GND" or "Earth" and on "+5V" too, so it goes by what is drawn, never by the net's name. Its tracks and
    // pads stay, and the zone is back once the tool has ended.
    // ###########################################################################################
    [Fact]
    public void Aligning_draws_the_tracks_and_pads_of_a_ground_net_but_not_its_copper_pour()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, _) = CreateAligningTab();

            KiCadNetRef ground = new() { Id = "1", Name = "GND", NormalizedName = "GND" };

            tab.UseKiCadProjectForTests(new KiCadProjectBundle
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
                            Nets = new KiCadPcbNets { List = [ground] },
                            Footprints =
                            [
                                new KiCadPcbFootprint
                                {
                                    Reference = "U1",
                                    Pads = [new KiCadPcbPad { Number = "1", AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 0 }, Layers = ["F.Cu"], Net = ground }],
                                },
                            ],
                            Routing = new KiCadPcbRouting
                            {
                                Segments = [new KiCadPcbSegment { Start = new KiCadPoint2D { X = 0, Y = 0 }, End = new KiCadPoint2D { X = 50, Y = 0 }, Width = 1, Layer = "F.Cu", Net = ground }],
                                Zones =
                                [
                                    new KiCadPcbZone
                                    {
                                        Layers = ["F.Cu"],
                                        Net = ground,
                                        FilledPolygons =
                                        [
                                            new KiCadPcbZonePolygon
                                            {
                                                Points =
                                                [
                                                    new KiCadPoint2D { X = 0, Y = 0 },
                                                    new KiCadPoint2D { X = 100, Y = 0 },
                                                    new KiCadPoint2D { X = 100, Y = 100 },
                                                    new KiCadPoint2D { X = 0, Y = 100 },
                                                ],
                                            },
                                        ],
                                    },
                                ],
                            },
                            HighlightIndex = new Dictionary<string, KiCadPcbHighlightBucket>
                            {
                                ["1"] = new KiCadPcbHighlightBucket
                                {
                                    Segments = [0],
                                    Zones = [0],
                                    Pads = [new KiCadPcbHighlightPadRef { FootprintIndex = 0, PadIndex = 0, Reference = "U1", PadNumber = "1" }],
                                },
                            },
                        },
                    ],
                },
            });

            var window = new Window { Width = 1000, Height = 700, Content = tab };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            KiCadOverlayRenderControl overlay = tab.GetControl<KiCadOverlayRenderControl>("SchematicsKiCadOverlayCanvas");

            int Traces() => overlay.Primitives.Count(primitive => primitive.Kind is KiCadOverlayPrimitiveKind.Line or KiCadOverlayPrimitiveKind.Polyline);
            int Shapes() => overlay.Primitives.Count(primitive => primitive.Kind == KiCadOverlayPrimitiveKind.Geometry);

            Align(tab, "Top");

            // The net's graph is built off the UI thread, so the lines arrive a moment later.
            for (int tries = 0; tries < 400 && Traces() < 1; tries++)
            {
                Thread.Sleep(5);
                Dispatcher.UIThread.RunJobs();
                tab.RefreshKiCadOverlayForTests();
            }

            Assert.True(Traces() >= 1, "The ground net's track was not drawn in time");

            // The pad's shape alone - it was two with the pour's.
            Assert.Equal(1, Shapes());

            window.Close();
        });
    }

    // ###########################################################################################
    // The plan's "Align: opens calibration on the Schematics tab with a Done - next picture button":
    // aligning the top, the bar offers the next picture with a view - the bottom, past the layout
    // drawing, which has none. Pressed: the top's alignment saved, the bottom asked for at once - no
    // end told in between, so CRT stays on the tab. On the last picture there is no next: Done ends
    // it, and that end is told.
    // ###########################################################################################
    [Fact]
    public void Done_next_picture_saves_this_alignment_and_aligns_the_next_picture()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateAligningTab();
            var saved = new List<string>();
            tab.KiCadCalibrationSaveOverrideForTests = row =>
            {
                saved.Add(row.SchematicName);
                return Task.FromResult<string?>(null);
            };
            tab.AlignedPicturesOverrideForTests = () => saved;

            Align(tab, "Top");

            Button next = tab.GetControl<Button>("PictureToolDoneNextButton");
            Assert.True(next.IsVisible);
            Assert.Equal(PictureToolWording.DoneNext, next.Content);
            Assert.Contains("Accept", next.Classes);
            Assert.Equal(PictureToolWording.DoneNextTip("Bottom"), ToolTip.GetTip(next));

            Click(tab, "PictureToolDoneNextButton");

            // Saved, and the bottom selected and waited for - the bar on it, no end told.
            Assert.Equal(["Top"], saved);
            Assert.False(tab.IsCalibrationModeForTests);
            Assert.Equal("Bottom", tab.GetCurrentSchematicName());
            Assert.Equal("Opening [Bottom]...", Text(tab, "PictureToolTitleText"));
            Assert.True(BarShown(tab));
            Assert.Empty(ended);

            tab.ImageSettledForTests("Bottom");

            Assert.True(tab.IsCalibrationModeForTests);
            Assert.Equal("Align KiCad data on [Bottom]", Text(tab, "PictureToolTitleText"));

            // The top is aligned and the layout has no view: nothing after the bottom.
            Assert.False(next.IsVisible);

            Click(tab, "PictureToolDoneButton");

            Assert.Equal(["Top", "Bottom"], saved);
            Assert.False(tab.IsCalibrationModeForTests);
            Assert.False(BarShown(tab));
            Assert.Equal(["ended"], ended);
        });
    }

    // ###########################################################################################
    // With Main there, the next picture is MAIN's to start (AlignNextRequested): a board read again
    // that waited for the tool to end is read first - run as the tool went off, its reset forgot the
    // request halfway. Meanwhile the bar says the next is opening, and no end is told.
    // ###########################################################################################
    [Fact]
    public void With_a_host_the_next_picture_is_asked_of_it()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateAligningTab();
            tab.KiCadCalibrationSaveOverrideForTests = _ => Task.FromResult<string?>(null);
            tab.AlignedPicturesOverrideForTests = () => [];
            var asked = new List<string>();
            tab.AlignNextRequested += (_, picture) => asked.Add(picture);

            Align(tab, "Top");
            Click(tab, "PictureToolDoneNextButton");

            Assert.Equal(["Bottom"], asked);
            Assert.False(tab.IsCalibrationModeForTests);
            Assert.Equal("Opening [Bottom]...", Text(tab, "PictureToolTitleText"));
            Assert.Equal("Top", tab.GetCurrentSchematicName());
            Assert.Empty(ended);
        });
    }

    // ###########################################################################################
    // A save that fails leaves calibration on, on THIS picture - nothing the user set up is lost, and
    // the next picture is not gone on to. Nor is it told as aligned: nothing is marked from it.
    // ###########################################################################################
    [Fact]
    public void A_failed_save_stays_on_this_picture()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateAligningTab();
            tab.KiCadCalibrationSaveOverrideForTests = _ => Task.FromResult<string?>("The draft could not be saved.");
            tab.AlignedPicturesOverrideForTests = () => [];
            var asked = new List<string>();
            tab.AlignNextRequested += (_, picture) => asked.Add(picture);
            var aligned = new List<string>();
            tab.AlignedFromWorkspace += (_, picture) => aligned.Add(picture.Picture);

            Align(tab, "Top");
            Click(tab, "PictureToolDoneNextButton");

            Assert.True(tab.IsCalibrationModeForTests);
            Assert.Equal("Top", tab.GetCurrentSchematicName());
            Assert.Equal("Align KiCad data on [Top]", Text(tab, "PictureToolTitleText"));
            Assert.Empty(asked);
            Assert.Empty(ended);
            Assert.Empty(aligned);
        });
    }

    // ###########################################################################################
    // *** EVERY PICTURE ALIGNED WITH DONE IS TOLD, BEFORE THE TOOL'S END (owner request, 2026-11-01:
    // Done marks the components - Main marks each picture told once the tool hands back). *** With
    // "Done - next picture" and with Done, each with its image file; Cancel tells nothing.
    // ###########################################################################################
    [Fact]
    public void Each_picture_aligned_with_Done_is_told_before_the_tools_end()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, List<string> ended) = CreateAligningTab();
            var saved = new List<string>();
            tab.KiCadCalibrationSaveOverrideForTests = row =>
            {
                saved.Add(row.SchematicName);
                return Task.FromResult<string?>(null);
            };
            tab.AlignedPicturesOverrideForTests = () => saved;
            tab.AlignedFromWorkspace += (_, picture) => ended.Add($"aligned {picture.Picture} {picture.ImageFile}");

            Align(tab, "Top");
            Click(tab, "PictureToolDoneNextButton");
            tab.ImageSettledForTests("Bottom");
            Click(tab, "PictureToolDoneButton");

            Assert.Equal(["aligned Top top.png", "aligned Bottom bottom.png", "ended"], ended);

            // Aligned again and cancelled: nothing told but the end.
            ended.Clear();
            Align(tab, "Top");
            Click(tab, "PictureToolCancelButton");

            Assert.Equal(["ended"], ended);
        });
    }

    // ###########################################################################################
    // Mark components has no next picture - the button is the alignment's alone.
    // ###########################################################################################
    [Fact]
    public void Mark_components_offers_no_next_picture()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, _) = CreateAligningTab();

            tab.StartPictureTool(PictureTool.MarkComponents, "Top");
            tab.ImageSettledForTests("Top");

            Assert.True(tab.IsLabelEditorModeForTests);
            Assert.False(tab.GetControl<Button>("PictureToolDoneNextButton").IsVisible);
        });
    }

    // ###########################################################################################
    // "Back to Contribute" (owner request, 2026-10-29): the bar says what was shown from the
    // Contribute tab and its button asks to go back. A picture tool's bar takes its place - the tool
    // goes back by itself - and it comes back once the tool's bar goes.
    // ###########################################################################################
    [Fact]
    public void Back_to_Contribute_says_what_was_shown_and_gives_way_to_a_tools_bar()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, _) = CreateTab();
            int asked = 0;
            tab.BackToContributeRequested += (_, _) => asked++;

            Assert.False(tab.IsBackToContributeShownForTests);

            tab.ShowBackToContribute(PictureToolWording.ComponentShown(" U1 "));
            Assert.True(tab.IsBackToContributeShownForTests);
            Assert.Equal("Showing [U1] from the \"Contribute\" tab", tab.BackToContributeTextForTests);
            Assert.Equal(PictureToolWording.BackToContribute, tab.GetControl<Button>("BackToContributeButton").Content);

            tab.PressBackToContributeForTests();
            Assert.Equal(1, asked);

            tab.StartPictureTool(PictureTool.MarkComponents, Picture);
            Assert.True(BarShown(tab));
            Assert.False(tab.IsBackToContributeShownForTests);

            Click(tab, "PictureToolCancelButton");
            Assert.True(tab.IsBackToContributeShownForTests);

            tab.ShowBackToContribute(null);
            Assert.False(tab.IsBackToContributeShownForTests);
        });
    }

    // ###########################################################################################
    // Every bar that takes a person back to the Contribute tab looks the same (owner request, 2026-11-11: "This panel should be
    // identical everywhere, and please let us try and make this lightgrey background with 2px black dashed border, to give it
    // some attention" - then white in place of the lightgray): "Back to Contribute" and the picture tool's bar are both a ReturnBar - white, framed by a 2px black
    // dash - and their words are black on it whichever theme is on.
    // ###########################################################################################
    [Fact]
    public void Every_bar_back_to_Contribute_is_white_with_a_2px_black_dashed_frame()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, _) = CreateTab();

            // Shown: the colors are the app's resources, which a control finds only in a window.
            var window = new Window { Content = tab, Width = 1000, Height = 700 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            foreach (string name in new[] { "BackToContributeBorder", "PictureToolBarBorder" })
            {
                ReturnBar bar = tab.GetControl<ReturnBar>(name);

                Assert.Equal(Avalonia.Media.Colors.White, ((Avalonia.Media.ISolidColorBrush)bar.Background!).Color);
                Assert.Equal(Avalonia.Media.Colors.Black, ((Avalonia.Media.ISolidColorBrush)bar.FrameBrush!).Color);
            }

            Assert.Equal(2, ReturnBar.FrameThickness);
            Assert.Equal(Avalonia.Media.Colors.Black, ((Avalonia.Media.ISolidColorBrush)tab.GetControl<TextBlock>("BackToContributeText").Foreground!).Color);
            Assert.Equal(Avalonia.Media.Colors.Black, ((Avalonia.Media.ISolidColorBrush)tab.GetControl<TextBlock>("PictureToolTitleText").Foreground!).Color);

            window.Close();
        });
    }
}
