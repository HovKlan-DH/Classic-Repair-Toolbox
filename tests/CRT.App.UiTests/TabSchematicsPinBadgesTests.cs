using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The pins of a component on the "Schematics" tab's schematic image (owner requests, 2026-11-17).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabSchematicsPinBadgesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public TabSchematicsPinBadgesTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    // A board's top with one 74LS08, U5, in a DIP-14 footprint: pins 1 to 7 down the left, 8 to 14 up the right, in mm.
    private static KiCadProjectBundle Project()
    {
        var pads = new List<KiCadPcbPad>();

        for (int pin = 1; pin <= 14; pin++)
        {
            pads.Add(new KiCadPcbPad
            {
                Number = pin.ToString(System.Globalization.CultureInfo.InvariantCulture),
                AbsoluteCenter = pin <= 7 ? new KiCadPoint2D { X = 20, Y = 10 + ((pin - 1) * 2.54) } : new KiCadPoint2D { X = 27.62, Y = 10 + ((14 - pin) * 2.54) },
                Size = new KiCadSize2D { X = 1.6, Y = 1.6 },
                Layers = ["F.Cu", "B.Cu"],
            });
        }

        return new KiCadProjectBundle
        {
            Root = new KiCadProjectRoot
            {
                Project = new KiCadProjectInfo
                {
                    Views = [new KiCadProjectView { DisplayName = "Board - PCB Top", Type = "pcb_top", SourceKind = "pcb", SourceIndex = 0 }],
                },
                // Two mounting holes at the board's corners, so the board (80 x 50 mm) is mapped onto the 800 x 500 picture
                // alike both ways, as an aligned board is - the footprint's own proportions kept.
                Pcb =
                [
                    new KiCadPcb
                    {
                        Footprints =
                        [
                            new KiCadPcbFootprint { Reference = "U5", Pads = pads },
                            new KiCadPcbFootprint
                            {
                                Reference = "H1",
                                Pads =
                                [
                                    new KiCadPcbPad { Number = "1", AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 0 }, Layers = ["F.Cu"] },
                                    new KiCadPcbPad { Number = "2", AbsoluteCenter = new KiCadPoint2D { X = 80, Y = 50 }, Layers = ["F.Cu"] },
                                ],
                            },
                        ],
                    },
                ],
            },
        };
    }

    // The tab in a window, its picture of the board's top shown and the board's KiCad data settled.
    private static (Window Window, TabSchematics Tab) OnTheBoardsTop(BoardData? onTheBoard = null)
    {
        var tab = new TabSchematics();
        var window = new Window { Width = 1200, Height = 800, Content = tab };

        var top = new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png", CadName = "Board - PCB Top" };
        tab.schematicByName["Top"] = top;

        BoardData board = onTheBoard ?? Board();
        tab.CurrentBoardDataOverrideForTests = new BoardData { Schematics = [top], Components = board.Components, ComponentPins = board.ComponentPins };

        var thumbnail = new SchematicThumbnail { Name = "Top" };
        tab.currentThumbnails.Add(thumbnail);
        tab.SelectedThumbnail = thumbnail;

        window.Show();
        Dispatcher.UIThread.RunJobs();

        tab.ImageSettledForTests("Top");
        tab.currentFullResBitmap = new WriteableBitmap(new PixelSize(800, 500), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        tab.UseKiCadProjectForTests(Project());
        tab.KiCadSettled(hasProject: true);
        Dispatcher.UIThread.RunJobs();

        return (window, tab);
    }

    // ###########################################################################################
    // *** EVERY PIN NUMBERED ON ITS PAD WHILE SHIFT IS HELD (owner request, 2026-11-17: "Number all pins holding SHIFT") - the
    // components selected or pointed at, those whose pin 1 is marked. *** Pin 1 told apart; each number in its pin's
    // direction's colors, the pinout's - or plain with "Show color for pin direction" unticked.
    //
    // AND WHILE THEY ARE IN THEIR DIRECTIONS' COLORS, A LEGEND SAYS WHAT THE COLORS MEAN (owner question, the same day: "If I
    // have selected "Show color for pin direction" then I think we need a legend somewhere that can explain the colors?"):
    // every direction in the pinout's words, and pin 1's ring - not while the numbers are plain, nor while none is shown.
    // ###########################################################################################
    [Fact]
    public void A_selected_chips_pins_are_numbered_holding_shift_with_a_legend_for_their_colors()
    {
        UiTest.Run(() =>
        {
            (Window window, TabSchematics tab) = OnTheBoardsTop();
            PinNumbersOverlay numbers = tab.GetControl<PinNumbersOverlay>("SchematicsPinNumbersOverlay");
            Border legend = tab.GetControl<Border>("PinDirectionLegend");

            tab.UpdateHighlightsForComponents(["U5"]);
            Dispatcher.UIThread.RunJobs();

            // Not holding SHIFT: no numbers, no legend.
            Assert.Empty(numbers.Badges);
            Assert.False(legend.IsVisible);

            // Holding SHIFT: every pad numbered, in its direction's colors, and the legend saying what they mean.
            tab.NoteKeyModifiers(KeyModifiers.Shift);

            Assert.Equal(14, numbers.Badges.Count);
            Assert.Equal(["1"], numbers.Badges.Where(badge => badge.IsFirst).Select(badge => badge.Number));
            Assert.Equal(PinDirectionKind.Input, numbers.Badges.Single(badge => badge.Number == "1").Direction);
            Assert.Equal(PinDirectionKind.Output, numbers.Badges.Single(badge => badge.Number == "3").Direction);
            Assert.Equal(PinDirectionKind.Other, numbers.Badges.Single(badge => badge.Number == "7").Direction);

            Assert.True(legend.IsVisible);
            Assert.Equal(["Input", "Output", "Input/output", "Other", "Pin 1"], tab.PinDirectionLegendWordsForTests);
            Assert.Equal(
                [.. Handlers.Theming.PinDirectionColors.LegendOrder.Select(Handlers.Theming.PinDirectionColors.WordOf), TabSchematics.LegendFirstPin],
                tab.PinDirectionLegendWordsForTests);

            // Each number on its pad: pin 8 at the bottom right, below and right of pin 1.
            PinNumberBadge first = numbers.Badges.Single(badge => badge.Number == "1");
            PinNumberBadge eighth = numbers.Badges.Single(badge => badge.Number == "8");
            Assert.True(eighth.Center.X > first.Center.X && eighth.Center.Y > first.Center.Y, $"pin 1 at {first.Center}, pin 8 at {eighth.Center}");

            // "Show color for pin direction" unticked: the numbers plain, and nothing for a legend to explain.
            tab.GetControl<CheckBox>("CheckGlobalShowPinDirection").IsChecked = false;
            Assert.All(numbers.Badges, badge => Assert.Null(badge.Direction));
            Assert.False(UserSettings.SchematicsShowPinDirection);
            Assert.False(legend.IsVisible);

            tab.GetControl<CheckBox>("CheckGlobalShowPinDirection").IsChecked = true;
            Assert.True(legend.IsVisible);

            // SHIFT let go of: the numbers go, and the legend with them.
            tab.NoteKeyModifiers(KeyModifiers.None);
            Assert.Empty(numbers.Badges);
            Assert.False(legend.IsVisible);

            // The numbers unticked: SHIFT draws nothing.
            tab.GetControl<CheckBox>("CheckGlobalNumberPins").IsChecked = false;
            tab.NoteKeyModifiers(KeyModifiers.Shift);
            Assert.Empty(numbers.Badges);
            Assert.False(legend.IsVisible);

            // Nothing selected: nothing drawn, holding SHIFT or not.
            tab.GetControl<CheckBox>("CheckGlobalNumberPins").IsChecked = true;
            tab.UpdateHighlightsForComponents([]);
            Dispatcher.UIThread.RunJobs();
            Assert.Empty(numbers.Badges);
            Assert.False(legend.IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // SHIFT numbers no pin while a tool draws on the image (code review, 2026-11-18): the label editor's SHIFT turns its
    // snapping off, and the numbers - over its overlay - covered the rectangle being drawn; "Draw trace" and a worklog's area
    // are drawn there too.
    // ###########################################################################################
    [Fact]
    public void Shift_numbers_no_pin_while_a_tool_draws_on_the_image()
    {
        UiTest.Run(() =>
        {
            (Window window, TabSchematics tab) = OnTheBoardsTop();
            PinNumbersOverlay numbers = tab.GetControl<PinNumbersOverlay>("SchematicsPinNumbersOverlay");

            tab.UpdateHighlightsForComponents(["U5"]);
            Dispatcher.UIThread.RunJobs();

            tab.SetDrawingTraces(true);
            Assert.True(tab.IsDrawingTraces);
            tab.NoteKeyModifiers(KeyModifiers.Shift);
            Assert.Empty(numbers.Badges);

            // The tool off, SHIFT held again: numbered.
            tab.NoteKeyModifiers(KeyModifiers.None);
            tab.SetDrawingTraces(false);
            tab.NoteKeyModifiers(KeyModifiers.Shift);
            Assert.Equal(14, numbers.Badges.Count);

            window.Close();
        });
    }

    // ###########################################################################################
    // The viewer reset as another board is shown lets go of the pins numbered and the legend (code review, 2026-11-18): they
    // stood over the next board's thumbnail until its picture was read.
    // ###########################################################################################
    [Fact]
    public void Another_board_shown_lets_go_of_the_pins_numbered_and_the_legend()
    {
        UiTest.Run(() =>
        {
            (Window window, TabSchematics tab) = OnTheBoardsTop();
            PinNumbersOverlay numbers = tab.GetControl<PinNumbersOverlay>("SchematicsPinNumbersOverlay");
            Border legend = tab.GetControl<Border>("PinDirectionLegend");

            tab.UpdateHighlightsForComponents(["U5"]);
            Dispatcher.UIThread.RunJobs();
            tab.NoteKeyModifiers(KeyModifiers.Shift);
            Assert.NotEmpty(numbers.Badges);
            Assert.True(legend.IsVisible);

            tab.ResetSchematicsViewer();

            Assert.Empty(numbers.Badges);
            Assert.Empty(numbers.FirstPins);
            Assert.False(legend.IsVisible);

            window.Close();
        });
    }

    // ###########################################################################################
    // SHIFT REBUILDS THE KiCad OVERLAY ONLY WHERE IT IS WHAT HIGHLIGHTS THE TRACES (code review, 2026-11-18): it reaches the
    // tab from anywhere in CRT's window - a capital letter typed in the component filter, a range picked in a list - and in
    // any other mode the overlay draws the same with it or without, yet each press and release rebuilt it.
    // ###########################################################################################
    [Theory]
    [InlineData("Always", false)]
    [InlineData("Disabled", false)]
    [InlineData("HoldShift", true)]
    public void Shift_rebuilds_the_kicad_overlay_only_where_it_highlights_the_traces(string mode, bool rebuilds)
    {
        UiTest.Run(() =>
        {
            UserSettings.InteractiveCadTraceHoverMode = mode;
            (Window window, TabSchematics tab) = OnTheBoardsTop();
            Dispatcher.UIThread.RunJobs();

            int before = tab.KiCadOverlayRefreshRequestsForTests;
            tab.NoteKeyModifiers(KeyModifiers.Shift);
            tab.NoteKeyModifiers(KeyModifiers.None);
            int after = tab.KiCadOverlayRefreshRequestsForTests;

            if (rebuilds)
            {
                Assert.True(after > before, $"{after - before} rebuilds asked for");
            }
            else
            {
                Assert.Equal(before, after);
            }

            window.Close();
        });
    }

    // ###########################################################################################
    // THE LEGEND IS A PANEL RIGHT OF "Important KiCad signals", ONE DIRECTION A LINE (owner request, 2026-11-17: "I would
    // like the pin direction legend to show vertically and right to the "Important KiCad signals" panel"), taking no click
    // from the schematic image under it.
    // ###########################################################################################
    [Fact]
    public void The_legend_stands_right_of_the_important_signals_panel_one_direction_a_line()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            Border legend = tab.GetControl<Border>("PinDirectionLegend");
            Border signals = tab.GetControl<Border>("ImportantSignalsPanel");

            Assert.Same(signals.Parent, legend.Parent);
            Assert.Equal(Grid.GetRow(signals), Grid.GetRow(legend));
            Assert.Equal(Grid.GetColumn(signals) + 1, Grid.GetColumn(legend));
            Assert.Equal(Avalonia.Layout.Orientation.Vertical, tab.GetControl<StackPanel>("PinDirectionLegendItems").Orientation);
            Assert.False(legend.IsHitTestVisible);
        });
    }

    // ###########################################################################################
    // *** THE BADGES IN THE TOP LEFT CORNER STAND OVER THE PINS' NUMBERS (owner report, 2026-11-17: hovering a component in the
    // top-left corner while holding SHIFT, "the "Khaki" badges are shown below the pin numbers"). *** And over the components'
    // labels - both drawn on the schematic image, which the badges name.
    // ###########################################################################################
    [Fact]
    public void The_badges_in_the_top_left_corner_stand_over_the_pin_numbers_and_the_labels()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            PinNumbersOverlay numbers = tab.GetControl<PinNumbersOverlay>("SchematicsPinNumbersOverlay");
            Canvas labels = tab.GetControl<Canvas>("SchematicsLabelsCanvas");

            foreach (string row in new[] { "SchematicsInfoBadgesRow", "SchematicsHoverBadgesRow" })
            {
                StackPanel badges = tab.GetControl<StackPanel>(row);

                // ZIndex orders the children of one panel only.
                Assert.Same(numbers.Parent, badges.Parent);
                Assert.Same(labels.Parent, badges.Parent);
                Assert.True(badges.ZIndex > numbers.ZIndex, $"{row} at {badges.ZIndex}, the pins' numbers at {numbers.ZIndex}");
                Assert.True(badges.ZIndex > labels.ZIndex, $"{row} at {badges.ZIndex}, the labels at {labels.ZIndex}");
            }
        });
    }

    // ###########################################################################################
    // *** PIN 1 OF A COMPONENT SELECTED IS ALWAYS MARKED IN ORANGE (owner decision, 2026-11-17: "remove the "Mark first pin on
    // component" checkbox, but keep its functionality enabled per default, so everyone gets this"). *** With "Hold SHIFT to
    // number pins on hover and select" ticked - as it is unless unticked - it was not marked at all until then.
    // ###########################################################################################
    [Fact]
    public void Pin_1_of_a_selected_component_is_marked_whatever_the_numbers_do()
    {
        UiTest.Run(() =>
        {
            (Window window, TabSchematics tab) = OnTheBoardsTop();
            KiCadOverlayRenderControl overlay = tab.GetControl<KiCadOverlayRenderControl>("SchematicsKiCadOverlayCanvas");

            bool PinOneMarked() => overlay.Primitives.Any(primitive =>
                primitive.Fill is Avalonia.Media.ISolidColorBrush { Color: var color } && color == Avalonia.Media.Colors.Orange);

            Assert.True(UserSettings.SchematicsNumberPinsHoldingShift);

            // *** AND FRAMED IN BLACK, A LINE JUST OUTSIDE THE ORANGE (owner requests, 2026-11-17: "the "normal" first pin
            // marking having this 1px 2xblack border so it is more visible", then "just the one outer black border is fine") -
            // drawn with the pins' numbers, in the panel's own pixels; a numbered pin 1 has its badge's ring instead. ***
            PinNumbersOverlay numbers = tab.GetControl<PinNumbersOverlay>("SchematicsPinNumbersOverlay");

            tab.UpdateHighlightsForComponents(["U5"]);
            Dispatcher.UIThread.RunJobs();
            Assert.True(PinOneMarked());
            FirstPinMark frame = Assert.Single(numbers.FirstPins);

            // Holding SHIFT, as without - the pins numbered, pin 1's badge ringed and on its frame's place.
            tab.NoteKeyModifiers(KeyModifiers.Shift);
            Dispatcher.UIThread.RunJobs();
            Assert.True(PinOneMarked());
            Assert.Empty(numbers.FirstPins);
            Assert.Equal(frame.Center, numbers.Badges.Single(badge => badge.IsFirst).Center);

            // The numbers unticked: SHIFT numbers nothing, and pin 1 is framed.
            tab.NoteKeyModifiers(KeyModifiers.None);
            tab.GetControl<CheckBox>("CheckGlobalNumberPins").IsChecked = false;
            tab.NoteKeyModifiers(KeyModifiers.Shift);
            Dispatcher.UIThread.RunJobs();
            Assert.True(PinOneMarked());
            Assert.Single(numbers.FirstPins);
            Assert.Empty(numbers.Badges);

            // Nothing selected: no pin 1 to mark.
            tab.UpdateHighlightsForComponents([]);
            Dispatcher.UIThread.RunJobs();
            Assert.False(PinOneMarked());
            Assert.Empty(numbers.FirstPins);

            window.Close();
        });
    }

    // ###########################################################################################
    // "Hold SHIFT to number pins on hover and select" is what "Show color for pin direction" waits for (owner request,
    // 2026-11-17: "make sure to enable the below "Show pin direction" checkbox, and likewise disable it, if this checkbox is
    // not checked"); neither is there on a schematic image with no KiCad pads.
    // ###########################################################################################
    [Fact]
    public void Numbering_the_pins_lets_their_direction_be_shown()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            tab.UpdatePinsOnImageRowsForTests(hasPadData: true);

            Assert.True(tab.GetControl<CheckBox>("CheckGlobalNumberPins").IsChecked);
            Assert.True(tab.GetControl<CheckBox>("CheckGlobalShowPinDirection").IsEnabled);

            tab.GetControl<CheckBox>("CheckGlobalNumberPins").IsChecked = false;
            tab.UpdatePinsOnImageRowsForTests(hasPadData: true);

            Assert.False(UserSettings.SchematicsNumberPinsHoldingShift);
            Assert.False(tab.GetControl<CheckBox>("CheckGlobalShowPinDirection").IsEnabled);
            Assert.True(tab.GetControl<StackPanel>("GlobalShowPinDirectionRow").IsVisible);

            // No KiCad pads: neither row.
            tab.UpdatePinsOnImageRowsForTests(hasPadData: false);
            Assert.False(tab.GetControl<StackPanel>("GlobalNumberPinsRow").IsVisible);
            Assert.False(tab.GetControl<StackPanel>("GlobalShowPinDirectionRow").IsVisible);
        });
    }

    private static BoardData Board() => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U5", Category = "IC", TechnicalNameOrValue = "74LS08", Package = "DIP-14" }],
        ComponentPins = [new ComponentPinEntry { BoardLabel = "U5", Pin = "1", Name = "1A", Direction = "Input" }],
    };

    // ###########################################################################################
    // THE PIN UNDER THE POINTER IS THREE BADGES (owner request, 2026-11-17: "split this in to 3 different badges
    // (horizontally aligned)": "Pin 20", "Net ~{CHAROM}", "Direction Input" - "Pin 20 | ~{CHAROM}" in one until then): its
    // direction the one its pinout gives it - the list's, else the datasheet's - and a badge with nothing to say not there.
    // ###########################################################################################
    [Fact]
    public void The_pin_under_the_pointer_is_three_badges_its_number_its_net_and_its_direction()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics { CurrentBoardDataOverrideForTests = Board() };

            tab.ShowHoverPinBadgesForTests("1", "~{CHAROM}", "U5");

            Assert.True(tab.GetControl<StackPanel>("SchematicsHoverPinBadges").IsVisible);
            Assert.Equal("1", tab.GetControl<TextBlock>("SchematicsHoverPadText").Text);
            Assert.True(tab.GetControl<Border>("SchematicsHoverNetBorder").IsVisible);
            Assert.Equal("~{CHAROM}", tab.GetControl<TextBlock>("SchematicsHoverNetText").Text);
            Assert.True(tab.GetControl<Border>("SchematicsHoverDirectionBorder").IsVisible);
            Assert.Equal("Input", tab.GetControl<TextBlock>("SchematicsHoverDirectionText").Text);

            // A pin its list says nothing of: the datasheet's direction; a pad on no net: no net badge.
            tab.ShowHoverPinBadgesForTests("3", string.Empty, "U5");
            Assert.Equal("Output", tab.GetControl<TextBlock>("SchematicsHoverDirectionText").Text);
            Assert.False(tab.GetControl<Border>("SchematicsHoverNetBorder").IsVisible);

            // A component with no direction for its pin: no direction badge.
            tab.ShowHoverPinBadgesForTests("2", "GND", "J1");
            Assert.False(tab.GetControl<Border>("SchematicsHoverDirectionBorder").IsVisible);
            Assert.True(tab.GetControl<Border>("SchematicsHoverPadBorder").IsVisible);

            // No pad: no badges.
            tab.ShowHoverPinBadgesForTests(string.Empty, "GND", "U5");
            Assert.False(tab.GetControl<StackPanel>("SchematicsHoverPinBadges").IsVisible);
        });
    }
}
