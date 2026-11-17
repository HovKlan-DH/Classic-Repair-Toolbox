using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE POPUP'S PINS GO TO THEIR IMAGES (owner request, 2026-11-01: "have the pins clickable with mouse
// (and highlighted when hovering) so the user quickly can go to that pin and see what images are
// available for that pin"). With real pointer events through a shown window: only what a control draws
// takes the pointer, which is the trap here.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class PinoutDrawingTests
{
    private static readonly ComponentPinEntry[] Pins =
    [
        new() { Pin = "1", Name = "VCC" },
        new() { Pin = "2", Name = "GND" },
    ];

    private static (PinoutDrawing Drawing, Window Window) Shown(bool canBePicked)
    {
        var drawing = new PinoutDrawing { PinsCanBePicked = canBePicked, HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left, VerticalAlignment = Avalonia.Layout.VerticalAlignment.Top };
        var window = new Window { Content = drawing, Width = 500, Height = 300 };

        window.Show();
        drawing.Show("DIP-8", Pins);
        Dispatcher.UIThread.RunJobs();

        return (drawing, window);
    }

    // Far out on a pin's row, over no name and no lead - where nothing but the clear fill is drawn.
    private static Point FarOutOn(PinoutLayout layout, PinoutPin pin) =>
        pin.Side == PinSide.Left ? new Point(2, pin.Lead.Center.Y) : new Point(layout.Size.Width - 2, pin.Lead.Center.Y);

    [Fact]
    public void A_pin_lights_up_under_the_pointer_and_goes_out_when_it_leaves()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: true);
            PinoutLayout layout = drawing.LayoutForTests!;

            window.MouseMove(FarOutOn(layout, layout.Pins[0]));
            Assert.Equal(1, drawing.LitPinForTests);
            Assert.NotNull(drawing.Cursor);

            window.MouseMove(FarOutOn(layout, layout.Pins[7]));
            Assert.Equal(8, drawing.LitPinForTests);

            window.MouseMove(layout.Body.Center);
            Assert.Null(drawing.LitPinForTests);
            Assert.Null(drawing.Cursor);

            window.Close();
        });
    }

    [Fact]
    public void A_click_on_a_pin_says_which_and_one_beside_the_pins_says_nothing()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: true);
            PinoutLayout layout = drawing.LayoutForTests!;
            var picked = new List<int>();
            drawing.PinPicked += (_, e) => picked.Add(e.Pin);

            window.MouseDown(FarOutOn(layout, layout.Pins[6]), MouseButton.Left);
            window.MouseUp(FarOutOn(layout, layout.Pins[6]), MouseButton.Left);
            Assert.Equal([7], picked);

            window.MouseDown(layout.Body.Center, MouseButton.Left);
            window.MouseUp(layout.Body.Center, MouseButton.Left);
            Assert.Equal([7], picked);

            // Another button picks nothing.
            window.MouseDown(FarOutOn(layout, layout.Pins[0]), MouseButton.Right);
            window.MouseUp(FarOutOn(layout, layout.Pins[0]), MouseButton.Right);
            Assert.Equal([7], picked);

            window.Close();
        });
    }

    // The component pane's drawing only shows: nothing lights up, nothing is picked and no pin is said to be pointed at.
    [Fact]
    public void A_drawing_whose_pins_cannot_be_picked_ignores_the_pointer()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: false);
            PinoutLayout layout = drawing.LayoutForTests!;
            var picked = new List<int>();
            var hovered = new List<int?>();
            drawing.PinPicked += (_, e) => picked.Add(e.Pin);
            drawing.PinHovered += (_, e) => hovered.Add(e.Pin);

            window.MouseMove(layout.Pins[0].Lead.Center);
            Assert.Null(drawing.LitPinForTests);

            window.MouseDown(layout.Pins[0].Lead.Center, MouseButton.Left);
            window.MouseUp(layout.Pins[0].Lead.Center, MouseButton.Left);
            Assert.Empty(picked);
            Assert.Empty(hovered);

            window.Close();
        });
    }

    // ###########################################################################################
    // The pointer moving onto a pin, to another and off them says so, once each (owner request, 2026-11-09: the list's row
    // lit too) - and the host lighting one (HoverPin, the list's row under the pointer) says nothing, or the two would
    // answer each other.
    // ###########################################################################################
    [Fact]
    public void The_pointer_moving_onto_and_off_a_pin_says_which_and_the_host_lighting_one_says_nothing()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: true);
            PinoutLayout layout = drawing.LayoutForTests!;
            var hovered = new List<int?>();
            drawing.PinHovered += (_, e) => hovered.Add(e.Pin);

            window.MouseMove(FarOutOn(layout, layout.Pins[0]));
            window.MouseMove(FarOutOn(layout, layout.Pins[0]) + new Point(1, 0));
            window.MouseMove(FarOutOn(layout, layout.Pins[7]));
            window.MouseMove(layout.Body.Center);
            Assert.Equal([1, 8, null], hovered);

            drawing.HoverPin(5);
            drawing.HoverPin(null);
            Assert.Equal([1, 8, null], hovered);

            window.Close();
        });
    }

    // The drawing is as wide as its names need, so the popup's list has the room beside it.
    [Fact]
    public void The_drawing_is_only_as_wide_as_its_names_need()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: false);

            Assert.True(drawing.LayoutForTests!.Size.Width < LogicChipGeometry.EmptyCaseFor("DIP-8")!.Case.Size.Width);

            drawing.Show("DIP-8", [new ComponentPinEntry { Pin = "1", Name = new string('x', 200) }]);

            Assert.Equal(LogicChipGeometry.EmptyCaseFor("DIP-8")!.Case.Size.Width, drawing.LayoutForTests!.Size.Width);

            window.Close();
        });
    }

    // The host lights one up - the list beside it does - and another board's pins start with none.
    [Fact]
    public void The_host_can_light_a_pin_and_showing_another_case_lets_go_of_it()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: true);

            drawing.MarkPin(3);
            Assert.Equal(3, drawing.LitPinForTests);

            drawing.MarkPin(null);
            Assert.Null(drawing.LitPinForTests);

            drawing.MarkPin(3);
            drawing.Show("DIP-14", Pins);
            Assert.Null(drawing.LitPinForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // A CHIP CRT KNOWS IS DRAWN WITH ITS INSIDES (owner request, 2026-11-16): its gates in a case of its own, each pin in
    // the color of its direction (owner request, 2026-11-17) - the board's, else the datasheet's, the ground and the supply
    // grey - the legend naming only the directions drawn, and every pin named outside: the board's name, else the
    // datasheet's, as nothing is written at the pins inside. The pointer finds its pins as on any case, on the taller pitch.
    // Shown again as a chip CRT knows nothing inside of, it is the same case empty (owner request, 2026-11-17: "unify ALL
    // pinout IC's so they look the same"); only a single row is the plain case.
    // ###########################################################################################
    [Fact]
    public void A_known_chip_is_drawn_with_its_insides_and_its_pins_are_found_by_the_pointer()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: true);
            window.Width = 900;
            window.Height = 600;

            drawing.Show(
                "DIP-14",
                [new ComponentPinEntry { Pin = "3", Name = "BA" }, new ComponentPinEntry { Pin = "1", Name = "1A", Direction = "Input/output" }],
                LogicChips.Named("74LS08"));
            Dispatcher.UIThread.RunJobs();

            LogicInside inside = Assert.IsType<LogicInside>(drawing.InsideForTests);
            PinoutLayout layout = drawing.LayoutForTests!;
            Assert.Same(inside.Case, layout);
            Assert.Equal(4, inside.Symbols.OfType<LogicGateShape>().Count());
            Assert.Equal(LogicChipGeometry.Pitch, layout.Pitch);

            Assert.Equal(PinDirectionKind.InputOutput, drawing.DirectionsForTests[1]);
            Assert.Equal(PinDirectionKind.Input, drawing.DirectionsForTests[2]);
            Assert.Equal(PinDirectionKind.Output, drawing.DirectionsForTests[3]);
            Assert.Equal(PinDirectionKind.Other, drawing.DirectionsForTests[7]);
            Assert.Equal(PinDirectionKind.Other, drawing.DirectionsForTests[14]);
            Assert.Equal([PinDirectionKind.Input, PinDirectionKind.Output, PinDirectionKind.InputOutput, PinDirectionKind.Other], drawing.LegendDirectionsForTests);

            Assert.Equal("BA", drawing.NamesForTests[3]);
            Assert.Equal("1A", drawing.NamesForTests[1]);
            Assert.Equal("GND", drawing.NamesForTests[7]);
            Assert.Equal("VCC", drawing.NamesForTests[14]);
            Assert.Equal(14, drawing.NamesForTests.Count);

            // Under the case, what the chip is and the legend - so the drawing is taller than the case.
            Assert.True(drawing.NaturalSize.Height > layout.Size.Height);
            Assert.Equal(drawing.NaturalSize, drawing.DesiredSize);

            window.MouseMove(FarOutOn(layout, layout.Pins[2]));
            Assert.Equal(3, drawing.LitPinForTests);

            window.MouseMove(FarOutOn(layout, layout.Pins[10]));
            Assert.Equal(11, drawing.LitPinForTests);

            window.MouseMove(layout.Body.Center);
            Assert.Null(drawing.LitPinForTests);

            // Shown as a chip CRT knows nothing inside of: the same case, empty, its pins grey but what its list says.
            drawing.Show("DIP-14", [new ComponentPinEntry { Pin = "2", Name = "D0", Direction = "Input/output" }]);
            LogicInside empty = Assert.IsType<LogicInside>(drawing.InsideForTests);
            Assert.Empty(empty.Symbols);
            Assert.Equal(LogicChipGeometry.Pitch, drawing.LayoutForTests!.Pitch);
            Assert.Equal(PinDirectionKind.InputOutput, drawing.DirectionsForTests[2]);
            Assert.Equal(13, drawing.DirectionsForTests.Values.Count(kind => kind == PinDirectionKind.Other));
            Assert.Equal([PinDirectionKind.InputOutput, PinDirectionKind.Other], drawing.LegendDirectionsForTests);

            // No pin with a direction a color says - a DIP-40 just given its package: no legend, nothing under the case (code
            // review, 2026-11-17: "Pin direction: Other" explained nothing).
            drawing.Show("DIP-40", [new ComponentPinEntry { Pin = "20", Name = "GND", Direction = "Ground" }]);
            Assert.IsType<LogicInside>(drawing.InsideForTests);
            Assert.Empty(drawing.LegendDirectionsForTests);
            Assert.Equal(drawing.LayoutForTests!.Size, drawing.NaturalSize);

            // A pin written with more after its number - "2 [64K]" - is named as it is colored, by its number (code review,
            // 2026-11-17: it was colored as pin 2 and its name left out).
            drawing.Show("DIP-14", [new ComponentPinEntry { Pin = "2 [64K]", Name = "A1", Direction = "Output" }], LogicChips.Named("74LS08"));
            Assert.Equal("A1", drawing.NamesForTests[2]);
            Assert.Equal(PinDirectionKind.Output, drawing.DirectionsForTests[2]);

            // A single row keeps the plain case.
            drawing.Show("TO-92", []);
            Assert.Null(drawing.InsideForTests);
            Assert.Equal(PinoutGeometry.Pitch, drawing.LayoutForTests!.Pitch);

            window.Close();
        });
    }

    // ###########################################################################################
    // In a place narrower than it, the chip is drawn smaller, never cut (as a part's picture is) - and the pointer still
    // finds each pin where it is drawn.
    // ###########################################################################################
    [Fact]
    public void A_chip_drawn_smaller_in_a_narrow_place_still_finds_its_pins()
    {
        UiTest.Run(() =>
        {
            var drawing = new PinoutDrawing { PinsCanBePicked = true, HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left, VerticalAlignment = Avalonia.Layout.VerticalAlignment.Top };
            var window = new Window { Content = drawing, Width = 200, Height = 600 };

            window.Show();
            drawing.Show("DIP-16", [], LogicChips.Named("74LS257"));
            Dispatcher.UIThread.RunJobs();

            PinoutLayout layout = drawing.LayoutForTests!;
            double scale = drawing.Bounds.Width / drawing.NaturalSize.Width;

            Assert.True(drawing.NaturalSize.Width > 200);
            Assert.True(drawing.Bounds.Width <= 200);

            // As tall for its width as it is drawn whole - to the pixel layout rounds to.
            Assert.InRange(drawing.Bounds.Height - (drawing.NaturalSize.Height * scale), -1, 1);

            window.MouseMove(new Point(layout.Pins[4].Lead.Center.X * scale, layout.Pins[4].Lead.Center.Y * scale));
            Assert.Equal(5, drawing.LitPinForTests);

            window.MouseMove(new Point(layout.Pins[12].Lead.Center.X * scale, layout.Pins[12].Lead.Center.Y * scale));
            Assert.Equal(13, drawing.LitPinForTests);

            window.Close();
        });
    }

    // The pin of the image shown stays lit whatever the pointer does (owner request, 2026-11-02); the
    // pin under the pointer in the list comes and goes beside it.
    [Fact]
    public void The_pin_marked_stays_while_another_is_pointed_at()
    {
        UiTest.Run(() =>
        {
            (PinoutDrawing drawing, Window window) = Shown(canBePicked: true);

            drawing.MarkPin(3);
            drawing.HoverPin(5);

            Assert.Equal(5, drawing.LitPinForTests);
            Assert.Equal(3, drawing.MarkedPinForTests);

            drawing.HoverPin(null);

            Assert.Equal(3, drawing.LitPinForTests);
            Assert.Equal(3, drawing.MarkedPinForTests);

            window.Close();
        });
    }
}
