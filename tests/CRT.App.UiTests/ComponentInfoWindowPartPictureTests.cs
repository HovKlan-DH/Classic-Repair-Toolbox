using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The component popup's picture of a leaded part (owner request, 2026-11-09: "could you also draw an axial and radial
// electrolytic capacitor in same style ... and a disc capacitor as a pinout image and mark (if applicable) its plus and
// minus side? And how about transistors? And diodes?"). PartPicture decides, PartPictureDrawing paints, in the pinout's
// place: the case drawn as a box gives way to it, a part with no pins listed still gets it, and a part CRT cannot draw
// takes no room. Its leads are clicked to go to their images, as the chip's pins are.
//
// COLLECTION NOTE: "HeadlessUi" because it constructs a Window, whose constructor reads UserSettings - pointed at a temp
// file first, as ComponentInfoWindowTests does.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentInfoWindowPartPictureTests : IDisposable
{
    private const string BoardLabel = "C61";

    private static readonly string DataRoot = Path.Combine(Path.GetTempPath(), "crt-no-such-data-root-" + Guid.NewGuid().ToString("N"));

    private readonly TempWorkspace thisWorkspace = new();

    public ComponentInfoWindowPartPictureTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
        this.thisWorkspace.Dispose();
    }

    // A radial electrolytic: drawn with its minus stripe on pin 2 by KiCad's convention, beside the pins listed.
    [Fact]
    public void A_radial_electrolytic_is_drawn_in_the_pinout_with_its_leads_named()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Capacitor", Package = "Radial", FriendlyName = "Electrolytic Radial", TechnicalNameOrValue = "10uF 16V" }, Pins("", ""));

            PartPictureDrawing part = window.GetControl<PartPictureDrawing>("InfoPartDrawing");
            Assert.True(window.GetControl<Control>("PinoutSection").IsVisible);
            Assert.True(part.IsVisible);
            Assert.Equal(PartPicture.Shape.RadialElectrolytic, part.PictureForTests!.Shape);
            Assert.Equal(["+", "-"], part.PictureForTests.Leads.Select(lead => lead.Name));
            Assert.False(window.GetControl<PinoutDrawing>("InfoPinoutDrawing").IsVisible);
            Assert.True(window.GetControl<Control>("InfoPinsPanel").IsVisible);
        });
    }

    // A TO-92 drawn as it looks takes the place of its case drawn as a box of three pins.
    [Fact]
    public void A_transistor_drawn_as_it_looks_takes_the_place_of_its_case_drawn_as_a_box()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Transistor", Package = "TO-92", FriendlyName = "NPN", TechnicalNameOrValue = "2SC1815" }, Pins("E", "C", "B"));

            Assert.True(window.GetControl<PartPictureDrawing>("InfoPartDrawing").IsVisible);
            Assert.False(window.GetControl<PinoutDrawing>("InfoPinoutDrawing").IsVisible);
        });
    }

    // With no pins listed the pinout section still shows the part - its leads numbered as the case numbers them - and no
    // empty list of pins beside it.
    [Fact]
    public void A_part_with_no_pins_listed_is_drawn_with_no_empty_pin_list()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Diode", Package = "DO-35", TechnicalNameOrValue = "1N914" }, []);

            PartPictureDrawing part = window.GetControl<PartPictureDrawing>("InfoPartDrawing");
            Assert.True(window.GetControl<Control>("PinoutSection").IsVisible);
            Assert.True(part.IsVisible);
            Assert.Equal(0, part.PictureForTests!.MarkedLead);
            Assert.False(window.GetControl<Control>("InfoPinsPanel").IsVisible);
        });
    }

    // A chip keeps its case drawn from above; a capacitor of a shape not known, and an IC with no pins, show no pinout.
    [Fact]
    public void A_chip_keeps_its_case_and_a_part_CRT_cannot_draw_takes_no_room()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" }, Pins("VCC", "", "", "", "", "", "", "GND"));
            Assert.True(window.GetControl<PinoutDrawing>("InfoPinoutDrawing").IsVisible);
            Assert.False(window.GetControl<PartPictureDrawing>("InfoPartDrawing").IsVisible);

            // "Ceramic" with no package: a disc, a dipped or an axial part - none is guessed.
            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Capacitor", FriendlyName = "Ceramic", TechnicalNameOrValue = "100nF" }, []);
            Assert.False(window.GetControl<PartPictureDrawing>("InfoPartDrawing").IsVisible);
            Assert.False(window.GetControl<Control>("PinoutSection").IsVisible);

            // A single resistor is its color bands, in the same place (owner request, 2026-11-09).
            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", Package = "Axial", TechnicalNameOrValue = "560Ω 1/4W 5%" }, Pins("", ""));
            Assert.False(window.GetControl<PartPictureDrawing>("InfoPartDrawing").IsVisible);
            Assert.False(window.GetControl<PinoutDrawing>("InfoPinoutDrawing").IsVisible);
            Assert.True(window.GetControl<ResistorPicture>("InfoResistorPicture").IsVisible);

            // The next part drawn lets go of it.
            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Transistor", Package = "TO-92", FriendlyName = "NPN", TechnicalNameOrValue = "2SC1815" }, Pins("E", "C", "B"));
            Assert.False(window.GetControl<ResistorPicture>("InfoResistorPicture").IsVisible);
            Assert.True(window.GetControl<PartPictureDrawing>("InfoPartDrawing").IsVisible);
        });
    }

    // A lead clicked goes to its pin's first image, as a chip's pin clicked does.
    [Fact]
    public void A_lead_clicked_goes_to_its_pins_first_image()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow { Width = 900, Height = 800 };
            window.Show();

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Transistor", Package = "TO-92", FriendlyName = "NPN", TechnicalNameOrValue = "2SC1815" }, Pins("E", "C", "B"));

            ListBox thumbnails = window.GetControl<ListBox>("ThumbnailList");
            thumbnails.ItemsSource = new List<ComponentImageItem> { new() { Pin = "1" }, new() { Pin = "3" }, new() { Pin = "2" } };

            for (int i = 0; i < 3; i++)
            {
                Dispatcher.UIThread.RunJobs();
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            }

            PartPictureDrawing part = window.GetControl<PartPictureDrawing>("InfoPartDrawing");
            PartPictureGeometry.LeadPlace collector = part.LayoutForTests!.Leads.Single(lead => lead.Number == 2);
            Point at = part.TranslatePoint(collector.Label.Center, window)!.Value;

            window.MouseDown(at, MouseButton.Left);
            window.MouseUp(at, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(2, thumbnails.SelectedIndex);
            Assert.Equal("2", window.GetControl<TextBlock>("PinFlashText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // A lead pointed at lights its pin's row in the list, as a chip's pin does (owner request, 2026-11-09: "Not only an
    // issue for ICs"); off the leads, no row is lit.
    // ###########################################################################################
    [Fact]
    public void A_lead_pointed_at_lights_its_pins_row_in_the_list()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow { Width = 900, Height = 800 };
            window.Show();

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Transistor", Package = "TO-92", FriendlyName = "NPN", TechnicalNameOrValue = "2SC1815" }, Pins("E", "C", "B"));

            for (int i = 0; i < 3; i++)
            {
                Dispatcher.UIThread.RunJobs();
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            }

            PartPictureDrawing part = window.GetControl<PartPictureDrawing>("InfoPartDrawing");
            PartPictureGeometry.LeadPlace collector = part.LayoutForTests!.Leads.Single(lead => lead.Number == 2);

            window.MouseMove(part.TranslatePoint(collector.Label.Center, window)!.Value);
            Dispatcher.UIThread.RunJobs();

            List<ComponentPopupPin> rows = window.GetControl<ItemsControl>("InfoPinsItemsControl").ItemsSource!.Cast<ComponentPopupPin>().ToList();
            Assert.Equal(["2"], rows.Where(pin => pin.IsPointedAt).Select(pin => pin.Pin));

            window.MouseMove(new Point(2, 2));
            Dispatcher.UIThread.RunJobs();

            rows = window.GetControl<ItemsControl>("InfoPinsItemsControl").ItemsSource!.Cast<ComponentPopupPin>().ToList();
            Assert.DoesNotContain(rows, pin => pin.IsPointedAt);

            window.Close();
        });
    }

    // ###########################################################################################
    // A transistor's picture says what its letters stand for (owner request, 2026-11-09: "E = Emitter, C = Collector, B =
    // Base"), the picture as tall as that takes; a RESISTOR'S PIN LIST STANDS FIRST, its picture under it (owner request,
    // 2026-11-09: "I want the pinout table to be shown first and then the image and explanation afterwards"), wide window
    // or not - another part's list beside its picture where both fit.
    // ###########################################################################################
    [Fact]
    public void A_transistor_explains_its_letters_and_a_resistors_pin_list_stands_over_its_picture()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow { Width = 1400, Height = 900 };
            window.Show();

            void Lay()
            {
                for (int i = 0; i < 3; i++)
                {
                    Dispatcher.UIThread.RunJobs();
                    AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                }
            }

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Transistor", Package = "TO-92", FriendlyName = "NPN", TechnicalNameOrValue = "2SC1815" }, Pins("E", "C", "B"));
            Lay();

            PartPictureDrawing part = window.GetControl<PartPictureDrawing>("InfoPartDrawing");
            Assert.Equal(["E = Emitter", "C = Collector", "B = Base"], part.LegendShownForTests);
            Assert.True(part.NaturalSize.Height > part.LayoutForTests!.Size.Height + (3 * 14), "the legend's lines under the picture");

            Control pins = window.GetControl<Control>("InfoPinsPanel");
            Assert.Equal((0, 1), (Grid.GetRow(pins), Grid.GetColumn(pins)));

            Show(window, new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", Package = "Axial", TechnicalNameOrValue = "4K7Ω 5%" }, Pins("", ""));
            Lay();

            ResistorPicture resistor = window.GetControl<ResistorPicture>("InfoResistorPicture");
            Assert.True(resistor.IsVisible);
            Assert.Equal((0, 0, 2), (Grid.GetRow(pins), Grid.GetColumn(pins), Grid.GetColumnSpan(pins)));
            Assert.Equal(1, Grid.GetRow(resistor));
            Assert.True(resistor.TranslatePoint(new Point(0, 0), window)!.Value.Y > pins.TranslatePoint(new Point(0, pins.Bounds.Height), window)!.Value.Y - 1, "under the list");

            window.Close();
        });
    }

    private static void Show(ComponentInfoWindow window, ComponentEntry entry, IReadOnlyList<ComponentPinEntry> pins) =>
        window.SetComponent(BoardLabel, BoardLabel, [entry], [], [], [], "PAL", DataRoot, regions: [], boardIdentity: string.Empty, pins: pins);

    private static List<ComponentPinEntry> Pins(params string[] names) =>
        names.Select((name, index) => new ComponentPinEntry { BoardLabel = BoardLabel, Pin = (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), Name = name }).ToList();
}
