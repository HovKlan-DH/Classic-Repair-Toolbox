using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Contribute tab's "Pinout" page draws a logic chip CRT knows with its insides (owner request, 2026-11-16), from what
// the pane's boxes say - so a technical name typed there is drawn at once - each pin in the color of its direction as the
// page's pins say it, else as the datasheet does (owner request, 2026-11-17), and again when a pin's direction changes.
// Several components picked at once get none, as their fields may differ under one box.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentPaneLogicChipTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData Published() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U27", FriendlyName = "AND", TechnicalNameOrValue = "74LS08", Category = "IC", Package = "DIP-14" },
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510", Category = "IC", Package = "DIP-40" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U27", Pin = "7", Name = "GND", Direction = "Power" },
            new ComponentPinEntry { BoardLabel = "U27", Pin = "3", Name = "BA", Direction = "Output" },
        ],
    };

    private DraftSession OpenDraft()
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, Published());

        return DraftSession.Open(this.DraftsRoot, BoardKey)!;
    }

    private static (ComponentPane Pane, Window Window) Shown(DraftSession session, string label)
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 1400 };
        window.Show();

        pane.Show(session, label, "PAL", Published());
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    [Fact]
    public void A_known_chip_is_drawn_with_its_insides_and_follows_its_pins_directions()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U27");
            pane.ShowPage(ComponentPage.Pinout);

            PinoutDrawing drawing = pane.GetControl<PinoutDrawing>("PinoutPicture");
            Assert.True(drawing.IsVisible);
            Assert.NotNull(drawing.InsideForTests);
            Assert.Equal(PinDirectionKind.Output, drawing.DirectionsForTests[3]);
            Assert.Equal(PinDirectionKind.Other, drawing.DirectionsForTests[7]);
            Assert.Equal(PinDirectionKind.Input, drawing.DirectionsForTests[1]);

            // A pin's direction changed on the page: its color follows.
            pane.PinsForTests.Single(pin => pin.Pin == "3").Direction = "Input/output";
            pane.CommitHeld();
            Assert.Equal(PinDirectionKind.InputOutput, drawing.DirectionsForTests[3]);

            window.Close();
        });
    }

    // A technical name typed is drawn at once: a chip CRT knows with its insides, another in the same case empty.
    [Fact]
    public void A_technical_name_typed_is_drawn_with_or_without_insides_at_once()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U8");
            pane.ShowPage(ComponentPage.Pinout);

            PinoutDrawing drawing = pane.GetControl<PinoutDrawing>("PinoutPicture");
            Assert.Empty(Assert.IsType<LogicInside>(drawing.InsideForTests).Symbols);

            pane.GetControl<TextBox>("TechnicalNameBox").Text = "74LS32";
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-14";
            pane.CommitHeld();

            Assert.NotNull(drawing.InsideForTests);
            Assert.Equal(LogicFunction.Or, drawing.InsideForTests!.Symbols.OfType<LogicGateShape>().First().Unit.Function);
            Assert.Equal(4, drawing.InsideForTests!.Symbols.OfType<LogicGateShape>().Count());

            pane.GetControl<TextBox>("TechnicalNameBox").Text = "6510";
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "DIP-40";
            pane.CommitHeld();

            Assert.Empty(Assert.IsType<LogicInside>(drawing.InsideForTests).Symbols);
            Assert.Equal(40, drawing.LayoutForTests!.Pins.Count);

            window.Close();
        });
    }
}
