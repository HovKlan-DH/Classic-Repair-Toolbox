using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Contribute tab's "Pinout" page draws a leaded part as it looks (owner request, 2026-11-09), from what the pane's
// boxes say - so a change typed there is drawn at once - in the place of the case drawn as a box; a chip keeps its case
// drawn from above, and several components picked at once get no picture, as their fields may differ under one box.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentPanePartPictureTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData Published() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "C61", FriendlyName = "Electrolytic Radial", TechnicalNameOrValue = "10uF 16V", Category = "Capacitor", Package = "Radial" },
            new ComponentEntry { BoardLabel = "C62", FriendlyName = "Electrolytic Radial", TechnicalNameOrValue = "10uF 16V", Category = "Capacitor", Package = "Radial" },
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510", Category = "IC", Package = "DIP-8" },
            new ComponentEntry { BoardLabel = "R5", FriendlyName = "Pull-up", TechnicalNameOrValue = "4K7Ω 5%", Category = "Resistor", Package = "Axial" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "C61", Pin = "1" },
            new ComponentPinEntry { BoardLabel = "C61", Pin = "2" },
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" },
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

    // ###########################################################################################
    // A SINGLE RESISTOR IS ITS COLOR BANDS ON THE "Pinout" PAGE TOO (owner request, 2026-11-09: "the image/drawing of the
    // resistor should be in the pinout instead (like the other components you are drawing, including the explanation of
    // colors"), in the case's place, following the value typed; a value the bands cannot say shows none.
    // ###########################################################################################
    [Fact]
    public void A_resistor_is_its_color_bands_on_the_pinout_page_and_follows_its_value()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "R5");
            pane.ShowPage(ComponentPage.Pinout);

            ResistorPicture resistor = pane.GetControl<ResistorPicture>("PinoutResistorPicture");
            Assert.True(resistor.IsVisible);

            // Under the pins, across the page (owner request, 2026-11-09: "the pinout table to be shown first").
            Assert.Equal((1, 2), (Grid.GetRow(resistor), Grid.GetColumnSpan(resistor)));
            Assert.Equal(["Yellow", "Violet", "Red", "Gold"], resistor.ColorsShownForTests);
            Assert.False(pane.GetControl<PinoutDrawing>("PinoutPicture").IsVisible);
            Assert.False(pane.GetControl<PartPictureDrawing>("PinoutPartDrawing").IsVisible);

            pane.GetControl<TextBox>("TechnicalNameBox").Text = "100KΩ 5%";
            pane.CommitHeld();
            Assert.Equal(["Brown", "Black", "Yellow", "Gold"], resistor.ColorsShownForTests);

            pane.GetControl<TextBox>("TechnicalNameBox").Text = "wire";
            pane.CommitHeld();
            Assert.False(resistor.IsVisible);

            // Another component shown lets go of it.
            pane.GetControl<TextBox>("TechnicalNameBox").Text = "4K7Ω 5%";
            pane.CommitHeld();
            pane.Show(session, "C61", "PAL", Published());
            Dispatcher.UIThread.RunJobs();
            Assert.False(resistor.IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void A_capacitor_is_drawn_as_it_looks_and_follows_what_is_typed()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "C61");

            PartPictureDrawing part = pane.GetControl<PartPictureDrawing>("PinoutPartDrawing");
            Assert.True(part.IsVisible);
            Assert.Equal(PartPicture.Shape.RadialElectrolytic, part.PictureForTests!.Shape);
            Assert.Equal("10µF\n16V", part.PictureForTests.Marking);
            Assert.False(pane.GetControl<PinoutDrawing>("PinoutPicture").IsVisible);

            // Typed as a ceramic one, it is drawn without a polarity at once.
            pane.GetControl<TextBox>("FriendlyNameBox").Text = "Ceramic";
            pane.CommitHeld();

            Assert.Equal(PartPicture.Shape.RadialCapacitor, part.PictureForTests!.Shape);
            Assert.Null(part.PictureForTests.MarkedLead);

            // A case CRT does not draw as a part: nothing.
            pane.GetControl<AutoCompleteBox>("PackageBox").Text = "Other";
            pane.CommitHeld();

            Assert.False(part.IsVisible);

            window.Close();
        });
    }

    // Two electrolytics picked together - their values could differ under one empty box, which would draw a wrong one.
    [Fact]
    public void Several_components_picked_at_once_get_no_picture()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            var pane = new ComponentPane { AutoSave = false };
            var window = new Window { Content = pane, Width = 1100, Height = 1400 };
            window.Show();

            pane.ShowSeveral(session, ["C61", "C62"], "PAL", Published());
            Dispatcher.UIThread.RunJobs();

            Assert.True(pane.IsSeveral);
            Assert.False(pane.GetControl<PartPictureDrawing>("PinoutPartDrawing").IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void A_chip_keeps_its_case_drawn_from_above()
    {
        DraftSession session = this.OpenDraft();

        UiTest.Run(() =>
        {
            (ComponentPane pane, Window window) = Shown(session, "U8");

            Assert.True(pane.GetControl<PinoutDrawing>("PinoutPicture").IsVisible);
            Assert.False(pane.GetControl<PartPictureDrawing>("PinoutPartDrawing").IsVisible);

            window.Close();
        });
    }
}
