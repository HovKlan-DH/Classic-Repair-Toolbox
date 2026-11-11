using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Avalonia;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The component info popup - the window that opens when a component is clicked on a schematic.
//
// It was the third-worst file in the codebase at 0% (804 lines), despite needing no seams at
// all: SetComponent is public and takes plain data (component entries, images, local files,
// links, a region and a data root), so a test can drive it directly. What kept it uncovered was
// simply that nothing opened it.
//
// WHAT THESE COVER. The content decisions SetComponent makes: which component entry wins for the
// current region, how the title and info lines are composed from it, and which sections appear
// or stay hidden based on what the board actually carries. These are the things a user sees
// wrong when board data changes shape.
//
// WHAT THEY DO NOT. Image loading is async and needs real files on disk (LoadImagesAsync), so
// the fixtures below carry image ENTRIES only where a test is about filtering, never about
// decoding. Anything involving the oscilloscope session, the IC test panel or opening an
// external file is out of scope here - the first two need hardware state and the third is
// Process.Start, which rule 6 puts out of bounds.
//
// COLLECTION NOTE: "HeadlessUi" because it constructs a Window. The constructor reads
// UserSettings for its saved size and splitter positions, so every test points UserSettings at
// a temp file first - without that, constructing the window reads the developer's real settings.
// ###########################################################################################
[Collection("HeadlessUi")]
public class ComponentInfoWindowTests : IDisposable
{
    private const string BoardLabel = "U8";
    // A root nothing is ever written under, built for THIS machine - never a literal @"C:\..."
    // (see TestPathLiteralTests).
    private static readonly string DataRoot = Path.Combine(Path.GetTempPath(), "crt-no-such-data-root-" + Guid.NewGuid().ToString("N"));

    private readonly TempWorkspace thisWorkspace = new();

    public ComponentInfoWindowTests()
    {
        // A fresh, empty settings file per test: the window's constructor reads window size,
        // splitter ratio and the two switch states from UserSettings.
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
        this.thisWorkspace.Dispose();
    }

    // -----------------------------------------------------------------------------------------
    // Construction
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_window_constructs_without_throwing()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            Assert.NotNull(window);
        });
    }

    // The popup's three helps that open a Wiki page read "Learn more" (CRT 4.0 step 5, the plan:
    // "The remaining help buttons read 'Learn more'") - a link beside the tool it explains, never
    // a bare "?" that does not say it opens a page.
    [Theory]
    [InlineData("MiniProHelpButton")]
    [InlineData("SyncOscilloscopeHelpButton")]
    [InlineData("NumpadOscilloscopeHelpButton")]
    public void A_help_that_opens_a_Wiki_page_reads_Learn_more(string name)
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();
            var button = window.GetControl<Avalonia.Controls.Button>(name);

            Assert.Contains("LearnMore", button.Classes);
            Assert.Equal("Learn more", ((Avalonia.Controls.TextBlock)button.Content!).Text);
            Assert.Contains("Wiki", Assert.IsType<string>(Avalonia.Controls.ToolTip.GetTip(button)));
        });
    }

    // -----------------------------------------------------------------------------------------
    // The title line
    // -----------------------------------------------------------------------------------------

    // The title joins board label, friendly name and technical name with " | ", skipping any that
    // are blank - so a sparsely filled component still reads correctly rather than showing
    // stranded separators.
    [Fact]
    public void The_title_joins_the_label_friendly_name_and_technical_name()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, entries: new[]
            {
                Component(friendlyName: "SuperPLA", technicalName: "251715"),
            });

            Assert.Equal("U8 | SuperPLA | 251715", TextOf(window, "TitleText"));
        });
    }

    [Fact]
    public void The_title_skips_blank_parts_rather_than_leaving_separators()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, entries: new[]
            {
                Component(friendlyName: "SuperPLA", technicalName: ""),
            });

            Assert.Equal("U8 | SuperPLA", TextOf(window, "TitleText"));
        });
    }

    // With no component entry the board label alone still forms the title - the label is always
    // known, because it is what was clicked on the schematic.
    [Fact]
    public void The_title_is_just_the_board_label_when_there_is_no_matching_entry()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, entries: Array.Empty<ComponentEntry>(), displayText: "U8 (unknown)");

            Assert.Equal("U8", TextOf(window, "TitleText"));
        });
    }

    // The supplied display text is the last resort, used only when there is nothing at all to
    // build a title from - no entry AND no board label. The popup must never open blank-headed.
    [Fact]
    public void The_title_falls_back_to_the_display_text_when_there_is_nothing_else()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            window.SetComponent(
                string.Empty,
                "Unknown component",
                new List<ComponentEntry>(),
                new List<ComponentImageEntry>(),
                new List<ComponentLocalFileEntry>(),
                new List<ComponentLinkEntry>(),
                "PAL",
                DataRoot,
                []);

            Assert.Equal("Unknown component", TextOf(window, "TitleText"));
        });
    }

    // -----------------------------------------------------------------------------------------
    // Region selection
    // -----------------------------------------------------------------------------------------

    // A board can carry a different part for PAL and NTSC. The popup must show the one matching
    // the region in force, or the user is reading the wrong part number for their machine.
    [Fact]
    public void The_entry_matching_the_current_region_is_the_one_shown()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            var entries = new[]
            {
                Component(friendlyName: "PAL PLA", technicalName: "251715", region: "PAL"),
                Component(friendlyName: "NTSC PLA", technicalName: "906114", region: "NTSC"),
            };

            SetComponent(window, entries, region: "NTSC");

            Assert.Equal("U8 | NTSC PLA | 906114", TextOf(window, "TitleText"));
        });
    }

    // A region-less entry is the generic fallback when nothing matches the current region.
    [Fact]
    public void A_region_less_entry_is_used_when_no_region_matches()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            var entries = new[]
            {
                Component(friendlyName: "PAL only", technicalName: "251715", region: "PAL"),
                Component(friendlyName: "Generic", technicalName: "0000", region: ""),
            };

            SetComponent(window, entries, region: "NTSC");

            Assert.Equal("U8 | Generic | 0000", TextOf(window, "TitleText"));
        });
    }

    // The region toggle buttons only make sense for a board that actually has region-specific
    // components; on any other board they are hidden rather than shown and inert.
    [Fact]
    public void The_region_buttons_are_hidden_when_the_board_has_no_region_components()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() }, hasExplicitRegionComponents: false);

            Assert.False(ControlOf<Button>(window, "PalRegionButton").IsVisible);
            Assert.False(ControlOf<Button>(window, "NtscRegionButton").IsVisible);
        });
    }

    [Fact]
    public void The_region_buttons_are_shown_when_the_board_has_region_components()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() }, hasExplicitRegionComponents: true);

            Assert.True(ControlOf<Button>(window, "PalRegionButton").IsVisible);
            Assert.True(ControlOf<Button>(window, "NtscRegionButton").IsVisible);
        });
    }

    // -----------------------------------------------------------------------------------------
    // The info lines
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_category_and_part_number_line_joins_both_when_present()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component(category: "IC", partNumber: "906114-01") });

            Assert.True(ControlOf<TextBlock>(window, "InfoCategoryPartNumber").IsVisible);
            Assert.Equal("IC | 906114-01", TextOf(window, "InfoCategoryPartNumber"));
        });
    }

    // With neither value the whole line is hidden - an empty line would leave a gap in the
    // layout that reads as a rendering fault.
    [Fact]
    public void The_category_and_part_number_line_is_hidden_when_both_are_blank()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component(category: "", partNumber: "") });

            Assert.False(ControlOf<TextBlock>(window, "InfoCategoryPartNumber").IsVisible);
        });
    }

    [Fact]
    public void The_description_section_shows_the_entrys_description()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component(description: "Programmable logic array") });

            Assert.True(ControlOf<Control>(window, "OneLinerSection").IsVisible);

            // A read-only TextBox rather than a TextBlock, so the description can be selected and
            // copied - part numbers and pin notes routinely get pasted into a search.
            Assert.Equal(
                "Programmable logic array",
                ControlOf<TextBox>(window, "InfoDescription").Text);
        });
    }

    // ###########################################################################################
    // "PURPOSE OR NOTE" IS SHOWN IN THE POPUP, FORMATTED AS TYPED (owner request, 2026-11-09: "It should be visible in the
    // component info popup"), and hidden without one. The description's heading says what the pane calls it.
    // ###########################################################################################
    [Fact]
    public void The_purpose_is_shown_formatted_and_hidden_without_one()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            ComponentEntry plain = Component();
            SetComponent(window, new[] { plain });
            Assert.False(ControlOf<Control>(window, "PurposeSection").IsVisible);

            var withPurpose = new ComponentEntry
            {
                BoardLabel = plain.BoardLabel,
                FriendlyName = plain.FriendlyName,
                Category = plain.Category,
                Description = plain.Description,
                Purpose = "Decodes the **memory map**\n- ROM and I/O",
            };

            SetComponent(window, new[] { withPurpose });

            Assert.True(ControlOf<Control>(window, "PurposeSection").IsVisible);
            Assert.Equal("Decodes the **memory map**\n- ROM and I/O", ControlOf<RichNoteTextBox>(window, "InfoPurpose").Markup);
            Assert.True(ControlOf<RichNoteTextBox>(window, "InfoPurpose").IsReadOnly);
        });
    }

    [Fact]
    public void The_description_section_is_hidden_when_there_is_no_description()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component(description: "   ") });

            Assert.False(ControlOf<Control>(window, "OneLinerSection").IsVisible);
        });
    }

    // ###########################################################################################
    // A BOARD OF ITS OWN REGIONS HAS A BUTTON PER REGION ON THE POPUP TOO (2026-11-09: "so we are prepared for other
    // regions"), each counting the pictures it shows, stacked from the footer's bottom; PAL and NTSC are not there for
    // such a board, and stand where they always stood for one in PAL and NTSC.
    // ###########################################################################################
    [Fact]
    public void A_board_of_its_own_regions_has_a_button_per_region_on_the_popup()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component(region: "EU"), Component(region: "US") }, region: "US", hasExplicitRegionComponents: true, regions: ["EU", "US"]);

            Assert.Equal(["EU (0)", "US (0)"], window.RegionButtonsShownForTests);
            Assert.False(ControlOf<Button>(window, "PalRegionButton").IsVisible);
            Assert.False(ControlOf<Button>(window, "NtscRegionButton").IsVisible);

            SetComponent(window, new[] { Component(region: "PAL") }, region: "PAL", hasExplicitRegionComponents: true);

            Assert.Equal(["PAL (0)", "NTSC (0)"], window.RegionButtonsShownForTests);
            Assert.Equal(2, Grid.GetRow(ControlOf<Button>(window, "PalRegionButton")));
            Assert.Equal(4, Grid.GetRow(ControlOf<Button>(window, "NtscRegionButton")));
            Assert.Contains("active", ControlOf<Button>(window, "PalRegionButton").Classes);
        });
    }

    // ###########################################################################################
    // A region the board gives colors of its own (owner request, 2026-11-11; RegionColors) is named in them over the popup's
    // picture, and its button carries them - a region with none keeps CRT's.
    // ###########################################################################################
    [Fact]
    public void A_region_in_the_boards_own_colors_is_named_in_them_on_the_popup()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            window.RegionColorRows = [new RegionColorEntry { Region = "US", Background = "#1F3A93" }];
            window.Show();

            SetComponent(window, new[] { Component(region: "EU"), Component(region: "US") }, region: "US", hasExplicitRegionComponents: true, regions: ["EU", "US"]);

            Border label = ControlOf<Border>(window, "InfoRegionBorder");
            TextBlock text = ControlOf<TextBlock>(window, "InfoRegionText");

            Assert.Equal(Avalonia.Media.Color.Parse("#1F3A93"), ((Avalonia.Media.ISolidColorBrush)label.Background!).Color);
            Assert.Equal(Avalonia.Media.Colors.White, ((Avalonia.Media.ISolidColorBrush)text.Foreground!).Color);

            SetComponent(window, new[] { Component(region: "EU"), Component(region: "US") }, region: "EU", hasExplicitRegionComponents: true, regions: ["EU", "US"]);

            Assert.True(Application.Current!.TryGetResource("Schematics_Region_Own1_Bg", Avalonia.Styling.ThemeVariant.Light, out object? own1));
            Assert.Equal(((Avalonia.Media.ISolidColorBrush)own1!).Color, ((Avalonia.Media.ISolidColorBrush)label.Background!).Color);

            window.Close();
        });
    }

    // ###########################################################################################
    // THE POPUP'S REGION BUTTONS LAID OUT AS ONE BOARD FOLLOWS ANOTHER, THE POPUP SHOWN (owner report, 2026-11-09: CRT threw
    // at launch - Avalonia's Grid measured a button against a column it had lost; RegionButtonLayout). Its two columns are
    // fixed: a fourth region stands in the second, 6px off the first, and with three or fewer that column is empty.
    // ###########################################################################################
    [Fact]
    public void The_popups_region_buttons_are_laid_out_as_boards_of_every_kind_follow_each_other()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            window.Show();

            void Lay(ComponentEntry[] components, string region, IReadOnlyList<string>? regions)
            {
                SetComponent(window, components, region: region, hasExplicitRegionComponents: true, regions: regions);
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();
            }

            Lay([Component(region: "EU")], "EU", ["EU", "US", "UK", "Japan", "Korea"]);
            Button fourth = window.GetVisualDescendants().OfType<Button>().Single(button => button.Tag as string == "Japan");
            Assert.Equal(1, Grid.GetColumn(fourth));
            Assert.Equal(6, fourth.Margin.Left);

            Lay([Component(region: "PAL")], "PAL", null);
            Assert.Equal(["PAL (0)", "NTSC (0)"], window.RegionButtonsShownForTests);
            Assert.Equal(2, ControlOf<Grid>(window, "PopupRegionGrid").ColumnDefinitions.Count);

            Lay([Component(region: "EU")], "EU", ["EU", "US", "UK", "Japan"]);
            Lay([Component(region: "PAL")], "PAL", null);
            Assert.Equal(0, ControlOf<Button>(window, "PalRegionButton").Margin.Left);

            // A long name is said whole - the button widens past PAL's 70px rather than cutting it off (code review,
            // 2026-11-09), its block the very one it was made with, the counts changing only its text.
            Lay([Component(region: "Japan (NTSC-J)")], "Japan (NTSC-J)", ["Japan (NTSC-J)", "US"]);
            Button longName = window.GetVisualDescendants().OfType<Button>().Single(button => button.Tag as string == "Japan (NTSC-J)");
            TextBlock name = Assert.IsType<TextBlock>(longName.Content);
            Assert.Equal("Japan (NTSC-J) (0)", name.Text);
            Assert.True(longName.Bounds.Width > 70, $"The button is {longName.Bounds.Width}px");
            Assert.True(name.Bounds.Height < 20, $"The name wraps to {name.Bounds.Height}px");

            Lay([Component(region: "US")], "US", ["Japan (NTSC-J)", "US"]);
            Assert.Same(name, longName.Content);

            window.Close();
        });
    }

    // -----------------------------------------------------------------------------------------
    // CRT 4.0's new component data (step 2, slice 3c)
    // -----------------------------------------------------------------------------------------

    // The package joins the category line.
    [Fact]
    public void The_package_joins_the_category_line()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[]
            {
                new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", PartNumber = "906114-01", Package = "DIP-28" },
            });

            Assert.Equal("IC | 906114-01 | DIP-28", TextOf(window, "InfoCategoryPartNumber"));
        });
    }

    [Fact]
    public void Without_new_data_the_new_sections_stay_hidden()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() });

            Assert.False(ControlOf<Control>(window, "PinoutSection").IsVisible);
            Assert.False(ControlOf<Control>(window, "ReplacementsSection").IsVisible);
            Assert.False(ControlOf<Control>(window, "BaselineConditionsSection").IsVisible);
        });
    }

    // ###########################################################################################
    // A SINGLE RESISTOR IS ITS COLOR BANDS, IN THE PINOUT (owner requests, 2026-11-07: "automatically create such an image
    // and show it under the pinout info in component info popup"; 2026-11-09: "the image/drawing of the resistor should be in
    // the pinout instead (like the other components you are drawing, including the explanation of colors"): drawn from its
    // value in the pinout section, where the case or a leaded part is drawn, with the value in words as the owner writes it
    // and its power rating apart - the section shown for it with no pins listed - and none for a resistor pack, or a value
    // the bands cannot say.
    // ###########################################################################################
    [Fact]
    public void A_resistor_shows_its_color_bands_and_value_in_the_pinout_and_a_pack_none()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", TechnicalNameOrValue = "560Ω 1/4W 5%" } });

            ResistorPicture resistor = ControlOf<ResistorPicture>(window, "InfoResistorPicture");
            Assert.True(resistor.IsVisible);
            Assert.True(ControlOf<Control>(window, "PinoutSection").IsVisible);
            Assert.Same(ControlOf<Grid>(window, "PinoutGrid"), resistor.Parent);
            Assert.Equal(
                [ResistorColorCode.Band.Green, ResistorColorCode.Band.Blue, ResistorColorCode.Band.Brown, ResistorColorCode.Band.Gold],
                resistor.BandsShownForTests.Select(band => band.Color));
            Assert.Equal("560 Ω 5%", resistor.ValueShownForTests);
            Assert.Equal("Power rating: 1/4W", resistor.PowerShownForTests);
            Assert.Equal(["Green", "Blue", "Brown", "Gold"], resistor.ColorsShownForTests);

            // A value with no ohm sign - a board made from KiCad data writes "8.2k" - reads the same (owner request,
            // 2026-11-07).
            SetComponent(window, new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", TechnicalNameOrValue = "8.2k" } });
            Assert.True(resistor.IsVisible);
            Assert.Equal("8.2K Ω", resistor.ValueShownForTests);
            Assert.Equal(["Grey", "Red", "Red"], resistor.ColorsShownForTests);

            SetComponent(window, new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", FriendlyName = "Resistor pack", TechnicalNameOrValue = "1KΩ, 6pin" } });
            Assert.False(resistor.IsVisible);

            SetComponent(window, new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", TechnicalNameOrValue = "wire" } });
            Assert.False(resistor.IsVisible);
            Assert.False(ControlOf<Control>(window, "PinoutSection").IsVisible);

            SetComponent(window, new[] { Component() });
            Assert.False(resistor.IsVisible);
        });
    }

    // ###########################################################################################
    // AN AXIAL CERAMIC CAPACITOR IS ITS COLOR BANDS TOO (owner request, 2026-11-09: "since that has color rings, would it then
    // be possible to show and translate those also, just alike a resistor?"): 250407's C83 as the data has it - "Ceramic",
    // "Axial", "82pF 5%" - in the pinout's place, as the owner's own parts carry their bands (CapacitorColorCode): the digits,
    // the multiplier, the tolerance in the resistor's colors, and a fifth band not known and said so, on the light-green body;
    // the plain tube not drawn beside it. A film one keeps the tube.
    // ###########################################################################################
    [Fact]
    public void An_axial_ceramic_capacitor_shows_its_color_bands_with_its_fifth_band_not_known()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            ComponentPinEntry[] pins =
            [
                new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "", Direction = "" },
                new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "2", Name = "", Direction = "" },
            ];

            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "Capacitor", FriendlyName = "Ceramic", TechnicalNameOrValue = "82pF 5%", Package = "Axial" } },
                pins: pins);

            ResistorPicture bands = ControlOf<ResistorPicture>(window, "InfoResistorPicture");
            Assert.True(bands.IsVisible);
            Assert.False(ControlOf<Control>(window, "InfoPartDrawing").IsVisible);
            Assert.Equal(
                [ResistorColorCode.Band.Grey, ResistorColorCode.Band.Red, ResistorColorCode.Band.Black, ResistorColorCode.Band.Gold, ResistorColorCode.Band.Unknown],
                bands.BandsShownForTests.Select(band => band.Color));
            Assert.Equal("Capacitor value:", bands.HeadingShownForTests);
            Assert.Equal("82 pF 5%", bands.ValueShownForTests);
            Assert.Equal(["Grey", "Red", "Black", "Gold", "Not known (5th band)"], bands.ColorsShownForTests);
            Assert.Contains("5th band", bands.UnknownBandsNoteShownForTests, StringComparison.Ordinal);
            Assert.NotNull(bands.BodyColorForTests);

            // A film one has no color code: the plain tube, no bands.
            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "Capacitor", FriendlyName = "Film", TechnicalNameOrValue = "16pF 5%", Package = "Axial" } },
                pins: pins);
            Assert.False(bands.IsVisible);
            Assert.True(ControlOf<Control>(window, "InfoPartDrawing").IsVisible);

            // A resistor after it: a resistor's heading, body and no line about bands not known.
            SetComponent(window, new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", TechnicalNameOrValue = "560Ω 1/4W 5%" } });
            Assert.Equal("Resistor value:", bands.HeadingShownForTests);
            Assert.Equal(string.Empty, bands.UnknownBandsNoteShownForTests);
            Assert.Null(bands.BodyColorForTests);
            Assert.Equal("Power rating: 1/4W", bands.PowerShownForTests);
        });
    }

    // ###########################################################################################
    // "DETACH PINOUT IMAGE IN COMPONENT INFO POPUP" (owner request, 2026-11-07: "do show the pinout and its information in
    // another popup window and do make sure to save size and location etc. (one global, and not per component/board)"):
    // the popup's pinout section stands in a window of its own while the component has pins, back in the popup when it
    // has none or the setting is unticked; the window's size is saved as it closes, one for all.
    // ###########################################################################################
    [Fact]
    public void A_detached_pinout_stands_in_a_window_of_its_own_and_its_size_is_kept()
    {
        UiTest.Run(() =>
        {
            // Forty pins: more than the popup has room for, so the window is wanted (PinoutFitsInPopup).
            ComponentPinEntry[] pins = FortyPins();
            ComponentEntry[] chip = [new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-40" }];

            UserSettings.DetachPinoutInComponentPopup = true;

            var window = CreateWindow();
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Control section = ControlOf<Control>(window, "PinoutSection");

            SetComponent(window, chip, pins: pins);
            Dispatcher.UIThread.RunJobs();

            ComponentPinoutWindow pinout = Assert.IsType<ComponentPinoutWindow>(window.PinoutWindowForTests);
            Assert.True(pinout.IsVisible);
            Assert.Same(pinout, TopLevel.GetTopLevel(section));
            Assert.StartsWith("Pinout - ", pinout.Title, StringComparison.Ordinal);

            // Its size the first time (owner request, 2026-11-07: "aprox 870x960px").
            Assert.Equal((ComponentPinoutWindow.DefaultWidth, ComponentPinoutWindow.DefaultHeight), (pinout.Width, pinout.Height));

            // No pins: no window, the section back in the popup - decided once the caller is done with the popup, before
            // anything is drawn (the jobs at Normal; the section's move is a Background job of its own).
            SetComponent(window, new[] { Component() });
            Dispatcher.UIThread.RunJobs(DispatcherPriority.Normal);
            Assert.Null(window.PinoutWindowForTests);
            Assert.False(pinout.IsVisible);
            Dispatcher.UIThread.RunJobs();
            Assert.Same(window, TopLevel.GetTopLevel(section));

            // Unticked while open: back in the popup.
            SetComponent(window, chip, pins: pins);
            Dispatcher.UIThread.RunJobs(DispatcherPriority.Normal);
            Assert.NotNull(window.PinoutWindowForTests);
            UserSettings.DetachPinoutInComponentPopup = false;
            Dispatcher.UIThread.RunJobs();
            Assert.Null(window.PinoutWindowForTests);
            Assert.Same(window, TopLevel.GetTopLevel(section));

            // Ticked again, resized, and the popup closed: the size is kept for the next one.
            UserSettings.DetachPinoutInComponentPopup = true;
            Dispatcher.UIThread.RunJobs();
            pinout = Assert.IsType<ComponentPinoutWindow>(window.PinoutWindowForTests);
            pinout.Width = 512;
            Dispatcher.UIThread.RunJobs();

            window.Close();
            Dispatcher.UIThread.RunJobs();

            Assert.False(pinout.IsVisible);
            Assert.True(UserSettings.HasPinoutWindowLayout);
            Assert.Equal(512, UserSettings.PinoutWindowWidth);

            UserSettings.DetachPinoutInComponentPopup = false;
        });
    }

    // ###########################################################################################
    // *** A LAYOUT PASS BETWEEN EVERY MOVE (owner report, 2026-11-07: a resistor clicked with the pinout detached, and
    // CRT ended on "Attempt to call InvalidateArrange on wrong LayoutManager"). *** Avalonia keeps a control queued for
    // the layout of the window it stood in after it is taken out, and arranging it there once it stands in another window
    // throws. So the popup's rows, rebuilt for the next component, were still the popup's to arrange as the section moved
    // into the pinout window - and the other way round. Every move here is followed by both windows' layout, as on screen.
    // ###########################################################################################
    [Fact]
    public void The_pinout_moves_between_its_windows_with_the_layout_of_both_running_in_between()
    {
        UiTest.Run(() =>
        {
            ComponentPinEntry[] chipPins = [new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "VCC", Direction = "Power" }];
            ComponentPinEntry[] resistorPins =
            [
                new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "1", Direction = "Unknown" },
                new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "2", Name = "2", Direction = "Unknown" },
            ];
            ComponentEntry[] chip = [new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" }];
            ComponentEntry[] resistor = [new ComponentEntry { BoardLabel = BoardLabel, Category = "Resistor", Package = "Axial", TechnicalNameOrValue = "100k" }];

            UserSettings.DetachPinoutInComponentPopup = true;

            // Jobs, then a render pass - each window's layout - then what it posted.
            static void Settle()
            {
                Dispatcher.UIThread.RunJobs();
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                Dispatcher.UIThread.RunJobs();
            }

            try
            {
                var window = CreateWindow();
                window.Width = 320;

                // At its lowest the popup has room for none of these pinouts, so each goes into its window (PinoutFitsInPopup).
                window.Height = window.MinHeight;
                SetComponent(window, new[] { Component() });
                window.Show();
                Settle();

                Control section = ControlOf<Control>(window, "PinoutSection");

                // No pins, then pins: out into a window of its own, the rows just made - a popup this narrow lists the pins
                // under the drawing, a window that wide beside it, so the move changes the section's layout. Taken out at
                // once, it stands in neither window until the popup has had its layout.
                SetComponent(window, resistor, pins: resistorPins);
                Dispatcher.UIThread.RunJobs(DispatcherPriority.Normal);
                Assert.Null(section.Parent);
                Assert.NotNull(window.PinoutWindowForTests);
                Settle();
                Assert.Same(window.PinoutWindowForTests, TopLevel.GetTopLevel(section));

                // Other pins, in the window.
                SetComponent(window, chip, pins: chipPins);
                Settle();

                // No pins: back into the popup, the rows just cleared.
                SetComponent(window, new[] { Component() });
                Settle();
                Assert.Same(window, TopLevel.GetTopLevel(section));

                // And out again, and back by the setting.
                SetComponent(window, resistor, pins: resistorPins);
                Settle();
                UserSettings.DetachPinoutInComponentPopup = false;
                Settle();
                Assert.Same(window, TopLevel.GetTopLevel(section));

                window.Close();
                Settle();
            }
            finally
            {
                UserSettings.DetachPinoutInComponentPopup = false;
            }
        });
    }

    // ###########################################################################################
    // *** ONLY A PINOUT THE POPUP HAS NO ROOM FOR GETS A WINDOW OF ITS OWN (owner request, 2026-11-09: "if no vertical
    // scrolling will be implied with this info inside of component info popup, then you will not open a new window").
    // *** With the setting ticked, a small pinout stays in the popup, which then has nothing to scroll; forty pins get the
    // window; and the small one shown after them comes back into the popup, measured while it stood in the window.
    // ###########################################################################################
    [Fact]
    public void Only_a_pinout_the_popup_has_no_room_for_gets_a_window_of_its_own()
    {
        UiTest.Run(() =>
        {
            ComponentEntry[] small = [new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" }];
            ComponentPinEntry[] smallPins =
            [
                new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "VCC", Direction = "Power" },
                new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "8", Name = "GND", Direction = "Power" },
            ];
            ComponentEntry[] large = [new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-40" }];

            UserSettings.DetachPinoutInComponentPopup = true;

            try
            {
                var window = CreateWindow();
                window.Height = 900;
                window.Show();
                Dispatcher.UIThread.RunJobs();

                Control section = ControlOf<Control>(window, "PinoutSection");
                ScrollViewer info = ControlOf<ScrollViewer>(window, "InfoScrollViewer");

                SetComponent(window, small, pins: smallPins);
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                Assert.Null(window.PinoutWindowForTests);
                Assert.Same(window, TopLevel.GetTopLevel(section));
                Assert.True(section.IsVisible);
                Assert.True(info.Extent.Height <= info.Viewport.Height, $"{info.Extent.Height} > {info.Viewport.Height}");

                SetComponent(window, large, pins: FortyPins());
                Dispatcher.UIThread.RunJobs();

                Assert.NotNull(window.PinoutWindowForTests);
                Assert.Same(window.PinoutWindowForTests, TopLevel.GetTopLevel(section));

                SetComponent(window, small, pins: smallPins);
                Dispatcher.UIThread.RunJobs();

                Assert.Null(window.PinoutWindowForTests);
                Assert.Same(window, TopLevel.GetTopLevel(section));

                window.Close();
            }
            finally
            {
                UserSettings.DetachPinoutInComponentPopup = false;
            }
        });
    }

    // ###########################################################################################
    // *** THE TALLEST NOTE OF ITS IMAGES COUNTS (owner request, 2026-11-09: "so the user can always see the pins when
    // navigating in the oscilloscope baseline"). *** The images' notes are shown one at a time as they are gone through,
    // so a pinout with room beside the first image's note can be pushed out of sight by a longer one further on: the
    // longest decides, before any image is shown, and the window is there from the start.
    // ###########################################################################################
    [Fact]
    public void The_longest_note_among_the_images_decides_so_the_pins_stay_in_sight_on_every_image()
    {
        UiTest.Run(() =>
        {
            ComponentEntry[] chip = [new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" }];
            ComponentPinEntry[] pins = [new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "VCC", Direction = "Power" }];
            string longNote = string.Join("\n", Enumerable.Range(1, 60).Select(line => $"- What pin {line} reads when the machine is on"));

            ComponentImageEntry Image(string file, string note, string pin = "") =>
                new() { BoardLabel = BoardLabel, Name = file, File = file, Note = note, Pin = pin };

            UserSettings.DetachPinoutInComponentPopup = true;

            try
            {
                var window = CreateWindow();
                window.Height = 900;
                window.Show();
                Dispatcher.UIThread.RunJobs();

                Control section = ControlOf<Control>(window, "PinoutSection");

                // A short note on every image: room enough.
                SetComponent(window, chip, pins: pins, images: [Image("pinout.png", "Seen from above"), Image("vcc.png", "5 V", pin: "1")]);
                Dispatcher.UIThread.RunJobs();

                Assert.Null(window.PinoutWindowForTests);
                Assert.Same(window, TopLevel.GetTopLevel(section));

                // The second image's note, sixty lines long, would push the pins out of sight there.
                SetComponent(window, chip, pins: pins, images: [Image("pinout.png", "Seen from above"), Image("vcc.png", longNote, pin: "1")]);
                Dispatcher.UIThread.RunJobs();

                Assert.NotNull(window.PinoutWindowForTests);
                Assert.Same(window.PinoutWindowForTests, TopLevel.GetTopLevel(section));

                window.Close();
            }
            finally
            {
                UserSettings.DetachPinoutInComponentPopup = false;
            }
        });
    }

    // ###########################################################################################
    // *** DECIDED ONCE THE CALLER IS DONE WITH THE POPUP, BEFORE ANYTHING IS DRAWN (code review, 2026-11-09). *** Main calls
    // OfferEdit and UpdateOscilloscopeSessionTitleState after SetComponent, and "Edit component" shown takes room from the
    // title row: decided inside SetComponent, the pinout was judged against a popup that then changed. So nothing is decided
    // while SetComponent and what follows it run, and the decision comes at Normal - above Render, so before the next frame.
    // ###########################################################################################
    [Fact]
    public void Where_the_pinout_stands_is_decided_after_what_follows_SetComponent_and_before_drawing()
    {
        UiTest.Run(() =>
        {
            ComponentEntry[] large = [new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-40" }];

            UserSettings.DetachPinoutInComponentPopup = true;

            try
            {
                // At its smallest the popup has no room for forty pins, whatever the fonts (PinoutFitsInPopup) - at its
                // default size they fitted with Windows' fonts and not with Linux', and this test is about WHEN it is
                // decided, not whether it fits.
                var window = CreateWindow();
                window.Width = window.MinWidth;
                window.Height = window.MinHeight;
                window.Show();
                Dispatcher.UIThread.RunJobs();

                SetComponent(window, large, pins: FortyPins());
                window.OfferEdit(_ => { }, string.Empty);

                // Not yet: what follows SetComponent may still change the room.
                Assert.Null(window.PinoutWindowForTests);

                Dispatcher.UIThread.RunJobs(DispatcherPriority.Normal);
                Assert.NotNull(window.PinoutWindowForTests);

                // Run ahead of layout and drawing.
                Assert.True(DispatcherPriority.Normal > DispatcherPriority.Render);

                window.Close();
            }
            finally
            {
                UserSettings.DetachPinoutInComponentPopup = false;
            }
        });
    }

    private static ComponentPinEntry[] FortyPins() =>
        Enumerable.Range(1, 40)
            .Select(pin => new ComponentPinEntry { BoardLabel = BoardLabel, Pin = pin.ToString(System.Globalization.CultureInfo.InvariantCulture), Name = $"P{pin}", Direction = "Input" })
            .ToArray();

    // The pins of the region shown are listed, the case drawn; switching the popup's region lists
    // that region's.
    [Fact]
    public void The_pinout_lists_the_pins_of_the_region_shown_and_follows_the_region_buttons()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" } },
                hasExplicitRegionComponents: true,
                pins:
                [
                    new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "VCC", Direction = "Power" },
                    new ComponentPinEntry { BoardLabel = BoardLabel, Region = "NTSC", Pin = "2", Name = "NTSC clock" },
                    new ComponentPinEntry { BoardLabel = "U9", Pin = "1", Name = "Another chip" },
                ]);

            Assert.True(ControlOf<Control>(window, "PinoutSection").IsVisible);
            Assert.Equal([new ComponentPopupPin("1", "VCC", "Power")], ItemsOf<ComponentPopupPin>(window, "InfoPinsItemsControl"));
            Assert.Equal(8, ControlOf<PinoutDrawing>(window, "InfoPinoutDrawing").LayoutForTests!.Pins.Count);

            ControlOf<Button>(window, "NtscRegionButton").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            Assert.Equal(["1", "2"], ItemsOf<ComponentPopupPin>(window, "InfoPinsItemsControl").Select(pin => pin.Pin));

            // Without the board's KiCad data there are no nets, and no heading for them.
            Assert.True(ControlOf<Control>(window, "InfoPinsHeading").IsVisible);
            Assert.False(ControlOf<Control>(window, "InfoPinsNetHeading").IsVisible);
        });
    }

    // ###########################################################################################
    // With the board's KiCad data loaded, each pin shows its net on this board beside its signal name
    // (the plan: "side by side"), under headings saying which is which.
    // ###########################################################################################
    [Fact]
    public void Each_pin_shows_its_net_on_this_board_beside_its_signal_name()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" } },
                pins:
                [
                    new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "VCC", Direction = "Power" },
                    new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "2", Name = "PHI2", Direction = "Input" },
                ],
                pinNets: new Dictionary<string, string> { ["1"] = "+5V" });

            Assert.True(ControlOf<Control>(window, "InfoPinsNetHeading").IsVisible);
            Assert.Equal(
                [new ComponentPopupPin("1", "VCC", "Power", "+5V"), new ComponentPopupPin("2", "PHI2", "Input", "")],
                ItemsOf<ComponentPopupPin>(window, "InfoPinsItemsControl"));
        });
    }

    // ###########################################################################################
    // The last column says what the scope should read on each pin, from the component's oscilloscope baselines (owner
    // request, 2026-11-09: "add another column ... named 'Expected oscilloscope reading' ... The header should probably
    // be 3 lines"), and follows the popup's region; a component with no reading shows no heading for it, as "Net name".
    // ###########################################################################################
    [Fact]
    public void Each_pin_shows_its_expected_oscilloscope_reading_in_the_last_column()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            ComponentPinEntry[] pins =
            [
                new() { BoardLabel = BoardLabel, Pin = "1", Name = "VCC", Direction = "Power" },
                new() { BoardLabel = BoardLabel, Pin = "2", Name = "PHI2", Direction = "Input" },
            ];

            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" } },
                hasExplicitRegionComponents: true,
                pins: pins,
                images:
                [
                    new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "1", ExpectedOscilloscopeReading = "5.0V (+5VDC power rail)" },
                    new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "2", Region = "PAL", ExpectedOscilloscopeReading = "985.232 kHz" },
                    new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "2", Region = "NTSC", ExpectedOscilloscopeReading = "1.02272 MHz" },
                ]);

            TextBlock heading = ControlOf<TextBlock>(window, "InfoPinsReadingHeading");
            Assert.True(heading.IsVisible);
            Assert.Equal(["Expected", "oscilloscope", "reading"], heading.Text!.Split('\n'));
            Assert.Equal(4, Grid.GetColumn(heading));
            Assert.Equal(4, ControlOf<Grid>(window, "InfoPinsHeading").Children.OfType<TextBlock>().Max(Grid.GetColumn));
            Assert.Equal(["5.0V (+5VDC power rail)", "985.232 kHz"], ItemsOf<ComponentPopupPin>(window, "InfoPinsItemsControl").Select(pin => pin.Reading));

            ControlOf<Button>(window, "NtscRegionButton").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            Assert.Equal(["5.0V (+5VDC power rail)", "1.02272 MHz"], ItemsOf<ComponentPopupPin>(window, "InfoPinsItemsControl").Select(pin => pin.Reading));

            SetComponent(window, new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" } }, pins: pins);
            Assert.False(heading.IsVisible);
        });
    }

    // ###########################################################################################
    // THE PIN UNDER THE POINTER ON THE DRAWING IS LIT IN THE LIST TOO (owner request, 2026-11-09: "hovering the IC image
    // pin will not highlight the table pin. This needs to be the same both places") - through a shown window and real
    // pointer moves, its row's frame taking the look of a row under the pointer; off the pins, no row is lit.
    // ###########################################################################################
    [Fact]
    public void A_pin_pointed_at_on_the_chip_drawing_lights_its_row_in_the_list()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow { Width = 1200, Height = 900 };
            window.Show();

            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" } },
                pins: FortyPins().Take(8).ToArray());
            LayOut();

            PinoutDrawing drawing = ControlOf<PinoutDrawing>(window, "InfoPinoutDrawing");
            PinoutLayout layout = drawing.LayoutForTests!;
            PinoutPin pin3 = layout.Pins.Single(pin => pin.Number == 3);
            PinoutPin pin6 = layout.Pins.Single(pin => pin.Number == 6);

            window.MouseMove(drawing.TranslatePoint(new Point(2, pin3.Lead.Center.Y), window)!.Value);
            LayOut();

            Assert.Equal(["3"], ItemsOf<ComponentPopupPin>(window, "InfoPinsItemsControl").Where(pin => pin.IsPointedAt).Select(pin => pin.Pin));
            Assert.Equal(["3"], LitRows(window));

            window.MouseMove(drawing.TranslatePoint(new Point(layout.Size.Width - 2, pin6.Lead.Center.Y), window)!.Value);
            LayOut();
            Assert.Equal(["6"], LitRows(window));

            window.MouseMove(drawing.TranslatePoint(layout.Body.Center, window)!.Value);
            LayOut();
            Assert.Empty(LitRows(window));

            window.Close();
        });
    }

    // ###########################################################################################
    // EVERY COLUMN OF THE PIN LIST IS AS WIDE AS ITS LONGEST TEXT, THE SAME SPACE BETWEEN EACH TWO (owner request,
    // 2026-11-09: "align on empty spacing between columns - it should be the same for all, and I do like the spacing
    // between 'Pin' and 'Signal'. Should be determined by longest text in each column + that spacing"). Laid out in a
    // shown window: each heading starts where the column before it ends - its widest text, heading or cell, and the gap -
    // and every cell but the last column's has the gap after it. In a narrow popup nothing stands past the list's right
    // edge: the reading and the net wrap.
    // ###########################################################################################
    [Fact]
    public void Each_column_of_the_pin_list_is_its_longest_text_and_the_same_gap_wide()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow { Width = 1600, Height = 900 };
            window.Show();

            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" } },
                pins:
                [
                    new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "D6", Direction = "Unknown" },
                    new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "2", Name = "PHI-COLOR", Direction = "Input" },
                    new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "8", Name = "VCC", Direction = "Power" },
                ],
                pinNets: new Dictionary<string, string> { ["1"] = "D6", ["2"] = "Net-(FB23-Pad1)", ["8"] = "Net-(C101-Pad1)" },
                images:
                [
                    new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "1", ExpectedOscilloscopeReading = "Pulsing" },
                    new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "2", ExpectedOscilloscopeReading = "14.3181 MHz" },
                    new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "8", ExpectedOscilloscopeReading = "12.0V (+12VDC power rail)" },
                ]);
            LayOut();

            IReadOnlyList<PinTableColumn> columns = window.PinColumnsForTests;
            Grid heading = ControlOf<Grid>(window, "InfoPinsHeading");
            List<double> starts = Enumerable.Range(0, 5)
                .Select(column => heading.Children.OfType<TextBlock>().Single(block => Grid.GetColumn(block) == column).TranslatePoint(default, heading)!.Value.X)
                .ToList();

            Assert.Equal(5, columns.Count(column => column.IsShown));
            Assert.Equal(0, starts[0]);

            for (int column = 1; column < 5; column++)
            {
                Assert.Equal(starts[column - 1] + columns[column - 1].Widest + PinTableGeometry.Gap, starts[column], 0.5);
            }

            List<TextBlock> cells = ControlOf<Control>(window, "InfoPinsItemsControl").GetVisualDescendants().OfType<TextBlock>().ToList();
            Assert.Equal(3 * 4, cells.Count(cell => cell.Classes.Contains("PinCell") && cell.Margin == new Thickness(0, 0, PinTableGeometry.Gap, 0)));
            Assert.Equal(3, cells.Count(cell => !cell.Classes.Contains("PinCell") && cell.Margin == default));

            // A narrow popup: the list has less room than its texts whole, though more than the columns that never wrap -
            // the signal, the net and the reading wrap, and nothing stands past its edge.
            window.Width = 900;
            LayOut();

            Control list = ControlOf<Control>(window, "InfoPinsPanel");
            Assert.True(PinTableGeometry.WidthOf(columns) > list.Bounds.Width, "the list has less room than its texts whole");
            Assert.True(
                columns[0].Widest + columns[3].Widest + (4 * PinTableGeometry.Gap) + (3 * PinTableGeometry.Narrowest) < list.Bounds.Width,
                "the list has room for what never wraps");
            Assert.All(
                list.GetVisualDescendants().OfType<TextBlock>().Where(block => block.IsEffectivelyVisible),
                block => Assert.True(block.TranslatePoint(new Point(block.Bounds.Width, 0), list)!.Value.X <= list.Bounds.Width + 0.5, $"\"{block.Text}\" stands past the list's edge"));

            window.Close();
        });
    }

    private static void LayOut()
    {
        for (int i = 0; i < 3; i++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
    }

    // The pins whose row in the popup's list is lit from the drawing - by the class its frame carries.
    private static List<string> LitRows(ComponentInfoWindow window) =>
        window.GetVisualDescendants()
            .OfType<Border>()
            .Where(row => row.Classes.Contains("PopupPinRow") && row.Classes.Contains("pointed") && row.IsEffectivelyVisible)
            .Select(row => ((ComponentPopupPin)row.Tag!).Pin)
            .ToList();

    // ###########################################################################################
    // Under a scope baseline, how the board's baselines were measured: its "Scope measurements"
    // files, each a link (2026-10-26). Read once the picture has loaded - the section follows the
    // picture selected.
    // ###########################################################################################
    [Fact]
    public async Task A_scope_baseline_links_the_boards_scope_measurement_files()
    {
        string root = this.thisWorkspace.Path_("Data-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(root, "Scope baseline"));
        File.WriteAllBytes(Path.Combine(root, "Scope baseline", "U8_1.png"), Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="));

        await UiTest.RunAsync(async () =>
        {
            var window = CreateWindow();

            window.SetComponent(
                BoardLabel,
                "U8",
                [Component()],
                [new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "1", Name = "PHI2", File = "Scope baseline/U8_1.png", TimeDiv = "1us" }],
                [],
                [],
                "PAL",
                root,
                regions: [],
                boardFiles:
                [
                    new BoardLocalFileEntry { Category = "Scope measurements", Name = "Introduction", File = "Scope baseline/intro.txt" },
                    new BoardLocalFileEntry { Category = "Resources", Name = "Manual", File = "Board local files/manual.pdf" },
                ]);

            for (int attempt = 0; attempt < 100 && ItemsOf<ComponentImageItem>(window, "ThumbnailList").Count == 0; attempt++)
            {
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.True(ControlOf<Control>(window, "BaselineConditionsSection").IsVisible);
            Assert.True(ControlOf<Control>(window, "InfoMeasuringNotesLabel").IsVisible);
            Assert.Equal(["Introduction"], ItemsOf<ComponentLocalFileItem>(window, "InfoMeasuringNotesItemsControl").Select(note => note.Name));
        });
    }

    // A note of one region is linked only under that region's baselines (code review, 2026-11-02) - the
    // popup's own PAL/NTSC switch moves it, as it moves the component's files.
    [Fact]
    public async Task A_scope_measurement_file_of_the_other_region_is_not_linked()
    {
        string root = this.thisWorkspace.Path_("Data-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(root, "Scope baseline"));
        File.WriteAllBytes(Path.Combine(root, "Scope baseline", "U8_1.png"), Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="));

        await UiTest.RunAsync(async () =>
        {
            var window = CreateWindow();

            window.SetComponent(
                BoardLabel,
                "U8",
                [Component()],
                [new ComponentImageEntry { BoardLabel = BoardLabel, Pin = "1", Name = "PHI2", File = "Scope baseline/U8_1.png", TimeDiv = "1us" }],
                [],
                [],
                "PAL",
                root,
                regions: BoardRegions.PalAndNtsc,
                boardFiles:
                [
                    new BoardLocalFileEntry { Category = "Scope measurements", Name = "Every region", File = "Scope baseline/every.txt" },
                    new BoardLocalFileEntry { Category = "Scope measurements", Name = "NTSC's", File = "Scope baseline/ntsc.txt", Region = "NTSC" },
                ]);

            for (int attempt = 0; attempt < 100 && ItemsOf<ComponentImageItem>(window, "ThumbnailList").Count == 0; attempt++)
            {
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal(["Every region"], ItemsOf<ComponentLocalFileItem>(window, "InfoMeasuringNotesItemsControl").Select(note => note.Name));
        });
    }

    [Fact]
    public void The_replacement_parts_are_listed()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(
                window,
                new[] { Component() },
                replacements: [new ComponentReplacementEntry { BoardLabel = BoardLabel, PartNumber = "PLAnkton", Kind = "Drop-in replacement" }]);

            Assert.True(ControlOf<Control>(window, "ReplacementsSection").IsVisible);
            Assert.Equal("PLAnkton - Drop-in replacement", Assert.Single(ItemsOf<ComponentPopupReplacement>(window, "InfoReplacementsItemsControl")).Title);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Local files and links
    // -----------------------------------------------------------------------------------------

    // Both lists are filtered by board label: the board's full file and link tables are passed in,
    // and only the rows belonging to THIS component may appear.
    [Fact]
    public void Only_this_components_local_files_are_listed()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() }, localFiles: new[]
            {
                LocalFile(BoardLabel, "Datasheet", "docs/pla.pdf"),
                LocalFile("U9", "Other datasheet", "docs/other.pdf"),
            });

            var items = ItemsOf<ComponentLocalFileItem>(window, "LocalFilesItemsControl");

            Assert.Single(items);
            Assert.Equal("Datasheet", items[0].Name);
        });
    }

    // The stored path is relative to the data root; the popup has to resolve it before anything
    // can open it.
    [Fact]
    public void A_local_files_path_is_resolved_against_the_data_root()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() }, localFiles: new[]
            {
                LocalFile(BoardLabel, "Datasheet", "docs/pla.pdf"),
            });

            var items = ItemsOf<ComponentLocalFileItem>(window, "LocalFilesItemsControl");

            Assert.Equal(Path.Combine(DataRoot, "docs", "pla.pdf"), items[0].FullPath);
        });
    }

    [Fact]
    public void The_local_files_section_is_hidden_when_this_component_has_none()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() }, localFiles: new[]
            {
                LocalFile("U9", "Someone elses", "docs/other.pdf"),
            });

            Assert.False(ControlOf<Control>(window, "LocalFilesSection").IsVisible);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Draft-fallback file resolution (NewContributeStrategy.md Phase 2, session 2c, task 8's
    // component/file/link attachment piece). A file attached only in a local draft has never been
    // synced under DataRoot, so it does not exist there at all - SetComponent must still resolve
    // it, via the draft folder DraftFileResolver falls back to, and record which root it actually
    // came from (SourceRoot) so a later open can be validated against the right containment root.
    // -----------------------------------------------------------------------------------------

    // ###########################################################################################
    // The ordinary case for a freshly attached file: nothing under DataRoot, a real copy inside
    // the draft folder.
    //
    // *** THE DRAFT FOLDER MUST BE A REALISTIC Manufacturer/Hardware/Board PATH (Phase 6). ***
    // DraftFileResolver strips the board's own prefix off a stored path, and it works out what
    // that prefix is from the folder's last three segments. A flat "draft" folder - which this
    // test used before - has no such segments, so nothing strips and the file is looked for in
    // the wrong place.
    // ###########################################################################################
    [Fact]
    public void A_draft_only_local_file_resolves_to_the_draft_folder_copy()
    {
        UiTest.Run(() =>
        {
            string draftFolder = Path.Combine(
                this.thisWorkspace.Path_("Drafts"), "Commodore", "C64", "250407");

            string draftedFile = Path.Combine(draftFolder, "docs", "pla.pdf");
            Directory.CreateDirectory(Path.GetDirectoryName(draftedFile)!);
            File.WriteAllText(draftedFile, "drafted copy");

            var window = CreateWindow();

            SetComponent(
                window,
                new[] { Component() },
                localFiles: new[] { LocalFile(BoardLabel, "Datasheet", "Commodore/C64/250407/docs/pla.pdf") },
                draftBoardFolder: draftFolder);

            var items = ItemsOf<ComponentLocalFileItem>(window, "LocalFilesItemsControl");

            Assert.Equal(draftedFile, items[0].FullPath);
            Assert.Equal(draftFolder, items[0].SourceRoot);
        });
    }

    // ###########################################################################################
    // *** THE DRAFTED COPY WINS WHEN BOTH EXIST - FLIPPED IN PHASE 6 (2026-09-23). ***
    //
    // It used to assert the official copy won, matching DraftFileResolver's official-first
    // rule. That rule was right while a draft was an OVERLAY holding only newly attached
    // extras; a draft is now a full copy of the board, so a drafted file is a REPLACEMENT and
    // showing the published bytes would put the old picture on screen for an image the
    // contributor has already replaced.
    //
    // *** THIS TEST ALSO USED A FLAT "draft" FOLDER, WHICH MADE IT PASS FOR THE WRONG
    // REASON. *** DraftFileResolver works out the board prefix to strip from the folder's
    // last three segments, so a one-segment folder strips nothing and the drafted file is
    // never found at all - the test went green because the lookup failed, not because
    // precedence worked. The folder is now a realistic Manufacturer/Hardware/Board path.
    // ###########################################################################################
    [Fact]
    public void A_drafted_local_file_is_preferred_over_a_same_named_official_copy()
    {
        UiTest.Run(() =>
        {
            string dataRoot = this.thisWorkspace.Path_("data");
            string officialFile = Path.Combine(
                dataRoot, "Commodore", "C64", "250407", "docs", "pla.pdf");
            Directory.CreateDirectory(Path.GetDirectoryName(officialFile)!);
            File.WriteAllText(officialFile, "official copy");

            string draftFolder = Path.Combine(
                this.thisWorkspace.Path_("Drafts"), "Commodore", "C64", "250407");

            string draftedFile = Path.Combine(draftFolder, "docs", "pla.pdf");
            Directory.CreateDirectory(Path.GetDirectoryName(draftedFile)!);
            File.WriteAllText(draftedFile, "drafted copy");

            var window = CreateWindow();

            window.SetComponent(
                BoardLabel,
                "U8",
                new[] { Component() }.ToList(),
                new List<ComponentImageEntry>(),
                new List<ComponentLocalFileEntry>
                {
                    LocalFile(BoardLabel, "Datasheet", "Commodore/C64/250407/docs/pla.pdf"),
                },
                new List<ComponentLinkEntry>(),
                "PAL",
                dataRoot,
                [],
                draftFolder);

            var items = ItemsOf<ComponentLocalFileItem>(window, "LocalFilesItemsControl");

            Assert.Equal(draftedFile, items[0].FullPath);

            // ###########################################################################################
            // *** AND THE ROOT MUST BE THE DRAFT FOLDER TOO (code review, 2026-09-25). ***
            //
            // This test used to stop at FullPath, and that let a real defect through: SetComponent
            // derived SourceRoot with the OLD official-first assumption ("drafted only when the
            // official copy is missing"), so with both copies present the path was the draft's
            // while the root was the data root. OnLocalFileButtonClick hands that pair to
            // ExternalTargetLauncher, whose containment check refuses a path outside its root - so
            // the datasheet button did nothing at all on every board that had a draft, which is
            // every file a seeded draft copies. The second assertion is the invariant the launcher
            // actually needs: the path lies inside the root it is opened against.
            // ###########################################################################################
            Assert.Equal(draftFolder, items[0].SourceRoot);
            Assert.StartsWith(
                items[0].SourceRoot + Path.DirectorySeparatorChar,
                items[0].FullPath,
                StringComparison.OrdinalIgnoreCase);
        });
    }

    // Neither copy exists (a stale entry, or a draft that was discarded after the popup was last
    // opened): SetComponent must still fall back to building the ordinary DataRoot-relative path
    // rather than leaving FullPath blank, and SourceRoot names DataRoot as the (missing) source.
    [Fact]
    public void A_local_file_missing_from_both_roots_still_falls_back_to_the_data_root_path()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(
                window,
                new[] { Component() },
                localFiles: new[] { LocalFile(BoardLabel, "Datasheet", "docs/pla.pdf") },
                draftBoardFolder: this.thisWorkspace.Path_("nonexistent-draft"));

            var items = ItemsOf<ComponentLocalFileItem>(window, "LocalFilesItemsControl");

            Assert.Equal(Path.Combine(DataRoot, "docs", "pla.pdf"), items[0].FullPath);
            Assert.Equal(DataRoot, items[0].SourceRoot);
        });
    }

    [Fact]
    public void Only_this_components_links_are_listed()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() }, links: new[]
            {
                Link(BoardLabel, "Pinout", "https://example.com/pla"),
                Link("U9", "Other", "https://example.com/other"),
            });

            var items = ItemsOf<ComponentLinkItem>(window, "LinksItemsControl");

            Assert.Single(items);
            Assert.Equal("Pinout", items[0].Name);
            Assert.Equal("https://example.com/pla", items[0].Url);
        });
    }

    [Fact]
    public void The_links_section_is_hidden_when_this_component_has_none()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(window, new[] { Component() }, links: Array.Empty<ComponentLinkEntry>());

            Assert.False(ControlOf<Control>(window, "LinksSection").IsVisible);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Reloading
    // -----------------------------------------------------------------------------------------

    // The popup is reused for the next component clicked rather than reopened, so a second
    // SetComponent must fully replace the first one's content - not merge with it.
    [Fact]
    public void Loading_a_second_component_replaces_the_first_ones_content()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(
                window,
                new[] { Component(friendlyName: "SuperPLA", technicalName: "251715") },
                links: new[] { Link(BoardLabel, "Pinout", "https://example.com/pla") });

            SetComponent(
                window,
                new[] { Component(friendlyName: "Other", technicalName: "0001") },
                links: Array.Empty<ComponentLinkEntry>());

            Assert.Equal("U8 | Other | 0001", TextOf(window, "TitleText"));
            Assert.False(ControlOf<Control>(window, "LinksSection").IsVisible);
        });
    }

    // ###########################################################################################
    // "EDIT COMPONENT" CLICKS ON THE PRESS THAT BRINGS THE POPUP TO THE FRONT (owner report, 2026-11-08: "it does not
    // react every time I click the 'Edit component' button ... it did that the second time I pressed it"). With CRT's
    // window in front, that press activates the popup, whose Activated handler puts the keyboard back in it - posted, so
    // it ran AFTER the press had put the focus on the button, and a Button that loses the focus lets go of its press
    // (Avalonia 12.1.4's Button.OnLostFocus): the release clicked nothing. The second press found the popup in front.
    // ###########################################################################################
    [Fact]
    public void Edit_component_clicks_on_the_press_that_brings_the_popup_to_the_front()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            string? edited = null;

            window.Show();
            SetComponent(window, new[] { Component() });
            window.OfferEdit(label => edited = label, boardShown: string.Empty);
            Dispatcher.UIThread.RunJobs();

            // CRT's window in front, as when a component was clicked on the schematic.
            var main = new Window { Width = 300, Height = 200 };
            main.Show();
            main.Activate();
            Dispatcher.UIThread.RunJobs();

            bool activated = false;
            window.Activated += (_, _) => activated = true;

            Button edit = ControlOf<Button>(window, "EditComponentButton");
            Assert.True(edit.IsVisible);
            Avalonia.Point centre = edit.TranslatePoint(new Avalonia.Point(edit.Bounds.Width / 2, edit.Bounds.Height / 2), window)!.Value;

            window.MouseDown(centre, MouseButton.Left);
            Assert.True(edit.IsPressed);

            // Windows activates the popup with that press, and the dispatcher runs what Activated posted before the release.
            window.Activate();
            Dispatcher.UIThread.RunJobs();
            Assert.True(activated);

            window.MouseUp(centre, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardLabel, edited);

            main.Close();
        });
    }

    // A popup opened, or brought back with nothing in it focused, still takes the keyboard from CRT's component filter.
    [Fact]
    public void A_popup_brought_to_the_front_with_nothing_focused_in_it_takes_the_keyboard()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            window.Show();
            SetComponent(window, new[] { Component() });
            Dispatcher.UIThread.RunJobs();

            Assert.True(window.IsKeyboardFocusWithin);
            window.Close();
        });
    }

    // -----------------------------------------------------------------------------------------
    // Fixture
    // -----------------------------------------------------------------------------------------

    private static ComponentInfoWindow CreateWindow() => new();

    // Drives the real public entry point. Image entries default to none: loading them is async
    // and needs files on disk, which is out of scope here (see the header).
    private static void SetComponent(
        ComponentInfoWindow window,
        IEnumerable<ComponentEntry> entries,
        string displayText = "U8",
        IEnumerable<ComponentLocalFileEntry>? localFiles = null,
        IEnumerable<ComponentLinkEntry>? links = null,
        string region = "PAL",
        bool hasExplicitRegionComponents = false,
        string draftBoardFolder = "",
        IReadOnlyList<ComponentPinEntry>? pins = null,
        IReadOnlyList<ComponentReplacementEntry>? replacements = null,
        IReadOnlyDictionary<string, string>? pinNets = null,
        IReadOnlyList<string>? regions = null,
        IEnumerable<ComponentImageEntry>? images = null)
    {
        window.SetComponent(
            BoardLabel,
            displayText,
            entries.ToList(),
            (images ?? Array.Empty<ComponentImageEntry>()).ToList(),
            (localFiles ?? Array.Empty<ComponentLocalFileEntry>()).ToList(),
            (links ?? Array.Empty<ComponentLinkEntry>()).ToList(),
            region,
            DataRoot,
            regions ?? (hasExplicitRegionComponents ? BoardRegions.PalAndNtsc : []),
            draftBoardFolder,
            boardIdentity: string.Empty,
            pins,
            replacements,
            boardFiles: null,
            pinNets: pinNets);
    }

    private static ComponentEntry Component(
        string friendlyName = "SuperPLA",
        string technicalName = "251715",
        string category = "IC",
        string partNumber = "906114-01",
        string description = "Programmable logic array",
        string region = "") =>
        new()
        {
            BoardLabel = BoardLabel,
            FriendlyName = friendlyName,
            TechnicalNameOrValue = technicalName,
            Category = category,
            PartNumber = partNumber,
            Description = description,
            Region = region,
        };

    private static ComponentLocalFileEntry LocalFile(string boardLabel, string name, string file) =>
        new() { BoardLabel = boardLabel, Name = name, File = file };

    private static ComponentLinkEntry Link(string boardLabel, string name, string url) =>
        new() { BoardLabel = boardLabel, Name = name, Url = url };

    private static T ControlOf<T>(ComponentInfoWindow window, string name) where T : Control =>
        window.GetControl<T>(name);

    private static string TextOf(ComponentInfoWindow window, string name) =>
        window.GetControl<TextBlock>(name).Text ?? string.Empty;

    private static List<T> ItemsOf<T>(ComponentInfoWindow window, string name)
    {
        var itemsSource = window.GetControl<ItemsControl>(name).ItemsSource;

        return itemsSource is null
            ? new List<T>()
            : itemsSource.Cast<T>().ToList();
    }

    // -----------------------------------------------------------------------------------------
    // The captured oscilloscope image, and releasing it again
    // -----------------------------------------------------------------------------------------

    /// <summary>A real PNG on disk, so the show path performs a genuine decode as a capture does.</summary>
    private string WriteCaptureImage()
    {
        string path = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".png");

        using var renderTarget = new Avalonia.Media.Imaging.RenderTargetBitmap(
            new Avalonia.PixelSize(24, 24), new Avalonia.Vector(96, 96));

        using (var context = renderTarget.CreateDrawingContext())
        {
            context.DrawRectangle(Avalonia.Media.Brushes.White, null, new Avalonia.Rect(0, 0, 24, 24));
        }

        using (var stream = File.Create(path))
        {
            renderTarget.Save(stream, Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }

        return path;
    }

    // ###########################################################################################
    // RELEASING THE CAPTURE MUST DROP IT OUT OF THE IMAGE CONTROL BEFORE DISPOSING IT.
    //
    // The clear disposes the decoded bitmap, and the popup's large Image was still pointing at it.
    // Every long-standing caller happened to reassign that Source immediately afterwards, so the
    // freed surface was never rendered - but the attach flow's own caller does not: it re-shows the
    // banner instead, which forces a layout pass over the still-visible Image. That is an
    // ObjectDisposedException on the RENDER thread, which is fatal in Avalonia - the same failure
    // mode the Workbooks board pane carries a warning about in its detach path.
    //
    // So the release now happens inside the clear itself rather than depending on each caller to
    // remember. Asserting on the Source is what makes that provable: a disposed bitmap cannot be
    // rendered by a headless test, but a Source still holding one is exactly the state that crashes.
    //
    // Fails against the version that disposed without clearing.
    // ###########################################################################################
    [Fact]
    public void Clearing_a_captured_image_does_not_leave_the_main_image_holding_the_disposed_bitmap()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            window.ShowTemporaryCapturedOscilloscopeImageForTests(this.WriteCaptureImage());
            Assert.NotNull(window.MainComponentImageSourceForTests);

            window.ClearTemporaryCapturedOscilloscopeImageForTests();

            Assert.Null(window.MainComponentImageSourceForTests);
        });
    }

    // The clear runs at the START of every show too (a second capture replaces the first), so it
    // must not sabotage the assignment that follows it - the new bitmap has to survive.
    [Fact]
    public void Showing_a_second_capture_leaves_the_new_bitmap_in_the_main_image()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            window.ShowTemporaryCapturedOscilloscopeImageForTests(this.WriteCaptureImage());
            var first = window.MainComponentImageSourceForTests;

            window.ShowTemporaryCapturedOscilloscopeImageForTests(this.WriteCaptureImage());
            var second = window.MainComponentImageSourceForTests;

            Assert.NotNull(second);
            Assert.NotSame(first, second);
        });
    }

    // A clear with no capture showing must leave whatever else the popup is displaying alone - the
    // release is targeted at the capture bitmap, not a blanket "empty the image".
    [Fact]
    public void Clearing_with_no_capture_showing_does_not_blank_the_main_image()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentInfoWindow();

            var unrelated = new Avalonia.Media.Imaging.RenderTargetBitmap(
                new Avalonia.PixelSize(8, 8), new Avalonia.Vector(96, 96));

            window.GetControl<Image>("MainComponentImage").Source = unrelated;

            window.ClearTemporaryCapturedOscilloscopeImageForTests();

            Assert.Same(unrelated, window.MainComponentImageSourceForTests);
        });
    }

    // ###########################################################################################
    // A pin clicked goes to its images (owner request, 2026-11-01: "so the user quickly can go to that pin
    // and see what images are available for that pin") - the first of them selected, the pin flashed over
    // it; one with none says so and leaves the image shown where it is.
    // ###########################################################################################
    [Fact]
    public void A_pin_goes_to_its_first_image_or_says_it_has_none()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            SetComponent(window, new[] { Component() });

            ListBox thumbnails = ControlOf<ListBox>(window, "ThumbnailList");
            thumbnails.ItemsSource = new List<ComponentImageItem> { new() { Pin = "1" }, new() { Pin = "7" }, new() { Pin = "7" }, new() { Pin = "9" } };

            window.GoToPin("7");

            Assert.Equal(1, thumbnails.SelectedIndex);
            Assert.True(ControlOf<Border>(window, "PinFlashBorder").IsVisible);
            Assert.Equal("7", ControlOf<TextBlock>(window, "PinFlashText").Text);

            window.GoToPin("3");

            Assert.Equal(1, thumbnails.SelectedIndex);
            Assert.Equal(ComponentPopupContent.NoImageOfPin("3"), ControlOf<TextBlock>(window, "PinFlashText").Text);
        });
    }

    // ###########################################################################################
    // A pin typed goes where the same pin clicked goes (code review, 2026-11-02): "2 [64K]" is pin 2 to
    // the drawing and the list, and typing 2 flashed "Not found", comparing the text exactly.
    // ###########################################################################################
    [Fact]
    public void A_pin_typed_finds_the_image_a_click_on_that_pin_finds()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            SetComponent(window, new[] { Component() });

            ListBox thumbnails = ControlOf<ListBox>(window, "ThumbnailList");
            thumbnails.ItemsSource = new List<ComponentImageItem> { new() { Pin = "1" }, new() { Pin = "2 [64K]" }, new() { Pin = "A3" } };

            window.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.D2 });

            Assert.Equal(1, thumbnails.SelectedIndex);
            Assert.Equal("2", ControlOf<TextBlock>(window, "PinFlashText").Text);
        });
    }

    // Beside the drawing while there is room, so both are seen at once; under it when the popup is narrow.
    [Fact]
    public void The_pin_list_stands_beside_the_drawing_and_under_it_in_a_narrow_popup()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(
                window,
                new[] { new ComponentEntry { BoardLabel = BoardLabel, Category = "IC", Package = "DIP-8" } },
                pins: [new ComponentPinEntry { BoardLabel = BoardLabel, Pin = "1", Name = "VCC", Direction = "Power" }]);

            Control list = ControlOf<Control>(window, "InfoPinsPanel");
            Grid grid = ControlOf<Grid>(window, "PinoutGrid");

            grid.RaiseEvent(new SizeChangedEventArgs(Control.SizeChangedEvent, grid, new Avalonia.Size(0, 0), new Avalonia.Size(600, 200)));

            Assert.Equal((1, 0), (Grid.GetColumn(list), Grid.GetRow(list)));

            grid.RaiseEvent(new SizeChangedEventArgs(Control.SizeChangedEvent, grid, new Avalonia.Size(600, 200), new Avalonia.Size(300, 200)));

            Assert.Equal((0, 1), (Grid.GetColumn(list), Grid.GetRow(list)));
            Assert.Equal(2, Grid.GetColumnSpan(list));
        });
    }

    // ###########################################################################################
    // A file or a link of one region shows in that region only, and follows the popup's own switch
    // (owner request, 2026-11-02).
    // ###########################################################################################
    [Fact]
    public void The_files_and_links_follow_the_region_buttons()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();

            SetComponent(
                window,
                new[] { Component() },
                localFiles: [new ComponentLocalFileEntry { BoardLabel = BoardLabel, Name = "PAL note", File = "a.pdf", Region = "PAL" }],
                links: [new ComponentLinkEntry { BoardLabel = BoardLabel, Name = "NTSC page", Url = "https://example.org", Region = "NTSC" }],
                region: "PAL",
                hasExplicitRegionComponents: true);

            Assert.True(ControlOf<Control>(window, "LocalFilesSection").IsVisible);
            Assert.False(ControlOf<Control>(window, "LinksSection").IsVisible);

            ControlOf<Button>(window, "NtscRegionButton").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            Assert.False(ControlOf<Control>(window, "LocalFilesSection").IsVisible);
            Assert.True(ControlOf<Control>(window, "LinksSection").IsVisible);
        });
    }

    // The headings say "Net name", in bold (owner request, 2026-11-02).
    [Fact]
    public void The_pin_headings_are_bold_and_the_net_heading_says_net_name()
    {
        UiTest.Run(() =>
        {
            var window = CreateWindow();
            TextBlock net = ControlOf<TextBlock>(window, "InfoPinsNetHeading");

            Assert.Equal("Net name", net.Text);
            Assert.Equal(Avalonia.Media.FontWeight.Bold, net.FontWeight);
        });
    }
}
