using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

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
    private const string DataRoot = @"C:\data";

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
                false);

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
        bool hasExplicitRegionComponents = false)
    {
        window.SetComponent(
            BoardLabel,
            displayText,
            entries.ToList(),
            new List<ComponentImageEntry>(),
            (localFiles ?? Array.Empty<ComponentLocalFileEntry>()).ToList(),
            (links ?? Array.Empty<ComponentLinkEntry>()).ToList(),
            region,
            DataRoot,
            hasExplicitRegionComponents);
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
}
