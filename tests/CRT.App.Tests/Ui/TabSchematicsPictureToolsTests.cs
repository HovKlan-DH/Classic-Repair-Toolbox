using Avalonia.Controls;
using Avalonia.Interactivity;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The picture tools and the component menu on the Schematics tab (TabSchematics.PictureTools.cs -
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

    private static string Text(TabSchematics tab, string name) => tab.GetControl<TextBlock>(name).Text ?? string.Empty;

    private static bool BarShown(TabSchematics tab) => tab.GetControl<Border>("PictureToolBarBorder").IsVisible;

    private static void Click(TabSchematics tab, string name) =>
        tab.GetControl<Button>(name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

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
            Assert.Equal("Mark components on [Board top]", Text(tab, "PictureToolTitleText"));
            Assert.Equal(PictureToolWording.Help(PictureTool.MarkComponents), Text(tab, "PictureToolHelpText"));
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
            Assert.Equal("KiCad cannot be aligned on [Board top]", Text(tab, "PictureToolTitleText"));
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
            Assert.Equal("Align KiCad on [Bottom]", Text(tab, "PictureToolTitleText"));

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
    // the next picture is not gone on to.
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

            Align(tab, "Top");
            Click(tab, "PictureToolDoneNextButton");

            Assert.True(tab.IsCalibrationModeForTests);
            Assert.Equal("Top", tab.GetCurrentSchematicName());
            Assert.Equal("Align KiCad on [Top]", Text(tab, "PictureToolTitleText"));
            Assert.Empty(asked);
            Assert.Empty(ended);
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
    // The component menu: its name, and "Edit this component" - handed to whoever opens it, the
    // menu closed.
    // ###########################################################################################
    [Fact]
    public void The_component_menu_names_the_component_and_hands_on_Edit_this_component()
    {
        UiTest.Run(() =>
        {
            (TabSchematics tab, _) = CreateTab();
            var asked = new List<string>();
            tab.EditComponentRequested += (_, label) => asked.Add(label);

            tab.ShowComponentMenuForTests(" U1 ");

            Assert.True(tab.GetControl<Border>("SchematicsMenuBorder").IsVisible);
            Assert.Equal("U1 - CPU", Text(tab, "SchematicsMenuTitleText"));
            Assert.Equal("Highlight [U1]", tab.GetControl<Button>("MenuHighlightButton").Content);
            Assert.Equal(PictureToolWording.EditComponent, tab.GetControl<Button>("MenuEditComponentButton").Content);

            Click(tab, "MenuEditComponentButton");

            Assert.Equal(["U1"], asked);
            Assert.False(tab.GetControl<Border>("SchematicsMenuBorder").IsVisible);
        });
    }
}
