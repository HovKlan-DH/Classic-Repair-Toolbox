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
