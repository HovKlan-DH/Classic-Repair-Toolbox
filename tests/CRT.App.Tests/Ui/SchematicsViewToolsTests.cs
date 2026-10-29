using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Schematics view's small tools (TabSchematics.ViewTools.cs - CRT 4.0 step 5, the plan's
// "Smaller items"): a right OR middle drag pans, with a one-time tip saying so; traces are drawn
// and changed only with "Draw trace" on; "Global settings" sit behind a gear button. Driven through
// a shown Main, the mouse in its window - the trace tool is made by Main's Initialize. StartAsync
// is never called (MainWindowTests' header).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SchematicsViewToolsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SchematicsViewToolsTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
    }

    public void Dispose()
    {
        TraceStorage.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    // Main shown, a picture in the Schematics view - as the view has it once a picture loaded.
    private static (CRT.Main Main, TabSchematics Tab) ShowWithPicture()
    {
        var main = new CRT.Main { Width = 1280, Height = 800 };
        main.Show();
        Dispatcher.UIThread.RunJobs();

        TabSchematics tab = main.TabSchematicsControl;
        var bitmap = new WriteableBitmap(new PixelSize(1200, 900), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        tab.currentFullResBitmap = bitmap;
        tab.SchematicsImage.Source = bitmap;
        tab.ShowViewToolsForPicture(true);

        // The picture chosen, and its traces read - traces are changed only while they are the
        // chosen picture's (TabSchematics.ViewTools.cs). TraceStorage emptied first: a trace drawn
        // is saved in memory under the picture's name, and the next test read it back.
        TraceStorage.LoadFrom(string.Empty);
        main.BoardKeyOverrideForTests = "Commodore/C64/250407";
        tab.SuppressThumbnailSelectionChangedForTests = true;
        tab.SelectedThumbnail = new Tabs.TabSchematics.SchematicThumbnail { Name = "Picture" };
        tab.SuppressThumbnailSelectionChangedForTests = false;
        tab.ImportTracesForTests("Picture");

        main.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
        return (main, tab);
    }

    private static void Close(CRT.Main main)
    {
        main.IsApplicationShuttingDownForTests = true;
        main.Close();
    }

    // A point of the picture's container, in the window's coordinates.
    private static Point At(CRT.Main main, TabSchematics tab, double x, double y) =>
        tab.SchematicsContainer.TranslatePoint(new Point(x, y), main)!.Value;

    private static void Drag(CRT.Main main, Point from, Point to, MouseButton button)
    {
        main.MouseDown(from, button);
        main.MouseMove(new Point((from.X + to.X) / 2, (from.Y + to.Y) / 2));
        main.MouseMove(to);
        main.MouseUp(to, button);
        Dispatcher.UIThread.RunJobs();
    }

    private static void ZoomIn(TabSchematics tab)
    {
        MethodInfo zoom = typeof(TabSchematics).GetMethod("ApplySchematicsZoom", BindingFlags.NonPublic | BindingFlags.Instance)!;
        zoom.Invoke(tab, [3.0, new Point(150, 200)]);
    }

    // ###########################################################################################
    // The middle button pans as the right one does (the plan: "a middle-button drag pans too").
    // ###########################################################################################
    [Theory]
    [InlineData(MouseButton.Middle)]
    [InlineData(MouseButton.Right)]
    public void A_middle_or_right_drag_moves_the_picture(MouseButton button)
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                ZoomIn(tab);
                Matrix before = tab.schematicsMatrix;

                Drag(main, At(main, tab, 150, 200), At(main, tab, 190, 230), button);

                Assert.NotEqual(before.M31, tab.schematicsMatrix.M31);
                Assert.NotEqual(before.M32, tab.schematicsMatrix.M32);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // The tip shows with the first picture until it is put away - by "Got it", or by the first pan -
    // and then never again (the plan: "a one-time hint in the corner").
    // ###########################################################################################
    [Fact]
    public void The_tip_shows_with_the_first_picture_and_Got_it_puts_it_away_for_good()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                Assert.True(tab.PanHintBorder.IsVisible);
                Assert.Equal(TabSchematics.PanHint, tab.PanHintText.Text);

                tab.PanHintGotItButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

                Assert.False(tab.PanHintBorder.IsVisible);
                Assert.True(UserSettings.SchematicsPanHintSeen);

                tab.ShowViewToolsForPicture(true);
                Assert.False(tab.PanHintBorder.IsVisible);
            }
            finally
            {
                Close(main);
            }
        });
    }

    [Fact]
    public void The_first_pan_puts_the_tip_away()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                Drag(main, At(main, tab, 150, 300), At(main, tab, 160, 310), MouseButton.Middle);

                Assert.False(tab.PanHintBorder.IsVisible);
                Assert.True(UserSettings.SchematicsPanHintSeen);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // A drag on the picture draws a trace ONLY with "Draw trace" on (the plan: "Traces are drawn
    // only with Draw trace") - off, a press there never starts one.
    // ###########################################################################################
    [Fact]
    public void Only_with_Draw_trace_on_does_a_drag_draw_a_trace()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                Assert.True(tab.DrawTraceButton.IsVisible);

                Drag(main, At(main, tab, 100, 300), At(main, tab, 250, 300), MouseButton.Left);
                Assert.Equal(0, tab.polylineManager!.TraceCountForTests);

                tab.DrawTraceButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                Assert.True(tab.IsDrawingTraces);
                Assert.Equal("Done drawing", tab.DrawTraceButton.Content);

                Drag(main, At(main, tab, 100, 300), At(main, tab, 250, 300), MouseButton.Left);
                Assert.Equal(1, tab.polylineManager.TraceCountForTests);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // With "Draw trace" off, a right-click on a trace leaves it - the right button pans there as
    // anywhere. On, it removes it, as it did.
    // ###########################################################################################
    [Fact]
    public void Without_Draw_trace_a_right_click_on_a_trace_leaves_it()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.SetDrawingTraces(true);
                Drag(main, At(main, tab, 100, 300), At(main, tab, 250, 300), MouseButton.Left);
                tab.SetDrawingTraces(false);
                Assert.Equal(1, tab.polylineManager!.TraceCountForTests);

                Point onTrace = At(main, tab, 175, 300);
                main.MouseDown(onTrace, MouseButton.Right);
                main.MouseUp(onTrace, MouseButton.Right);
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(1, tab.polylineManager.TraceCountForTests);

                tab.SetDrawingTraces(true);
                main.MouseDown(onTrace, MouseButton.Right);
                main.MouseUp(onTrace, MouseButton.Right);
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(0, tab.polylineManager.TraceCountForTests);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // Escape stops drawing wherever the keyboard is - the window's own Escape, since the component
    // filter usually holds the focus - and so does another board.
    // ###########################################################################################
    [Fact]
    public void Escape_or_another_board_stops_drawing()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.SetDrawingTraces(true);
                main.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape, Source = main });

                Assert.False(tab.IsDrawingTraces);
                Assert.Equal("Draw trace", tab.DrawTraceButton.Content);

                tab.SetDrawingTraces(true);
                tab.ResetSchematicsViewer();

                Assert.False(tab.IsDrawingTraces);
                Assert.False(tab.DrawTraceButton.IsVisible);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // *** ESCAPE IS "DRAW TRACE"'S ONLY ON ITS OWN TAB (code review, 2026-10-29). *** The window
    // takes Escape on the tunnel route; with drawing left on after going to another tab, Escape meant
    // for a box there (a picture's name put back, a table cell) stopped drawing instead, and the
    // box never saw it. Leaving the tab stops drawing, and Escape elsewhere is never taken for it.
    // ###########################################################################################
    [Fact]
    public void Leaving_the_Schematics_tab_stops_drawing_and_Escape_elsewhere_is_not_taken()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.SetDrawingTraces(true);
                main.MainTabControl.SelectedItem = main.ContributeTabItem;
                Dispatcher.UIThread.RunJobs();

                Assert.False(tab.IsDrawingTraces);

                // Drawing on with another tab shown (whatever turned it on): Escape there is not its.
                tab.SetDrawingTraces(true);
                var escape = new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape, Source = main };
                main.RaiseEvent(escape);

                Assert.False(escape.Handled);
                Assert.True(tab.IsDrawingTraces);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // *** NO PICTURE, NO DRAWING (code review, 2026-10-29). *** A picture that could not be shown hid
    // the button but left drawing on, so a drag still changed the previous picture's traces - saved
    // under the new picture's name.
    // ###########################################################################################
    [Fact]
    public void A_picture_that_cannot_be_shown_stops_drawing()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.SetDrawingTraces(true);

                tab.ShowViewToolsForPicture(false);

                Assert.False(tab.IsDrawingTraces);
                Assert.False(tab.DrawTraceButton.IsVisible);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // *** ANOTHER BOARD DROPS A DRAG, NEVER SAVES IT (code review, 2026-10-29). *** Turning drawing
    // off on a board change ended a point being dragged as a change - and the old board's traces
    // were saved under the new board's key, which the trace layer's own Reset was written to avoid.
    // ###########################################################################################
    [Fact]
    public void Another_board_drops_a_point_being_dragged_without_saving_it()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.SetDrawingTraces(true);
                Drag(main, At(main, tab, 100, 300), At(main, tab, 250, 300), MouseButton.Left);

                int saves = 0;
                tab.polylineManager!.TracesModified += () => saves++;

                // A point of the trace taken and moved, the button still held.
                main.MouseDown(At(main, tab, 250, 300), MouseButton.Left);
                main.MouseMove(At(main, tab, 260, 320));
                Dispatcher.UIThread.RunJobs();

                tab.ResetSchematicsViewer();

                Assert.Equal(0, saves);
                Assert.False(tab.IsDrawingTraces);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // *** ANOTHER PICTURE'S TRACES ARE NEVER SAVED UNDER THE PICTURE CHOSEN (code review,
    // 2026-10-29). *** While the next picture decoded, the last one's traces stayed loaded and
    // editable, and a point dragged then - or a colour ticked on a picture with no file - saved them
    // under the new picture's name, over its own traces. A picture's traces are taken in with it now,
    // and saved only under the picture they were read for.
    // ###########################################################################################
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Another_picture_s_traces_are_never_saved_under_the_picture_chosen(bool nextHasAFile)
    {
        UiTest.Run(() =>
        {
            const string board = "Commodore/C64/250407";
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();
            main.BoardKeyOverrideForTests = board;

            TraceStorage.LoadFrom(this.thisWorkspace.Path_("traces.json"));
            TraceStorage.SaveTraces(board, "B", [new TraceModel { Color = "#FF0000", Nodes = [new() { X = 1, Y = 2 }, new() { X = 3, Y = 4 }] }]);
            string bStored = System.Text.Json.JsonSerializer.Serialize(TraceStorage.GetTraces(board, "B"));

            try
            {
                var a = new Tabs.TabSchematics.SchematicThumbnail { Name = "A" };
                var b = new Tabs.TabSchematics.SchematicThumbnail
                {
                    Name = "B",
                    ImageFilePath = nextHasAFile ? this.thisWorkspace.WriteFile("b.png", "decoded later") : string.Empty,
                };

                tab.SuppressThumbnailSelectionChangedForTests = true;
                tab.SelectedThumbnail = a;
                tab.SuppressThumbnailSelectionChangedForTests = false;
                tab.ImportTracesForTests("A");

                // A trace drawn on A, and its end point taken, the button still held.
                tab.SetDrawingTraces(true);
                Drag(main, At(main, tab, 100, 300), At(main, tab, 250, 300), MouseButton.Left);
                Assert.Single(TraceStorage.GetTraces(board, "A"));

                main.MouseDown(At(main, tab, 250, 300), MouseButton.Left);
                main.MouseMove(At(main, tab, 260, 320));

                tab.SelectedThumbnail = b;

                // At once - before the next picture (if any) is read: A's traces are let go.
                Assert.Equal(0, tab.polylineManager!.TraceCountForTests);

                main.MouseMove(At(main, tab, 270, 340));
                main.MouseUp(At(main, tab, 270, 340), MouseButton.Left);

                // A picture with no file never reads traces: a colour ticked then saves nothing.
                if (!nextHasAFile)
                    tab.polylineManager.SetVisibilityByColor(Avalonia.Media.Colors.Red, false);

                Assert.Equal(bStored, System.Text.Json.JsonSerializer.Serialize(TraceStorage.GetTraces(board, "B")));
            }
            finally
            {
                Close(main);
                TraceStorage.LoadFrom(string.Empty);
            }
        });
    }

    // ###########################################################################################
    // *** UNDER ANOTHER TOOL "DRAW TRACE" IS GREYED OUT, SAYING WHY (code review, 2026-10-29). *** It
    // stayed clickable while marking a worklog area, the label editor or aligning KiCad, and a press
    // did nothing at all - no change, no reason.
    // ###########################################################################################
    [Fact]
    public void Draw_trace_is_greyed_out_with_its_reason_while_another_tool_is_on()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                Assert.True(tab.DrawTraceButton.IsEnabled);

                Assert.True(tab.BeginWorklogEntryMode(1));
                Assert.False(tab.DrawTraceButton.IsEnabled);
                Assert.Equal(TabSchematics.DrawTraceNotNow, ToolTip.GetTip(tab.DrawTraceButton));

                tab.CancelWorklogEntryMode();
                Assert.True(tab.DrawTraceButton.IsEnabled);
                Assert.Null(ToolTip.GetTip(tab.DrawTraceButton));
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // *** A PRESS ON THE GREYED-OUT "DRAW TRACE" IS NOT THE PICTURE'S (code review, 2026-10-29). *** A
    // disabled button takes no pointer, so the press went through to the picture under it - in the
    // label editor, where any left press lets go of the rectangles selected, the selection was gone.
    // ###########################################################################################
    [Fact]
    public void A_press_on_the_greyed_out_Draw_trace_leaves_the_label_editor_s_selection()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.CurrentBoardDataOverrideForTests = new BoardData
                {
                    Schematics = [new BoardSchematicEntry { SchematicName = "Picture" }],
                    Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
                    ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Picture", BoardLabel = "U1", X = "10", Y = "10", Width = "40", Height = "20" }],
                };

                tab.BeginLabelEditorModeForTests();
                tab.SelectLabelEditorRowForTests(0);
                Assert.Equal(1, tab.SelectedLabelEditorCountForTests);
                Assert.False(tab.DrawTraceButton.IsEnabled);

                Point button = tab.DrawTraceButton.TranslatePoint(
                    new Point(tab.DrawTraceButton.Bounds.Width / 2, tab.DrawTraceButton.Bounds.Height / 2), main)!.Value;

                main.MouseDown(button, MouseButton.Left);
                main.MouseUp(button, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(1, tab.SelectedLabelEditorCountForTests);
            }
            finally
            {
                tab.CancelLabelEditorChangesForTests();
                Close(main);
            }
        });
    }

    // "Done drawing" is plain: it turns a mode off and saves nothing, so it is no green button (code
    // review, 2026-10-29 - CLAUDE.md: Accept is for what goes ahead, saves, adds or sends).
    [Fact]
    public void Done_drawing_is_a_plain_button()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.SetDrawingTraces(true);

                Assert.Equal("Done drawing", tab.DrawTraceButton.Content);
                Assert.DoesNotContain("Accept", tab.DrawTraceButton.Classes);
                Assert.DoesNotContain("Danger", tab.DrawTraceButton.Classes);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // With no picture chosen no traces are read - and none are drawn, even with the bitmap of one still
    // there (code review, 2026-10-29: "none chosen" and "none read" counted as the same picture).
    [Fact]
    public void With_no_picture_chosen_no_trace_is_drawn()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                tab.SuppressThumbnailSelectionChangedForTests = true;
                tab.SelectedThumbnail = null;
                tab.SuppressThumbnailSelectionChangedForTests = false;
                tab.ResetTracesForTests();

                tab.SetDrawingTraces(true);
                Drag(main, At(main, tab, 100, 300), At(main, tab, 250, 300), MouseButton.Left);

                Assert.Equal(0, tab.polylineManager!.TraceCountForTests);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // "Draw trace" cannot be on without a picture.
    [Fact]
    public void Drawing_needs_a_picture()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            tab.SetDrawingTraces(true);

            Assert.False(tab.IsDrawingTraces);
        });
    }

    // ###########################################################################################
    // The gear shows only while its popover has a setting to show: every row is a KiCad one, so a
    // picture of a board without KiCad data has none - an empty popover is no setting.
    // ###########################################################################################
    [Fact]
    public void The_gear_is_hidden_while_no_setting_applies()
    {
        UiTest.Run(() =>
        {
            (CRT.Main main, TabSchematics tab) = ShowWithPicture();

            try
            {
                Assert.All(tab.GlobalSettingsListPanel.Children, row => Assert.False(row.IsVisible));
                Assert.False(tab.GlobalSettingsButton.IsVisible);
            }
            finally
            {
                Close(main);
            }
        });
    }

    // ###########################################################################################
    // "Global settings" are behind a gear button: its popover holds the rows the panel held (the
    // plan: "Global settings become a gear button that opens a popover").
    // ###########################################################################################
    [Fact]
    public void Global_settings_are_in_the_gear_buttons_popover()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            Flyout popover = Assert.IsType<Flyout>(tab.GlobalSettingsButton.Flyout);
            StackPanel content = Assert.IsType<StackPanel>(popover.Content);

            Assert.Contains(tab.GlobalSettingsListPanel, content.Children);
            Assert.Contains(tab.GlobalHoverHighlightsTracesRow, tab.GlobalSettingsListPanel.Children);
            Assert.Equal("Global settings", ToolTip.GetTip(tab.GlobalSettingsButton));
        });
    }
}
