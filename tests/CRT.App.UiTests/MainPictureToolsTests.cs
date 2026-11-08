using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The picture tools as Main carries them between the tabs (Main.PictureTools.cs - CRT 4.0 step 2,
// slice 6): a tool asked for on the Contribute tab opens the Schematics tab - on the board read again
// first when asked, the tool started by THAT board's load - and its end goes back to the Contribute
// tab. "Edit component" on a component's popup opens the Contribute tab. Main's StartAsync is never
// called (MainWindowTests' header).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainPictureToolsTests : IDisposable
{
    private const string BoardKey = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();

    public MainPictureToolsTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static void Click(Control owner, string name) =>
        owner.GetControl<Button>(name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    // ###########################################################################################
    // The Schematics tab with the tool asked for - here refused, the board shown having no such
    // picture - and its end back on the Contribute tab.
    // ###########################################################################################
    [Fact]
    public void A_tool_from_the_Contribute_tab_opens_the_Schematics_tab_and_its_end_goes_back()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: false);

            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);
            TabSchematics tab = window.TabSchematicsControl;
            Assert.Equal("Components cannot be marked on [Board top]", tab.GetControl<TextBlock>("PictureToolTitleText").Text);

            Click(tab, "PictureToolCancelButton");

            Assert.Same(window.ContributeTabItem, window.MainTabControl.SelectedItem);
        });
    }

    // ###########################################################################################
    // THE DELETE KEY THE WAY A REAL PRESS TAKES IT (code review, 2026-11-04: LabelEditorInteractionTests call the tab's
    // handler straight, whatever holds the focus, so none could tell where the key really goes). With the picture
    // holding the focus, as a press on it leaves it, the rectangles selected go. In the component filter, in Main's
    // left panel, the tab is not on the key's route: it takes a letter out of the search, and the rectangles stay.
    // ###########################################################################################
    [Fact]
    public void The_Delete_key_is_the_pictures_while_it_has_the_focus_and_the_filters_while_that_has()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { Width = 1300, Height = 800 };
            TabSchematics tab = window.TabSchematicsControl;

            tab.CurrentBoardDataOverrideForTests = new BoardData
            {
                Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }, new ComponentEntry { BoardLabel = "U2", Category = "IC" }],
                ComponentHighlights =
                [
                    new ComponentHighlightEntry { SchematicName = "Board top", BoardLabel = "U1", X = "100", Y = "100", Width = "50", Height = "40" },
                    new ComponentHighlightEntry { SchematicName = "Board top", BoardLabel = "U2", X = "400", Y = "400", Width = "60", Height = "30" },
                ],
            };

            tab.SelectedThumbnail = new SchematicThumbnail { Name = "Board top" };

            window.Show();
            Dispatcher.UIThread.RunJobs();

            tab.BeginLabelEditorModeForTests();
            tab.SelectLabelEditorRowForTests(0);

            TextBox filter = window.ComponentSearchTextBox;

            filter.Text = "U1";
            filter.Focus();
            filter.CaretIndex = 0;
            Assert.True(filter.IsFocused);

            window.KeyPress(Key.Delete, RawInputModifiers.None, PhysicalKey.Delete, keySymbol: null);

            Assert.Equal(["U1", "U2"], tab.LabelEditorWorkingRowsForTests.Select(row => row.BoardLabel));
            Assert.Equal("1", filter.Text);

            tab.SchematicsContainer.Focus();
            Assert.True(tab.SchematicsContainer.IsFocused);

            window.KeyPress(Key.Delete, RawInputModifiers.None, PhysicalKey.Delete, keySymbol: null);

            Assert.Equal(["U2"], tab.LabelEditorWorkingRowsForTests.Select(row => row.BoardLabel));
            Assert.Equal("1", filter.Text);

            window.Close();
        });
    }

    // Read again first: started by that board's load, once - and never by another board's.
    [Fact]
    public void After_reading_the_board_again_the_tool_is_started_by_that_boards_load_only()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainPictureToolsTests.BoardKey };
            TabSchematics tab = window.TabSchematicsControl;

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: true);
            Dispatcher.UIThread.RunJobs();

            window.DoWhatWasAskedAfterLoad("Commodore 64|250425");
            Assert.False(tab.GetControl<Border>("PictureToolBarBorder").IsVisible);

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: true);
            Dispatcher.UIThread.RunJobs();
            window.DoWhatWasAskedAfterLoad(MainPictureToolsTests.BoardKey);

            Assert.True(tab.GetControl<Border>("PictureToolBarBorder").IsVisible);
        });
    }

    // ###########################################################################################
    // A picture tool on while the Contribute tab asks for the board to be read again (code review,
    // 2026-10-26): the reload reset the tab, which ended the label editor and threw its rectangles
    // away without a word. It waits for the tool to end now - and "Mark components" asked for again
    // shows the tool that is on rather than queueing another.
    // ###########################################################################################
    [Fact]
    public void A_reload_asked_for_while_a_tool_is_on_waits_until_the_tool_ends()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainPictureToolsTests.BoardKey };
            TabSchematics tab = window.TabSchematicsControl;
            tab.CurrentBoardDataOverrideForTests = new BoardData
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Board top" }],
                Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
                ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Board top", BoardLabel = "U1", X = "10", Y = "10", Width = "40", Height = "20" }],
            };

            var thumbnail = new Tabs.TabSchematics.SchematicThumbnail { Name = "Board top" };
            tab.currentThumbnails.Add(thumbnail);
            tab.SelectedThumbnail = thumbnail;
            tab.BeginLabelEditorModeForTests();
            Assert.Single(tab.LabelEditorWorkingRowsForTests);

            window.ReloadCurrentBoardFromDisk("Board top");
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.IsLabelEditorModeForTests);
            Assert.Single(tab.LabelEditorWorkingRowsForTests);
            Assert.True(window.IsReloadWaitingForToolForTests);

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: true);
            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);
            Assert.True(tab.IsLabelEditorModeForTests);

            // The tool ends: the board is read again then.
            tab.CancelLabelEditorChangesForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.IsReloadWaitingForToolForTests);
        });
    }

    // ###########################################################################################
    // "Done - next picture" (step 3, slice 3d) with a reload waiting for the tool: the board is read
    // ONCE, now, with the NEXT picture to come back to - and the next picture's alignment started by
    // that load. Run as the tool went off instead, the reload reset the tab and forgot the next
    // picture halfway, leaving the Schematics tab with no tool and no way back. CRT stays on the
    // Schematics tab throughout.
    // ###########################################################################################
    [Fact]
    public void Done_next_picture_reads_a_waiting_board_once_with_the_next_picture()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainPictureToolsTests.BoardKey };
            TabSchematics tab = window.TabSchematicsControl;

            foreach ((string name, string cadName) in new[] { ("Top", "Board - PCB Top"), ("Bottom", "Board - PCB Bottom") })
            {
                tab.schematicByName[name] = new BoardSchematicEntry { SchematicName = name, CadName = cadName };
                tab.currentThumbnails.Add(new Tabs.TabSchematics.SchematicThumbnail { Name = name });
            }

            // The top on screen, its image read (images are not decoded headlessly).
            tab.SelectedThumbnail = tab.currentThumbnails[0];
            tab.ImageSettledForTests("Top");
            tab.currentFullResBitmap = new Avalonia.Media.Imaging.WriteableBitmap(
                new Avalonia.PixelSize(200, 100), new Avalonia.Vector(96, 96), Avalonia.Platform.PixelFormat.Bgra8888, Avalonia.Platform.AlphaFormat.Premul);
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

            var saved = new List<string>();
            tab.KiCadCalibrationSaveOverrideForTests = row =>
            {
                saved.Add(row.SchematicName);
                return Task.FromResult<string?>(null);
            };
            tab.AlignedPicturesOverrideForTests = () => saved;

            window.StartPictureTool(PictureTool.AlignKiCad, "Top", readBoardFirst: false);
            tab.KiCadSettled(hasProject: true);
            Assert.True(tab.IsCalibrationModeForTests);

            // The Contribute tab's saves left the board behind: its reload waits for the tool.
            window.ReloadCurrentBoardFromDisk("Top");
            Assert.True(window.IsReloadWaitingForToolForTests);

            Click(tab, "PictureToolDoneNextButton");
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Top"], saved);
            Assert.False(window.IsReloadWaitingForToolForTests);
            Assert.Equal("Bottom", UserSettings.GetLastSchematicForBoard(MainPictureToolsTests.BoardKey));
            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);

            // What that load shows - a test's Main has no board to read, so its pictures are put back
            // as the load puts them, the one to come back to selected and read - and then its end:
            // aligning the bottom asked for, waiting for the KiCad data the load reads again.
            foreach (string name in new[] { "Top", "Bottom" }.Where(name => !tab.currentThumbnails.Any(thumbnail => thumbnail.Name == name)))
            {
                tab.currentThumbnails.Add(new Tabs.TabSchematics.SchematicThumbnail { Name = name });
            }

            tab.SelectSchematicByName(UserSettings.GetLastSchematicForBoard(MainPictureToolsTests.BoardKey)!);
            tab.ImageSettledForTests("Bottom");
            window.DoWhatWasAskedAfterLoad(MainPictureToolsTests.BoardKey);

            Assert.Equal("Bottom", tab.GetCurrentSchematicName());
            Assert.Equal("Opening [Bottom]...", tab.GetControl<TextBlock>("PictureToolTitleText").Text);
            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);
        });
    }

    // ###########################################################################################
    // *** EVERY PICTURE ALIGNED WITH DONE IS MARKED ONCE THE TOOL HANDS BACK (owner request,
    // 2026-11-01: "it is not clear that I must first do the Align KiCad, before I can do the Mark all
    // automatically from KiCad" - Done marks without asking). *** Done on the top: back on the
    // Contribute tab, it is marked - with its image file, never while the tool is on. A cancelled
    // alignment is not. (That "Done - next picture" tells each picture it saves is
    // TabSchematicsPictureToolsTests'; a test's Main cannot read the next picture's image.)
    // ###########################################################################################
    // ###########################################################################################
    // *** ANOTHER BOARD SHOWN BETWEEN TWO MARKS STOPS THEM (code review, 2026-11-01). *** Each picture
    // is marked on whatever board is shown THEN (TabDrafts.OnCurrentDraftAsync), so the board is asked
    // before each one - asked once before the first, a press on the board's drop-down in the moment the
    // overlay let go between two marks had the next picture looked for on the other board's draft.
    // ###########################################################################################
    [Fact]
    public async Task Another_board_shown_between_two_marks_stops_the_rest()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = "Commodore|C64|250407" };

            window.RememberAligned("Top", "top.png");
            window.RememberAligned("Bottom", "bottom.png");

            var marked = new List<string>();
            window.MarkAfterAligningOverrideForTests = (picture, _) =>
            {
                marked.Add(picture);
                window.BoardKeyOverrideForTests = "Commodore|C64|250425";
                return Task.CompletedTask;
            };

            await window.MarkPicturesAlignedAsync();

            Assert.Equal(["Top"], marked);
        });
    }

    [Fact]
    public async Task The_pictures_aligned_with_Done_are_marked_once_the_tool_hands_back()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainPictureToolsTests.BoardKey };
            TabSchematics tab = window.TabSchematicsControl;

            foreach ((string name, string cadName) in new[] { ("Top", "Board - PCB Top"), ("Bottom", "Board - PCB Bottom") })
            {
                tab.schematicByName[name] = new BoardSchematicEntry { SchematicName = name, SchematicImageFile = $"{name.ToLowerInvariant()}.png", CadName = cadName };
                tab.currentThumbnails.Add(new Tabs.TabSchematics.SchematicThumbnail { Name = name });
            }

            tab.SelectedThumbnail = tab.currentThumbnails[0];
            tab.ImageSettledForTests("Top");
            tab.currentFullResBitmap = new Avalonia.Media.Imaging.WriteableBitmap(
                new Avalonia.PixelSize(200, 100), new Avalonia.Vector(96, 96), Avalonia.Platform.PixelFormat.Bgra8888, Avalonia.Platform.AlphaFormat.Premul);
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

            var saved = new List<string>();
            tab.KiCadCalibrationSaveOverrideForTests = row =>
            {
                saved.Add(row.SchematicName);
                return Task.FromResult<string?>(null);
            };
            tab.AlignedPicturesOverrideForTests = () => saved;

            var marked = new List<string>();
            window.MarkAfterAligningOverrideForTests = (picture, imageFile) =>
            {
                Assert.False(tab.IsCalibrationModeForTests);
                Assert.Same(window.ContributeTabItem, window.MainTabControl.SelectedItem);
                marked.Add($"{picture} {imageFile}");
                return Task.CompletedTask;
            };

            window.StartPictureTool(PictureTool.AlignKiCad, "Top", readBoardFirst: false);
            tab.KiCadSettled(hasProject: true);
            Assert.True(tab.IsCalibrationModeForTests);

            Click(tab, "PictureToolDoneButton");
            Dispatcher.UIThread.RunJobs();
            await window.PicturesAlignedMarkedForTests;

            Assert.Equal(["Top"], saved);
            Assert.Equal(["Top top.png"], marked);

            // Aligned again and cancelled: nothing marked.
            marked.Clear();
            window.StartPictureTool(PictureTool.AlignKiCad, "Top", readBoardFirst: false);
            tab.currentFullResBitmap = new Avalonia.Media.Imaging.WriteableBitmap(
                new Avalonia.PixelSize(200, 100), new Avalonia.Vector(96, 96), Avalonia.Platform.PixelFormat.Bgra8888, Avalonia.Platform.AlphaFormat.Premul);
            tab.ImageSettledForTests("Top");
            Assert.True(tab.IsCalibrationModeForTests);
            Click(tab, "PictureToolCancelButton");
            Dispatcher.UIThread.RunJobs();
            await window.PicturesAlignedMarkedForTests;

            Assert.Empty(marked);
        });
    }

    // ###########################################################################################
    // The menu's first item: the component's highlight on, then off - what a right-click on a
    // component did by itself before the menu.
    // ###########################################################################################
    // "Edit component": the Contribute tab.
    [Fact]
    public async Task Edit_component_opens_the_Contribute_tab()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main();

            await window.EditComponentAsync("U8");

            Assert.Same(window.ContributeTabItem, window.MainTabControl.SelectedItem);
        });
    }

    // ###########################################################################################
    // "EDIT COMPONENT" ON THE POPUP (owner request, 2026-11-05: "maybe there could be a button in the component
    // info popup for 'Edit component' or alike"): the popup goes - it stands over CRT's window - and the
    // Contribute tab shows. A popup stays open as another board is shown, whose same label is another
    // component, so it no longer offers it. The Contribute tab's own preview never does: that is where the
    // component is being edited.
    // ###########################################################################################
    [Fact]
    public void Edit_component_on_the_popup_closes_it_and_opens_the_Contribute_tab()
    {
        UiTest.Run(() =>
        {
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };
            var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC" }] };
            var window = new CRT.Main { BoardKeyOverrideForTests = BoardKey };

            window.UseBoardForTests(entry, board);
            window.Show();
            Dispatcher.UIThread.RunJobs();

            window.OpenComponentInfoPopup("U8", "U8 | CPU");
            ComponentInfoWindow popup = window.SingleComponentInfoWindowForTests!;
            Button edit = popup.GetControl<Button>("EditComponentButton");

            Assert.True(edit.IsVisible);
            Assert.Equal(PictureToolWording.EditComponent, edit.Content);

            edit.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.False(popup.IsVisible);
            Assert.Same(window.ContributeTabItem, window.MainTabControl.SelectedItem);

            window.MainTabControl.SelectedItem = window.SchematicsTabItem;
            Dispatcher.UIThread.RunJobs();

            // Open as another board is chosen: not this board's component.
            window.OpenComponentInfoPopup("U8", "U8 | CPU");
            popup = window.SingleComponentInfoWindowForTests!;
            edit = popup.GetControl<Button>("EditComponentButton");
            Assert.True(edit.IsVisible);

            popup.FollowBoardShown("Commodore/C64/250425/board.json");
            Assert.False(edit.IsVisible);

            // No drop-down list is filled here, so the board is known to both as none.
            popup.FollowBoardShown(string.Empty);
            Assert.True(edit.IsVisible);

            // The Contribute tab's preview: what will be published, with nothing to go and edit.
            window.OpenComponentPreview(board, "U8", string.Empty, entry.ExcelDataFile);
            Assert.False(window.SingleComponentInfoWindowForTests!.GetControl<Button>("EditComponentButton").IsVisible);

            window.IsApplicationShuttingDownForTests = true;
            window.Close();
        });
    }
}
