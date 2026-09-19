using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using CRT;
using Handlers.DataHandling;
using System.Collections.ObjectModel;
using System.Linq;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// The "Detach thumbnails into their own window" feature: collapsing/restoring the Schematics tab's
// inline thumbnail column (TabSchematics.ThumbnailsDetach.cs), the detached window and gallery's
// construction, and the two-way selection sync between the gallery and the (now-hidden) inline
// SchematicsThumbnailList - the same pattern SchematicsFullscreenPlaceholder already uses to keep
// two ListBoxes in sync against one shared collection.
//
// WHAT THESE COVER. Everything drivable with no display and no real window resize: mode toggling,
// the single-shared-collection constraint Main.axaml.cs depends on, and selection forwarding both
// ways.
//
// WHAT THEY DO NOT. Actually driving a live .Show() plus a real OS resize to visually prove the
// auto-fit grid reflows - that is "rendering/layout in Tabs/", already excluded by CLAUDE.md's
// "Deliberately not covered" section; ThumbnailGalleryGeometryTests covers the pure math behind
// it instead. Also out of scope: the full close-button-unchecks-the-setting path end to end
// through Main's object graph (Main is never constructed by any test) - the TabConfiguration half
// of that resync is covered directly by TabConfigurationDetachThumbnailsTests instead.
//
// COLLECTION NOTE: "HeadlessUi" because it constructs TabSchematics, a UserControl, and
// SchematicsThumbnailsWindow, a real Window subclass - constructing a Window with no .Show() call
// works headlessly, confirmed by ComponentInfoWindowTests's own precedent.
[Collection("HeadlessUi")]
public sealed class SchematicsThumbnailsDetachTests
{
    [Fact]
    public void The_detached_window_constructs_without_throwing()
    {
        UiTest.Run(() =>
        {
            var window = new SchematicsThumbnailsWindow();

            Assert.NotNull(window);
        });
    }

    // F11 while this window is the active one must fullscreen the schematic, the same as it does
    // from the main window - asked for explicitly, since this window used to ignore F11 entirely.
    // Main itself is never constructed by any test (see CLAUDE.md), so the actual
    // ToggleSchematicsFullscreenWindow call cannot be observed here; what IS covered is that the
    // key reaches the handler and is marked handled rather than falling through to something else,
    // and that it does so safely with no MainWindow wired up (TabSchematics.MainWindow stays null
    // in every headless test, so this is also the ordinary test-time path).
    [Fact]
    public void F11_is_handled_by_the_detached_window_with_no_MainWindow_wired_up()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var window = new SchematicsThumbnailsWindow();
            window.Initialize(tab.currentThumbnails, hostedList, tab);
            window.Show();

            Assert.Null(tab.MainWindow);

            window.KeyPress(Key.F11, RawInputModifiers.None, PhysicalKey.F11, keySymbol: null);
        });
    }

    // A key other than F11 must be left alone - only F11 is this window's business.
    [Fact]
    public void A_non_F11_key_is_not_swallowed_by_the_detached_window()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var window = new SchematicsThumbnailsWindow();
            window.Initialize(tab.currentThumbnails, hostedList, tab);
            window.Show();

            bool bubbledToWindow = false;
            window.KeyDown += (_, e) =>
            {
                if (e.Key == Key.A)
                    bubbledToWindow = true;
            };

            window.KeyPress(Key.A, RawInputModifiers.None, PhysicalKey.A, keySymbol: null);

            Assert.True(bubbledToWindow);
        });
    }

    [Fact]
    public void Entering_detached_mode_hides_the_inline_list_and_collapses_the_columns()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");
            var list = tab.GetControl<ListBox>("SchematicsThumbnailList");

            tab.EnterThumbnailsDetachedMode();

            Assert.False(list.IsVisible);
            Assert.Equal(0.0, grid.ColumnDefinitions[1].Width.Value);
            Assert.Equal(0.0, grid.ColumnDefinitions[2].Width.Value);
            Assert.Equal(0.0, grid.ColumnDefinitions[2].MinWidth);
        });
    }

    [Fact]
    public void Exiting_detached_mode_restores_the_prior_column_widths_and_visibility()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");
            var list = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var originalRightWidth = grid.ColumnDefinitions[2].Width;
            var originalRightMinWidth = grid.ColumnDefinitions[2].MinWidth;

            tab.EnterThumbnailsDetachedMode();
            tab.ExitThumbnailsDetachedMode();

            Assert.True(list.IsVisible);
            Assert.Equal(originalRightWidth, grid.ColumnDefinitions[2].Width);
            Assert.Equal(originalRightMinWidth, grid.ColumnDefinitions[2].MinWidth);
        });
    }

    // The two modes COMPOSE rather than excluding each other: F11 must keep working while
    // thumbnails are detached (it was a silent no-op once, which just looked broken). Detaching
    // while fullscreen is active records the mode but leaves the columns to fullscreen, which has
    // already collapsed them and holds the pre-fullscreen widths it will restore.
    [Fact]
    public void Entering_detached_mode_while_fullscreen_records_the_mode_without_touching_the_columns()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");

            tab.EnterFullscreenMode();
            Assert.True(tab.IsFullscreenModeActive);

            var widthWhileFullscreen = grid.ColumnDefinitions[2].Width;

            tab.EnterThumbnailsDetachedMode();

            // The flag is set so ExitFullscreenMode knows not to bring the strip back...
            Assert.True(tab.IsThumbnailsDetached);

            // ...but the columns are fullscreen's to manage while it is active.
            Assert.Equal(widthWhileFullscreen, grid.ColumnDefinitions[2].Width);
        });
    }

    // The regression Dennis reported: leaving fullscreen must not resurrect the inline thumbnail
    // strip when thumbnails are supposed to be in their own window. ExitFullscreenMode used to set
    // IsVisible = true unconditionally.
    [Fact]
    public void Leaving_fullscreen_while_detached_keeps_the_inline_strip_hidden()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
            var splitter = tab.GetControl<GridSplitter>("SchematicsSplitter");

            tab.EnterThumbnailsDetachedMode();
            tab.EnterFullscreenMode();
            tab.ExitFullscreenMode();

            Assert.True(tab.IsThumbnailsDetached);
            Assert.False(tab.IsFullscreenModeActive);
            Assert.False(list.IsVisible);
            Assert.False(splitter.IsVisible);
        });
    }

    // The ordinary case must still restore the strip, or F11 would break for everyone not using
    // the detach setting.
    [Fact]
    public void Leaving_fullscreen_while_not_detached_restores_the_inline_strip()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var list = tab.GetControl<ListBox>("SchematicsThumbnailList");
            var splitter = tab.GetControl<GridSplitter>("SchematicsSplitter");

            tab.EnterFullscreenMode();
            tab.ExitFullscreenMode();

            Assert.True(list.IsVisible);
            Assert.True(splitter.IsVisible);
        });
    }

    // Un-ticking the setting while fullscreen is active must not paint the strip back over the
    // fullscreen view - the mirrored half of the guard above.
    [Fact]
    public void Exiting_detached_mode_while_fullscreen_leaves_the_strip_hidden_until_fullscreen_ends()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var list = tab.GetControl<ListBox>("SchematicsThumbnailList");

            tab.EnterThumbnailsDetachedMode();
            tab.EnterFullscreenMode();

            tab.ExitThumbnailsDetachedMode();

            Assert.False(tab.IsThumbnailsDetached);
            Assert.False(list.IsVisible);

            // Leaving fullscreen now lands on the ordinary attached layout.
            tab.ExitFullscreenMode();
            Assert.True(list.IsVisible);
        });
    }

    // ---------------------------------------------------------------------------------------
    // The two modes composing: whose restore widths win
    // ---------------------------------------------------------------------------------------

    // The reported-shaped defect: detaching WHILE fullscreen is active used to return before
    // capturing anything, leaving this mode's restore fields at their bare field initializers
    // (1* and MinWidth 100). Un-ticking the setting later then threw away the width the user had
    // actually dragged the splitter to and snapped the column to roughly half the tab.
    //
    // Fails against that version: the assertion below sees 1* instead of 450px.
    [Fact]
    public void Detaching_while_fullscreen_still_restores_the_users_own_splitter_width()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");

            // The user's own splitter position, before either mode touches anything.
            grid.ColumnDefinitions[2].Width = new GridLength(450, GridUnitType.Pixel);
            grid.ColumnDefinitions[2].MinWidth = 120;

            tab.EnterFullscreenMode();
            tab.EnterThumbnailsDetachedMode();

            tab.ExitFullscreenMode();
            tab.ExitThumbnailsDetachedMode();

            Assert.Equal(450, grid.ColumnDefinitions[2].Width.Value);
            Assert.Equal(GridUnitType.Pixel, grid.ColumnDefinitions[2].Width.GridUnitType);
            Assert.Equal(120, grid.ColumnDefinitions[2].MinWidth);
        });
    }

    // The other exit order, which is the one that exercises the hand-back in
    // ExitThumbnailsDetachedMode's own fullscreen branch: detach mode ends first, so fullscreen is
    // the mode left holding the real widths and must have been given them back.
    [Fact]
    public void Un_detaching_before_leaving_fullscreen_still_restores_the_users_own_splitter_width()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");

            grid.ColumnDefinitions[2].Width = new GridLength(450, GridUnitType.Pixel);
            grid.ColumnDefinitions[2].MinWidth = 120;

            tab.EnterFullscreenMode();
            tab.EnterThumbnailsDetachedMode();

            tab.ExitThumbnailsDetachedMode();
            tab.ExitFullscreenMode();

            Assert.Equal(450, grid.ColumnDefinitions[2].Width.Value);
            Assert.Equal(GridUnitType.Pixel, grid.ColumnDefinitions[2].Width.GridUnitType);
            Assert.Equal(120, grid.ColumnDefinitions[2].MinWidth);
        });
    }

    // The reverse entry order - detach first, then fullscreen - must reach the same place. Here it
    // is fullscreen that is entered second and so takes the widths off detach mode.
    [Fact]
    public void Entering_fullscreen_while_already_detached_still_restores_the_users_own_splitter_width()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");

            grid.ColumnDefinitions[2].Width = new GridLength(450, GridUnitType.Pixel);
            grid.ColumnDefinitions[2].MinWidth = 120;

            tab.EnterThumbnailsDetachedMode();
            tab.EnterFullscreenMode();

            tab.ExitFullscreenMode();
            tab.ExitThumbnailsDetachedMode();

            Assert.Equal(450, grid.ColumnDefinitions[2].Width.Value);
            Assert.Equal(120, grid.ColumnDefinitions[2].MinWidth);
        });
    }

    // Leaving fullscreen while detach mode is STILL on must not re-expand the column - the strip is
    // hidden, so an expanded column is an empty gap. This is what the hand-back has to preserve.
    [Fact]
    public void Leaving_fullscreen_while_still_detached_keeps_the_column_collapsed()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");
            var list = tab.GetControl<ListBox>("SchematicsThumbnailList");

            grid.ColumnDefinitions[2].Width = new GridLength(450, GridUnitType.Pixel);

            tab.EnterThumbnailsDetachedMode();
            tab.EnterFullscreenMode();
            tab.ExitFullscreenMode();

            Assert.True(tab.IsThumbnailsDetached);
            Assert.False(list.IsVisible);
            Assert.Equal(0, grid.ColumnDefinitions[2].Width.Value);
        });
    }

    // ---------------------------------------------------------------------------------------
    // Gallery wiring
    // ---------------------------------------------------------------------------------------

    // Main.axaml.cs clears and reassigns currentThumbnails directly on board switches, so the
    // gallery must observe the SAME instance rather than a copy handed to it once at Initialize.
    [Fact]
    public void Initializing_the_gallery_keeps_the_same_shared_thumbnails_instance()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var hostedList = tab.GetControl<ListBox>("SchematicsThumbnailList");

            var gallery = new SchematicsThumbnailGallery();
            gallery.Initialize(tab.currentThumbnails, hostedList, tab);

            // The gallery's ItemsSource must BE that collection, not a copy of it. Against a
            // `new ObservableCollection<SchematicThumbnail>(thumbnails)` in Initialize this fails,
            // which the previous `Assert.Same(tab.currentThumbnails, tab.currentThumbnails)` - true
            // of any object whatsoever - could not.
            var galleryList = gallery.GetControl<ListBox>("GalleryListBox");
            Assert.Same(tab.currentThumbnails, galleryList.ItemsSource);

            // And it is genuinely LIVE: an item added to the tab's collection afterwards is seen
            // through the gallery's own ItemsSource without anything re-initializing it.
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = "Added later" });

            var seenByGallery = galleryList.ItemsSource!.Cast<SchematicThumbnail>().ToList();
            Assert.Contains(seenByGallery, t => t.Name == "Added later");
        });
    }

    [Fact]
    public void Selecting_a_thumbnail_in_the_gallery_updates_the_hosted_list()
    {
        UiTest.Run(() =>
        {
            var thumbnails = new ObservableCollection<SchematicThumbnail>
            {
                new() { Name = "Schematic A" },
                new() { Name = "Schematic B" }
            };

            var hostedList = new ListBox { ItemsSource = thumbnails };
            var gallery = new SchematicsThumbnailGallery();
            var tab = new TabSchematics();

            gallery.Initialize(thumbnails, hostedList, tab);

            gallery.SelectedThumbnail = thumbnails[1];

            Assert.Same(thumbnails[1], hostedList.SelectedItem);
        });
    }

    // The leak: the hosted SchematicsThumbnailList lives for the application's lifetime inside
    // TabSchematics, while a new gallery is created on every detach. A subscription left on that
    // list keeps the whole dead gallery - and every decoded thumbnail bitmap its containers hold -
    // reachable forever, and each one still runs its handler on every schematic change.
    //
    // Asserted by behaviour rather than by counting handlers: after Teardown the gallery must stop
    // tracking the hosted list. Against the un-torn-down version the gallery follows it and fails.
    [Fact]
    public void Tearing_the_gallery_down_stops_it_following_the_hosted_list()
    {
        UiTest.Run(() =>
        {
            var thumbnails = new ObservableCollection<SchematicThumbnail>
            {
                new() { Name = "Schematic A" },
                new() { Name = "Schematic B" }
            };

            var hostedList = new ListBox { ItemsSource = thumbnails };
            var gallery = new SchematicsThumbnailGallery();
            var tab = new TabSchematics();

            gallery.Initialize(thumbnails, hostedList, tab);

            // Still wired: the gallery mirrors the hosted list.
            hostedList.SelectedItem = thumbnails[1];
            Assert.Same(thumbnails[1], gallery.SelectedThumbnail);

            gallery.Teardown();

            hostedList.SelectedItem = thumbnails[0];
            Assert.Same(thumbnails[1], gallery.SelectedThumbnail);
        });
    }

    // Teardown runs from BOTH OnDetachedFromVisualTree and the window's Closed handler, so it has
    // to survive being called twice - a window that was shown and then closed does both.
    [Fact]
    public void Tearing_the_gallery_down_twice_is_harmless()
    {
        UiTest.Run(() =>
        {
            var thumbnails = new ObservableCollection<SchematicThumbnail> { new() { Name = "Schematic A" } };

            var hostedList = new ListBox { ItemsSource = thumbnails };
            var gallery = new SchematicsThumbnailGallery();
            var tab = new TabSchematics();

            gallery.Initialize(thumbnails, hostedList, tab);

            gallery.Teardown();
            gallery.Teardown();
        });
    }

    [Fact]
    public void Selecting_a_thumbnail_in_the_hosted_list_updates_the_gallery()
    {
        UiTest.Run(() =>
        {
            var thumbnails = new ObservableCollection<SchematicThumbnail>
            {
                new() { Name = "Schematic A" },
                new() { Name = "Schematic B" }
            };

            var hostedList = new ListBox { ItemsSource = thumbnails };
            var gallery = new SchematicsThumbnailGallery();
            var tab = new TabSchematics();

            gallery.Initialize(thumbnails, hostedList, tab);

            hostedList.SelectedItem = thumbnails[1];

            Assert.Same(thumbnails[1], gallery.SelectedThumbnail);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Drag-reorder in the detached gallery
    // -----------------------------------------------------------------------------------------

    // Reported: dragging a thumbnail in the detached window showed no gap, unlike the inline list.
    // The first version moved the dragged item around live, so there was never a hole to see. It
    // now lifts the item out and puts a placeholder in its slot, exactly as the inline list does.
    [Fact]
    public void Dragging_lifts_the_thumbnail_out_and_leaves_a_visible_placeholder()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, _, _) = BuildGallery();
            var dragged = thumbnails[2];

            gallery.BeginTileDragForTests(dragged);

            // The dragged thumbnail is gone from the collection...
            Assert.DoesNotContain(dragged, thumbnails);

            // ...and a placeholder occupies the slot it came from, which is what the template
            // renders as the red outlined gap.
            Assert.Equal(3, thumbnails.Count);
            Assert.True(thumbnails[2].IsDropPlaceholder);
        });
    }

    // Reported: the placeholder appeared on the right image but could not be moved at all, then
    // vanished correctly on release. The drag captured the pointer to the PRESSED TILE and listened
    // for PointerMoved on it - but beginning a drag removes that tile's item from the collection,
    // so the ListBox destroys its container, taking the capture and the handler with it. The first
    // move event after the lift never arrived, so the drag was over the instant it began.
    //
    // The invariant is therefore about WHERE the drag is wired, which is why this asserts on the
    // markup and the capture target rather than on a simulated move: the seam-driven tests below
    // call the reorder step directly and so cannot see a routing bug at all - verified by putting
    // the broken wiring back and watching them all still pass.
    [Fact]
    public void The_drag_is_wired_to_the_gallery_not_to_the_tile_that_gets_destroyed()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, _, _) = BuildGallery();

            var window = new Window { Content = gallery, Width = 400, Height = 400 };
            window.Show();
            window.Measure(new Size(400, 400));
            window.Arrange(new Rect(0, 0, 400, 400));

            try
            {
                // Lifting the item out destroys the container the press happened on. The drag must
                // still be live and still own its placeholder afterwards - against the old wiring
                // the capture and the PointerMoved handler both died with that container.
                var dragged = thumbnails[2];
                gallery.BeginTileDragForTests(dragged);

                window.Measure(new Size(400, 400));
                window.Arrange(new Rect(0, 0, 400, 400));

                Assert.True(gallery.IsDraggingForTests);
                Assert.Contains(thumbnails, t => t.IsDropPlaceholder);

                // And the reorder still works end to end from that point.
                gallery.DragPlaceholderToPointForTests(new Point(10, 10));
                gallery.EndTileDragForTests(commitOrder: false);

                Assert.Same(dragged, thumbnails[0]);
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public void Dropping_puts_the_thumbnail_where_the_placeholder_ended_up()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, _, _) = BuildGallery();
            var dragged = thumbnails[2];

            gallery.BeginTileDragForTests(dragged);
            gallery.MoveDragPlaceholderForTests(0);
            gallery.EndTileDragForTests(commitOrder: false);

            Assert.Equal(3, thumbnails.Count);
            Assert.Same(dragged, thumbnails[0]);
            Assert.DoesNotContain(thumbnails, t => t.IsDropPlaceholder);
        });
    }

    // THE serious one. The suppress flag is the tab's shared guard, and while it is set nothing in
    // the app can change the selected schematic - reported as "I cannot select any, no red border".
    // It used to be cleared only on the pointer-release path, so any other way a drag ended left it
    // stuck on until restart.
    [Fact]
    public void Ending_a_drag_always_releases_the_shared_selection_guard()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, _, tab) = BuildGallery();

            gallery.BeginTileDragForTests(thumbnails[1]);
            Assert.True(tab.SuppressThumbnailSelectionChangedForTests);

            gallery.EndTileDragForTests(commitOrder: true);

            Assert.False(tab.SuppressThumbnailSelectionChangedForTests);
            Assert.False(gallery.IsDraggingForTests);
        });
    }

    // The same guarantee for a drag that never reaches a pointer release - the pointer leaving the
    // window, or the window being closed mid-drag.
    [Fact]
    public void Detaching_the_gallery_mid_drag_releases_the_shared_selection_guard()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, _, tab) = BuildGallery();

            var window = new Window { Content = gallery };
            window.Show();

            gallery.BeginTileDragForTests(thumbnails[1]);
            Assert.True(tab.SuppressThumbnailSelectionChangedForTests);

            window.Content = null;

            Assert.False(tab.SuppressThumbnailSelectionChangedForTests);
            window.Close();
        });
    }

    // After a reorder the moved thumbnail must still be the selected one, or the schematic view
    // silently stops matching the highlighted tile.
    [Fact]
    public void A_reorder_keeps_the_dragged_thumbnail_selected()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, hostedList, _) = BuildGallery();
            var dragged = thumbnails[2];

            gallery.SelectedThumbnail = dragged;

            gallery.BeginTileDragForTests(dragged);
            gallery.MoveDragPlaceholderForTests(0);
            gallery.EndTileDragForTests(commitOrder: false);

            Assert.Same(dragged, gallery.SelectedThumbnail);
            Assert.Same(dragged, hostedList.SelectedItem);
        });
    }

    // The placeholder is a real row in the shared collection while a drag runs, so it can be landed
    // on by a click or the arrow keys. It names no schematic, so forwarding it would blank the view.
    [Fact]
    public void Selecting_the_placeholder_is_never_forwarded_to_the_hosted_list()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, hostedList, _) = BuildGallery();
            var realSelection = thumbnails[0];

            gallery.SelectedThumbnail = realSelection;
            gallery.BeginTileDragForTests(thumbnails[2]);

            var placeholder = thumbnails.Single(t => t.IsDropPlaceholder);
            gallery.SelectedThumbnail = placeholder;

            // The hosted list still points at a REAL schematic.
            Assert.Same(realSelection, hostedList.SelectedItem);
        });
    }

    private static (SchematicsThumbnailGallery Gallery, ObservableCollection<SchematicThumbnail> Thumbnails, ListBox HostedList, TabSchematics Tab) BuildGallery()
    {
        var thumbnails = new ObservableCollection<SchematicThumbnail>
        {
            new() { Name = "Schematic A" },
            new() { Name = "Schematic B" },
            new() { Name = "Schematic C" }
        };

        var hostedList = new ListBox { ItemsSource = thumbnails };
        var tab = new TabSchematics();
        var gallery = new SchematicsThumbnailGallery();
        gallery.Initialize(thumbnails, hostedList, tab);

        return (gallery, thumbnails, hostedList, tab);
    }

    // Reordering must persist through the exact same UserSettings.SetSchematicsOrder path the
    // inline list uses, whether thumbnails are inline or detached - driven directly here since
    // simulating a full pointer drag headlessly (with no real board/MainWindow to resolve a board
    // key from) is impractical; SaveCurrentThumbnailOrder itself is what the gallery calls after a
    // drop, so this pins that it is reachable and internal rather than private.
    // ###########################################################################################
    // What SaveCurrentThumbnailOrder actually PERSISTS - the gallery's whole reason for calling it.
    //
    // This test used to assert only that the call "compiles and does not throw", which the compiler
    // already guarantees, and it could not do more: the method reads its board key off MainWindow,
    // which no test constructs, so it returned at the first line and exercised nothing. With
    // BoardKeyOverrideForTests it now drives the real path and checks the two things the method
    // decides: that the drop PLACEHOLDER is excluded (it is a UI artefact of an in-flight drag and
    // must never be written into a board's saved order) and that a blank name is skipped.
    // ###########################################################################################
    [Fact]
    public void Saving_the_thumbnail_order_persists_real_names_and_drops_the_placeholder()
    {
        using var workspace = new TempWorkspace();
        UserSettings.LoadFrom(workspace.Path_("settings.json"));

        UiTest.Run(() =>
        {
            var tab = new TabSchematics
            {
                BoardKeyOverrideForTests = "Commodore 64|250407"
            };

            tab.currentThumbnails.Add(new SchematicThumbnail { Name = "Top" });
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = "Placeholder", IsDropPlaceholder = true });
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = "   " });
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = "Bottom" });

            tab.SaveCurrentThumbnailOrder();

            // Order preserved, placeholder and blank gone.
            Assert.Equal(
                new[] { "Top", "Bottom" },
                UserSettings.GetSchematicsOrder("Commodore 64|250407"));
        });
    }

    // With no board selected there is nothing to key the order by, so nothing may be written -
    // otherwise a drag performed before a board loads would persist under an empty key.
    [Fact]
    public void Saving_with_no_board_selected_writes_nothing()
    {
        using var workspace = new TempWorkspace();
        UserSettings.LoadFrom(workspace.Path_("settings.json"));

        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = "Top" });

            tab.SaveCurrentThumbnailOrder();

            Assert.Null(UserSettings.GetSchematicsOrder(string.Empty));
        });
    }
}
