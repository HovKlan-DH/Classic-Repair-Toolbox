using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// SchematicsThumbnailGallery wherever it is shown. Since 2026-10-15 (owner request: the tab's own
// thumbnail panel "should behave exactly as the detached window") it is the Schematics tab's
// thumbnail column as well as the detached window's content and the fullscreen stand-in's, so
// three of them can be alive at once, all showing the one selection the tab holds.
//
// What these pin: the tab's column IS that auto-fit gallery (several columns once wide, never
// scrolled); a click, a tap, a Ctrl+click and a drag on it with real pointer input; it keeps
// working after the tab leaves the visual tree and comes back (every tab switch does that), and
// out of the tree it neither listens nor changes anything; the label editor's lock reaches every
// gallery; one gallery never lifts the shared drag guard another holds. The drag steps through
// their seams are covered in SchematicsThumbnailsDetachTests, and the auto-fit maths in
// ThumbnailGalleryGeometryTests.
//
// COLLECTION NOTE: "HeadlessUi" because it constructs TabSchematics and shows windows.
[Collection("HeadlessUi")]
public sealed class SchematicsThumbnailGalleryTests
{
    private static SchematicsThumbnailGallery InlineGalleryOf(TabSchematics tab) =>
        tab.GetControl<SchematicsThumbnailGallery>("InlineThumbnailGallery");

    private static ListBox GalleryListOf(SchematicsThumbnailGallery gallery) =>
        gallery.GetControl<ListBox>("GalleryListBox");

    // A tab holding three thumbnails, as LoadSortedThumbnails leaves it.
    private static TabSchematics BuildTabWithThumbnails(int count = 3)
    {
        var tab = new TabSchematics();

        for (int index = 0; index < count; index++)
        {
            tab.currentThumbnails.Add(new SchematicThumbnail { Name = $"Schematic {index + 1}" });
        }

        return tab;
    }

    private static void Layout(Window window)
    {
        window.Measure(new Size(window.Width, window.Height));
        window.Arrange(new Rect(0, 0, window.Width, window.Height));
        Dispatcher.UIThread.RunJobs();
    }

    // The tab shown in a window with a 300px thumbnail column, its three tiles stacked one above
    // the other, so a test can press, drag and tap them with real pointer input.
    private static (Window Window, TabSchematics Tab) ShowTabWithThumbnails()
    {
        var tab = BuildTabWithThumbnails();
        tab.GetControl<Grid>("SchematicsInnerGrid").ColumnDefinitions[2].Width = new GridLength(300, GridUnitType.Pixel);

        var window = new Window { Content = tab, Width = 1000, Height = 600 };
        window.Show();
        Layout(window);

        return (window, tab);
    }

    // On the tile's name strip, which every tile draws - these test thumbnails have no image.
    private static Point NameStripOfTile(Window window, TabSchematics tab, int index) =>
        PointInTile(window, tab, index, tile => new Point(tile.Width / 2, 10));

    // The middle of the tile: with no image drawn there, only the tile's own background is.
    private static Point CentreOfTile(Window window, TabSchematics tab, int index) =>
        PointInTile(window, tab, index, tile => new Point(tile.Width / 2, tile.Height / 2));

    private static Point PointInTile(Window window, TabSchematics tab, int index, System.Func<Size, Point> where)
    {
        var container = GalleryListOf(InlineGalleryOf(tab)).ContainerFromIndex(index);
        Assert.NotNull(container);

        var point = container!.TranslatePoint(where(container.Bounds.Size), window);
        Assert.NotNull(point);
        return point!.Value;
    }

    // -----------------------------------------------------------------------------------------
    // The tab's own thumbnail column
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_tabs_thumbnail_column_is_the_gallery()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails();
            var gallery = InlineGalleryOf(tab);

            Assert.True(gallery.IsVisible);
            Assert.Equal(2, Grid.GetColumn(gallery));

            // Bound to the tab's own collection, not a copy, so board switches reach it.
            Assert.Same(tab.currentThumbnails, GalleryListOf(gallery).ItemsSource);
        });
    }

    // What the owner asked for: widening the panel spreads the thumbnails over several columns
    // and every one of them stays in view, exactly as in the detached window; a narrow panel keeps
    // one column. The panel's own maths is ThumbnailGalleryGeometry's - this pins that the TAB'S
    // column is laid out by it, which a single-column scrolling list would fail.
    [Theory]
    [InlineData(600.0, true)]
    [InlineData(150.0, false)]
    public void The_tabs_thumbnail_column_fills_its_width_and_never_scrolls(double columnWidth, bool expectsSeveralColumns)
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails(6);
            var grid = tab.GetControl<Grid>("SchematicsInnerGrid");
            grid.ColumnDefinitions[2].Width = new GridLength(columnWidth, GridUnitType.Pixel);

            var window = new Window { Content = tab, Width = 1000, Height = 600 };
            window.Show();

            try
            {
                Layout(window);

                var panel = Assert.IsType<ThumbnailGalleryPanel>(GalleryListOf(InlineGalleryOf(tab)).ItemsPanelRoot);
                var layout = panel.CurrentLayout;

                Assert.Equal(6, layout.Positions.Count);

                if (expectsSeveralColumns)
                    Assert.True(layout.Shape.Columns > 1, $"{layout.Shape.Columns} column(s) in a {columnWidth}px panel");
                else
                    Assert.Equal(1, layout.Shape.Columns);

                // Every thumbnail inside the panel - nothing left below the fold to scroll to.
                foreach (Point topLeft in layout.Positions)
                {
                    Assert.True(
                        topLeft.Y + layout.CellSize.Height <= panel.Bounds.Height + 0.5,
                        $"a cell ends at {topLeft.Y + layout.CellSize.Height}, below the panel's {panel.Bounds.Height}");
                    Assert.True(
                        topLeft.X + layout.CellSize.Width <= panel.Bounds.Width + 0.5,
                        $"a cell ends at {topLeft.X + layout.CellSize.Width}, right of the panel's {panel.Bounds.Width}");
                }
            }
            finally
            {
                window.Close();
            }
        });
    }

    // -----------------------------------------------------------------------------------------
    // Real pointer input on the tab's column (code review, 2026-10-15). Each of these went wrong
    // only through the real press/capture/release path - the ListBoxItem's own selection on press,
    // and the capture being let go before the release was acted on - which no seam-driven test saw.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_mouse_click_on_a_thumbnail_in_the_tabs_column_selects_that_schematic()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                tab.SelectedThumbnail = thumbnails[0];

                var point = NameStripOfTile(window, tab, 2);
                window.MouseDown(point, MouseButton.Left);
                window.MouseUp(point, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.Same(thumbnails[2], tab.SelectedThumbnail);
                Assert.Same(thumbnails[2], InlineGalleryOf(tab).SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A touch or pen press does not select on the press; the release went to the gallery, which
    // holds the capture, and letting go of the capture first ended the click before it was read.
    // A tap selected nothing.
    [Fact]
    public void A_touch_tap_on_a_thumbnail_in_the_tabs_column_selects_that_schematic()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                tab.SelectedThumbnail = thumbnails[0];

                var point = NameStripOfTile(window, tab, 1);
                using (var touch = window.TouchBegin(point))
                {
                    window.TouchEnd(touch, point);
                }

                Dispatcher.UIThread.RunJobs();

                Assert.Same(thumbnails[1], tab.SelectedThumbnail);
                Assert.Same(thumbnails[1], InlineGalleryOf(tab).SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Reordering is not choosing: the press selected the tile, which loaded its schematic into the
    // view, ended any worklog area-marking and saved it as the board's last schematic - and the
    // user, who only moved a tile, was left looking at another schematic.
    [Fact]
    public void Dragging_a_thumbnail_in_the_tabs_column_reorders_it_without_switching_the_schematic()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                var shown = thumbnails[0];
                var dragged = thumbnails[2];
                tab.SelectedThumbnail = shown;

                var selectedOnTheWay = new List<object?>();
                tab.SelectedThumbnailChanged += (_, _) => selectedOnTheWay.Add(tab.SelectedThumbnail);

                DragWithMouse(window, NameStripOfTile(window, tab, 2), NameStripOfTile(window, tab, 0));

                Assert.Same(dragged, thumbnails[0]);
                Assert.Same(shown, tab.SelectedThumbnail);
                Assert.DoesNotContain((object)dragged, selectedOnTheWay);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A tile's image is fitted inside it, so a press can land on the tile where no image is drawn,
    // or on the thin ring between the tile's frame and its cell's edge. Neither was the tile's own:
    // the press went to the list item around it, which selected the schematic at once, and nothing
    // could be dragged from that spot.
    [Theory]
    [InlineData("middle, with no image")]
    [InlineData("outermost pixel of the cell")]
    public void A_thumbnail_can_be_dragged_from_anywhere_in_its_cell(string where)
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                var shown = thumbnails[0];
                var dragged = thumbnails[2];
                tab.SelectedThumbnail = shown;

                System.Func<int, Point> pointIn = where.StartsWith("middle", System.StringComparison.Ordinal)
                    ? index => CentreOfTile(window, tab, index)
                    : index => PointInTile(window, tab, index, cell => new Point(0.5, cell.Height / 2));

                DragWithMouse(window, pointIn(2), pointIn(0));

                Assert.Same(dragged, thumbnails[0]);
                Assert.Same(shown, tab.SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The right button opens the panel's "Detach thumbnails" menu; it is not a pick. The list item
    // selected on a right press too, so opening the menu also loaded that schematic.
    [Fact]
    public void A_right_click_on_a_thumbnail_does_not_switch_the_schematic()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                tab.SelectedThumbnail = thumbnails[0];

                var point = NameStripOfTile(window, tab, 2);
                window.MouseDown(point, MouseButton.Right);
                window.MouseUp(point, MouseButton.Right);
                Dispatcher.UIThread.RunJobs();

                Assert.Same(thumbnails[0], tab.SelectedThumbnail);
                Assert.Same(thumbnails[0], InlineGalleryOf(tab).SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Another control taking the pointer mid-drag (Alt+Tab, the OS) is not a drop: the tile goes
    // back where it started, as for every other drag that ends without one.
    [Fact]
    public void Losing_the_pointer_mid_drag_puts_the_thumbnail_back()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var before = tab.currentThumbnails.ToList();

                IPointer? mouse = null;
                window.AddHandler(InputElement.PointerPressedEvent, (_, e) => mouse = e.Pointer, RoutingStrategies.Tunnel, handledEventsToo: true);

                var from = NameStripOfTile(window, tab, 2);
                var to = NameStripOfTile(window, tab, 0);

                window.MouseDown(from, MouseButton.Left);
                window.MouseMove(from + new Point(0, -20), RawInputModifiers.LeftMouseButton);
                Layout(window);
                window.MouseMove(to, RawInputModifiers.LeftMouseButton);
                Layout(window);

                Assert.NotNull(mouse);
                mouse!.Capture(window);
                window.MouseUp(to, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(before, tab.currentThumbnails);
                Assert.DoesNotContain(tab.currentThumbnails, t => t.IsDropPlaceholder);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Only the button that started a drag ends it. The right button pressed and let go halfway
    // dropped the tile there, and the rest of the drag did nothing.
    [Fact]
    public void Another_button_let_go_during_a_drag_does_not_drop_it()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                var dragged = thumbnails[2];

                var from = NameStripOfTile(window, tab, 2);
                var halfway = NameStripOfTile(window, tab, 0);
                var to = NameStripOfTile(window, tab, 1);

                window.MouseDown(from, MouseButton.Left);
                window.MouseMove(from + new Point(0, -20), RawInputModifiers.LeftMouseButton);
                Layout(window);
                window.MouseMove(halfway, RawInputModifiers.LeftMouseButton);
                Layout(window);

                window.MouseDown(halfway, MouseButton.Right, RawInputModifiers.LeftMouseButton);
                window.MouseUp(halfway, MouseButton.Right, RawInputModifiers.LeftMouseButton);

                window.MouseMove(to, RawInputModifiers.LeftMouseButton);
                Layout(window);
                window.MouseUp(to, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.Same(dragged, thumbnails[1]);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Only the finger that pressed a tile completes the tap. A second finger's press took the
    // first one's place, and its release then chose the second tile.
    [Fact]
    public void A_second_finger_does_not_take_over_the_first_fingers_tap()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                tab.SelectedThumbnail = thumbnails[0];

                var first = NameStripOfTile(window, tab, 1);
                var second = NameStripOfTile(window, tab, 2);

                using var firstFinger = window.TouchBegin(first);
                using (var secondFinger = window.TouchBegin(second))
                {
                    window.TouchEnd(secondFinger, second);
                }

                Dispatcher.UIThread.RunJobs();
                Assert.Same(thumbnails[0], tab.SelectedThumbnail);

                window.TouchEnd(firstFinger, first);
                Dispatcher.UIThread.RunJobs();

                Assert.Same(thumbnails[1], tab.SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The arrow keys move through the thumbnails as they are laid out, as they did through the tab's
    // old one-column list. The gallery's grid panel did not say how to move through it, so the
    // keys did nothing.
    [Fact]
    public void The_arrow_keys_move_the_selection_through_the_thumbnails()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;

                var point = NameStripOfTile(window, tab, 0);
                window.MouseDown(point, MouseButton.Left);
                window.MouseUp(point, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();
                Assert.Same(thumbnails[0], tab.SelectedThumbnail);

                window.KeyPress(Key.Down, RawInputModifiers.None, PhysicalKey.ArrowDown, keySymbol: null);
                Dispatcher.UIThread.RunJobs();
                Assert.Same(thumbnails[1], tab.SelectedThumbnail);

                window.KeyPress(Key.End, RawInputModifiers.None, PhysicalKey.End, keySymbol: null);
                Dispatcher.UIThread.RunJobs();
                Assert.Same(thumbnails[2], tab.SelectedThumbnail);

                window.KeyPress(Key.Up, RawInputModifiers.None, PhysicalKey.ArrowUp, keySymbol: null);
                Dispatcher.UIThread.RunJobs();
                Assert.Same(thumbnails[1], tab.SelectedThumbnail);
                Assert.Same(thumbnails[1], InlineGalleryOf(tab).SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Past the drag threshold first, then to the target, with a layout pass after each step - the
    // gallery finds the nearest cell from the panel's last layout.
    private static void DragWithMouse(Window window, Point from, Point to)
    {
        window.MouseDown(from, MouseButton.Left);
        window.MouseMove(from + new Point(0, -20), RawInputModifiers.LeftMouseButton);
        Layout(window);
        window.MouseMove(to, RawInputModifiers.LeftMouseButton);
        Layout(window);
        window.MouseUp(to, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    // Ctrl+click (Cmd+click on macOS) toggled the selected tile OFF on the press, the gallery
    // passed that "nothing selected" on, and the schematic view went blank.
    [Fact]
    public void Ctrl_clicking_the_selected_thumbnail_keeps_it_selected()
    {
        UiTest.Run(() =>
        {
            var (window, tab) = ShowTabWithThumbnails();

            try
            {
                var thumbnails = tab.currentThumbnails;
                tab.SelectedThumbnail = thumbnails[1];

                var point = NameStripOfTile(window, tab, 1);
                window.MouseDown(point, MouseButton.Left, RawInputModifiers.Control);
                window.MouseUp(point, MouseButton.Left, RawInputModifiers.Control);
                Dispatcher.UIThread.RunJobs();

                Assert.Same(thumbnails[1], tab.SelectedThumbnail);
                Assert.Same(thumbnails[1], InlineGalleryOf(tab).SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A TabControl takes the tab out of the visual tree on every switch away and puts it back on
    // the way in, and the gallery let go of the selection for good on the first detach - the
    // right thing for a closed detached window, and a dead thumbnail column after one tab switch.
    // Fails against that version: the selection made meanwhile (the Workbooks tab picking a
    // schematic) never shows, and a click afterwards selects nothing.
    [Fact]
    public void The_tabs_column_keeps_working_after_the_tab_was_switched_away_and_back()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails();
            var gallery = InlineGalleryOf(tab);
            var thumbnails = tab.currentThumbnails;

            var window = new Window { Content = tab, Width = 800, Height = 600 };
            window.Show();

            try
            {
                tab.SelectedThumbnail = thumbnails[0];
                Assert.Same(thumbnails[0], gallery.SelectedThumbnail);

                // Away from the tab, and a schematic chosen elsewhere meanwhile.
                window.Content = null;
                tab.SelectedThumbnail = thumbnails[1];

                window.Content = tab;
                Layout(window);

                // Caught up on what happened while away...
                Assert.Same(thumbnails[1], gallery.SelectedThumbnail);

                // ...and working again in both directions - its own list moving the highlight, as
                // the arrow keys do, asks the tab for that schematic.
                GalleryListOf(gallery).SelectedItem = thumbnails[2];
                Assert.Same(thumbnails[2], tab.SelectedThumbnail);

                tab.SelectedThumbnail = thumbnails[0];
                Assert.Same(thumbnails[0], gallery.SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Out of the visual tree it must not stay subscribed to the tab, which lives as long as the
    // application - that is what kept closed detached windows alive. It catches up instead.
    [Fact]
    public void A_gallery_out_of_the_visual_tree_does_not_follow_the_selection()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, tab) = BuildGallery();

            var window = new Window { Content = gallery };
            window.Show();

            try
            {
                window.Content = null;

                tab.SelectedThumbnail = thumbnails[1];
                Assert.Null(gallery.SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Nor may it push its own, out-of-date highlight back. A gallery taken out of the tree with X
    // highlighted (the tab's column while another tab is shown) dropped X when X left the shared
    // collection - a drag in the detached window lifts it out - and passed that "nothing selected"
    // on, wiping the schematic chosen meanwhile.
    [Fact]
    public void A_gallery_out_of_the_visual_tree_never_changes_the_selection()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails();
            var thumbnails = tab.currentThumbnails;
            var gallery = InlineGalleryOf(tab);

            var window = new Window { Content = tab, Width = 800, Height = 600 };
            window.Show();

            try
            {
                tab.SelectedThumbnail = thumbnails[0];
                Assert.Same(thumbnails[0], gallery.SelectedThumbnail);

                window.Content = null;
                tab.SelectedThumbnail = thumbnails[1];

                // The tile the stale gallery still highlights leaves the collection.
                thumbnails.RemoveAt(0);
                Dispatcher.UIThread.RunJobs();

                Assert.Same(thumbnails[0], tab.SelectedThumbnail);
                Assert.Equal("Schematic 2", tab.SelectedThumbnail!.Name);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // -----------------------------------------------------------------------------------------
    // The label editor's lock
    // -----------------------------------------------------------------------------------------

    // The editor locks picking another schematic, and every gallery follows the tab's lock. Only
    // the tab's own thumbnail list of the time honoured it - the detached window's gallery passed
    // clicks on to a disabled list, whose SelectedItem a program can still set, and changed the
    // schematic under an open editor. The tab itself refuses a pick while locked, whichever gallery
    // asks.
    [Fact]
    public void The_tabs_lock_locks_the_gallery_and_drops_its_drag()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, tab) = BuildGallery();
            var galleryList = GalleryListOf(gallery);

            var window = new Window { Content = gallery };
            window.Show();

            try
            {
                tab.SelectedThumbnail = thumbnails[0];

                gallery.BeginTileDragForTests(thumbnails[1]);
                Assert.True(tab.SuppressThumbnailSelectionChangedForTests);

                tab.ThumbnailNavigationLockedForTests = true;

                Assert.False(galleryList.IsEnabled);
                Assert.False(gallery.IsDraggingForTests);
                Assert.DoesNotContain(thumbnails, t => t.IsDropPlaceholder);
                Assert.False(tab.SuppressThumbnailSelectionChangedForTests);

                tab.SelectThumbnailFromGallery(thumbnails[2]);
                Assert.Same(thumbnails[0], tab.SelectedThumbnail);

                tab.ThumbnailNavigationLockedForTests = false;

                Assert.True(galleryList.IsEnabled);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A detached window opened while the editor is already open starts locked.
    [Fact]
    public void A_gallery_wired_to_a_locked_tab_starts_locked()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics { ThumbnailNavigationLockedForTests = true };

            var gallery = new SchematicsThumbnailGallery();
            gallery.Initialize(tab);

            Assert.False(GalleryListOf(gallery).IsEnabled);
        });
    }

    // Its own list's highlight can move by itself, or by a key, to something the tab does not
    // hold - here the arrow keys reach another tile while a drag in another gallery holds the
    // guard, so the tab refuses it. The gallery shows the tab's selection again rather than a tile
    // that is not the schematic on screen.
    [Fact]
    public void A_pick_the_tab_refuses_does_not_stay_highlighted()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, tab) = BuildGallery();

            var window = new Window { Content = gallery };
            window.Show();

            try
            {
                tab.SelectedThumbnail = thumbnails[0];
                tab.SuppressThumbnailSelectionChangedForTests = true;

                GalleryListOf(gallery).SelectedItem = thumbnails[2];
                Dispatcher.UIThread.RunJobs();

                Assert.Same(thumbnails[0], tab.SelectedThumbnail);
                Assert.Same(thumbnails[0], gallery.SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // -----------------------------------------------------------------------------------------
    // Several galleries, one drag guard
    // -----------------------------------------------------------------------------------------

    // Up to three galleries share the tab's guard. Leaving the visual tree ends a drag and used to
    // release the guard unconditionally - so the tab's column, taken out by a tab switch, would
    // lift the guard of a drag running in the detached window. Fails against that version.
    [Fact]
    public void A_gallery_with_no_drag_of_its_own_leaves_another_gallerys_guard_alone()
    {
        UiTest.Run(() =>
        {
            var (dragging, thumbnails, tab) = BuildGallery();

            var idle = new SchematicsThumbnailGallery();
            idle.Initialize(tab);

            var window = new Window { Content = idle };
            window.Show();

            try
            {
                dragging.BeginTileDragForTests(thumbnails[1]);

                idle.CancelTileDrag();
                window.Content = null;

                Assert.True(tab.SuppressThumbnailSelectionChangedForTests);
                Assert.True(dragging.IsDraggingForTests);

                dragging.EndTileDragForTests(commitOrder: false);
                Assert.False(tab.SuppressThumbnailSelectionChangedForTests);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A board load resets the viewer and empties the thumbnails. It used to end only the tab's own
    // column's drag; a drag in the fullscreen stand-in kept the guard, so the new board's schematic
    // never loaded and the view stayed blank.
    [Fact]
    public void Resetting_the_viewer_ends_a_drag_in_every_gallery()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails();

            var placeholder = new SchematicsFullscreenPlaceholder();
            placeholder.Initialize(tab, 0.7);
            var standInGallery = placeholder.GetControl<SchematicsThumbnailGallery>("ThumbnailGalleryPlaceholder");

            var window = new Window { Content = placeholder, Width = 800, Height = 600 };
            window.Show();

            try
            {
                standInGallery.BeginTileDragForTests(tab.currentThumbnails[1]);
                Assert.True(tab.SuppressThumbnailSelectionChangedForTests);

                tab.ResetSchematicsViewer();

                Assert.False(standInGallery.IsDraggingForTests);
                Assert.False(tab.SuppressThumbnailSelectionChangedForTests);
                Assert.DoesNotContain(tab.currentThumbnails, t => t.IsDropPlaceholder);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A press held over a board load ends on a thumbnail of the board that is gone. Picking it
    // blanked the view, since nothing of the new board was selected any more.
    [Fact]
    public void A_thumbnail_no_longer_on_the_board_is_never_picked()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails();
            var current = tab.currentThumbnails[0];
            tab.SelectedThumbnail = current;

            tab.SelectThumbnailFromGallery(new SchematicThumbnail { Name = "From the board before" });

            Assert.Same(current, tab.SelectedThumbnail);
        });
    }

    // Teardown let go of the tab without ending a drag, and only a gallery that still has its tab
    // can release the guard - so a Teardown mid-drag would have left schematic selection dead
    // everywhere until restart. Safe today only because a closing window detaches first.
    [Fact]
    public void Tearing_a_gallery_down_mid_drag_releases_the_guard()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, tab) = BuildGallery();
            var before = thumbnails.ToList();

            gallery.BeginTileDragForTests(thumbnails[1]);
            gallery.Teardown();

            Assert.False(tab.SuppressThumbnailSelectionChangedForTests);
            Assert.Equal(before, thumbnails);
        });
    }

    // The selection is the tab's, not a list bound to the tiles: lifting the selected tile out for
    // a drag emptied it, and for the whole drag the tab had no schematic - its name label and
    // component labels were cleared by anything that redrew them meanwhile (F11 in the detached
    // window), and nothing put them back after the drop.
    [Fact]
    public void Dragging_the_selected_thumbnail_keeps_it_selected_throughout()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, tab) = BuildGallery();
            var dragged = thumbnails[2];
            tab.SelectedThumbnail = dragged;

            gallery.BeginTileDragForTests(dragged);
            gallery.MoveDragPlaceholderForTests(0);

            Assert.Same(dragged, tab.SelectedThumbnail);
            Assert.Equal("Schematic 3", tab.GetCurrentSchematicName());

            gallery.EndTileDragForTests(commitOrder: true);

            Assert.Same(dragged, tab.SelectedThumbnail);
        });
    }

    // The tab cancels its column's drag when the board is reset or the column collapses, F11
    // cancels it, and so do the label editor's lock and a tab switch. A cancelled drag saves no
    // order, so the tile goes back where it STARTED: left where the gap had got to, every gallery
    // showed an order the next board load quietly undid - or the next, unrelated drag saved.
    [Fact]
    public void Cancelling_a_drag_puts_the_thumbnail_back_where_it_started_and_releases_the_guard()
    {
        UiTest.Run(() =>
        {
            var (gallery, thumbnails, tab) = BuildGallery();
            var before = thumbnails.ToList();

            gallery.BeginTileDragForTests(before[2]);
            gallery.MoveDragPlaceholderForTests(0);

            gallery.CancelTileDrag();

            Assert.False(gallery.IsDraggingForTests);
            Assert.False(tab.SuppressThumbnailSelectionChangedForTests);
            Assert.Equal(before, thumbnails);
        });
    }

    // -----------------------------------------------------------------------------------------
    // The fullscreen stand-in
    // -----------------------------------------------------------------------------------------

    // While the schematic is fullscreen the tab shows a stand-in; its thumbnails are the same
    // gallery as the tab's column, so they look and behave alike there too.
    [Fact]
    public void The_fullscreen_stand_in_shows_the_same_gallery_and_selects_through_it()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails();

            var placeholder = new SchematicsFullscreenPlaceholder();
            placeholder.Initialize(tab, 0.7);

            var gallery = placeholder.GetControl<SchematicsThumbnailGallery>("ThumbnailGalleryPlaceholder");
            Assert.Same(tab.currentThumbnails, GalleryListOf(gallery).ItemsSource);

            var window = new Window { Content = placeholder, Width = 800, Height = 600 };
            window.Show();

            try
            {
                GalleryListOf(gallery).SelectedItem = tab.currentThumbnails[1];
                Assert.Same(tab.currentThumbnails[1], tab.SelectedThumbnail);

                tab.SelectedThumbnail = tab.currentThumbnails[2];
                Assert.Same(tab.currentThumbnails[2], gallery.SelectedThumbnail);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // F11 from another tab builds a stand-in the Schematics tab never shows, so it never leaves
    // the visual tree either. Following the tab from the moment it was wired kept every such
    // stand-in alive, and at work on every schematic change, for the rest of the session.
    [Fact]
    public void A_fullscreen_stand_in_that_is_never_shown_does_not_follow_the_selection()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTabWithThumbnails();
            tab.SelectedThumbnail = tab.currentThumbnails[0];

            var placeholder = new SchematicsFullscreenPlaceholder();
            placeholder.Initialize(tab, 0.7);

            var gallery = placeholder.GetControl<SchematicsThumbnailGallery>("ThumbnailGalleryPlaceholder");

            // It starts out showing the selection as it is...
            Assert.Same(tab.currentThumbnails[0], gallery.SelectedThumbnail);

            // ...and does not listen for changes until it is on screen.
            tab.SelectedThumbnail = tab.currentThumbnails[1];
            Assert.Same(tab.currentThumbnails[0], gallery.SelectedThumbnail);
        });
    }

    // A gallery of its own on a tab holding three thumbnails - as the detached window's is. Not
    // shown: a test that needs it listening to the tab puts it in a window.
    private static (SchematicsThumbnailGallery Gallery, ObservableCollection<SchematicThumbnail> Thumbnails, TabSchematics Tab) BuildGallery()
    {
        var tab = BuildTabWithThumbnails();

        var gallery = new SchematicsThumbnailGallery();
        gallery.Initialize(tab);

        return (gallery, tab.currentThumbnails, tab);
    }
}
