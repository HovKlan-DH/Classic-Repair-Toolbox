using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// The "#N" pills the thumbnail gallery draws for the worklog entries on each schematic.
//
// The rule these exist for: a thumbnail must draw an entry the SAME way the main schematic view
// does. An entry with "Show marked area" ticked gets its pill on its marked area; an entry with it
// UNTICKED has no area drawn at all, so its pill is parked in the image's top-right corner
// instead. Drawing an unticked entry's pill at its marker regardless was a reported bug - the main
// view parked it and the thumbnail did not, so the two views disagreed about the same entry, and
// the thumbnail pointed at a location nothing was marking.
//
// Asserted through LayOutPills rather than through the rendered output: this overlay draws straight
// to a DrawingContext, so there is no control on the visual tree carrying a position, and pixel
// sampling a RenderTargetBitmap needs a display.
[Collection("HeadlessUi")]
public class ThumbnailWorklogPillsTests
{
    // A square image in a square control, so the content rect fills the control exactly and every
    // expected position can be reasoned about without letterboxing in the way. The letterboxed case
    // has its own test below.
    private const double OverlaySize = 200;

    private static readonly PixelSize BitmapSize = new(400, 400);

    // Lays the overlay out for real - LayOutPills reads Bounds, which is only meaningful after an
    // arrange pass - and hands back what it placed.
    private static void WithOverlay(
        IReadOnlyList<ThumbnailWorklogPillsOverlay.Pill> pills,
        Action<IReadOnlyList<(ThumbnailWorklogPillsOverlay.Pill Pill, Point Center, FormattedText Text)>> body,
        double width = OverlaySize,
        double height = OverlaySize,
        PixelSize? bitmapSize = null)
    {
        UiTest.Run(() =>
        {
            var overlay = new ThumbnailWorklogPillsOverlay
            {
                BitmapPixelSize = bitmapSize ?? BitmapSize,
                Pills = pills
            };

            var window = new Window { Width = width, Height = height, Content = overlay };

            try
            {
                window.Show();
                window.Measure(new Size(width, height));
                window.Arrange(new Rect(0, 0, width, height));
                Dispatcher.UIThread.RunJobs();

                body(overlay.LayOutPills());
            }
            finally
            {
                window.Close();
            }
        });
    }

    private static ThumbnailWorklogPillsOverlay.Pill Pill(double x, double y, int id, bool isParked) =>
        new(new Point(x, y), Colors.IndianRed, id, isParked);

    [Fact]
    public void No_pills_lays_out_nothing()
    {
        WithOverlay(Array.Empty<ThumbnailWorklogPillsOverlay.Pill>(), placed => Assert.Empty(placed));
    }

    // A "show marked area" ON entry keeps the behaviour that already worked: its pill sits on the
    // entry's own area, scaled from bitmap pixels into the thumbnail's content rect.
    [Fact]
    public void A_shown_entrys_pill_sits_at_its_marked_area()
    {
        // (100,300) in a 400x400 bitmap drawn into a 200x200 control is (50,150).
        WithOverlay(new[] { Pill(100, 300, 1, isParked: false) }, placed =>
        {
            var (_, center, _) = Assert.Single(placed);

            Assert.Equal(50, center.X, precision: 3);
            Assert.Equal(150, center.Y, precision: 3);
        });
    }

    // THE REPORTED BUG. The entry's marker is at the BOTTOM-LEFT of the board; if the parked pill
    // were still drawn at the marker it would land there. Asserting where it actually IS - not
    // merely that it moved - is what makes this test fail against the version that shipped.
    [Fact]
    public void A_hidden_entrys_pill_is_parked_in_the_top_right_not_at_its_marker()
    {
        WithOverlay(new[] { Pill(20, 380, 1, isParked: true) }, placed =>
        {
            var (_, center, _) = Assert.Single(placed);

            // The marker maps to (10,190) - bottom-left. The pill must be nowhere near it.
            Assert.True(center.X > OverlaySize / 2, $"parked pill should be on the right, was at X={center.X}");
            Assert.True(center.Y < OverlaySize / 4, $"parked pill should be near the top, was at Y={center.Y}");
        });
    }

    // The two branches must genuinely differ for the SAME marker position - otherwise a test that
    // only looked at one of them could pass while both drew identically.
    [Fact]
    public void The_same_marker_lands_in_different_places_depending_on_show_marked_area()
    {
        Point shown = default;
        Point parked = default;

        WithOverlay(new[] { Pill(20, 380, 1, isParked: false) }, placed => shown = placed[0].Center);
        WithOverlay(new[] { Pill(20, 380, 1, isParked: true) }, placed => parked = placed[0].Center);

        Assert.NotEqual(shown, parked);
    }

    // A thumbnail can carry both kinds at once, and the shown one must NOT be dragged into the
    // corner along with the parked one.
    [Fact]
    public void Shown_and_parked_pills_on_one_thumbnail_keep_their_own_placements()
    {
        var shownPill = Pill(100, 300, 1, isParked: false);
        var parkedPill = Pill(100, 300, 2, isParked: true);

        WithOverlay(new[] { shownPill, parkedPill }, placed =>
        {
            Assert.Equal(2, placed.Count);

            var shown = placed.Single(p => p.Pill.EntryId == 1).Center;
            var parked = placed.Single(p => p.Pill.EntryId == 2).Center;

            Assert.Equal(50, shown.X, precision: 3);
            Assert.Equal(150, shown.Y, precision: 3);

            Assert.True(parked.X > OverlaySize / 2, $"parked pill should be on the right, was at X={parked.X}");
            Assert.True(parked.Y < OverlaySize / 4, $"parked pill should be near the top, was at Y={parked.Y}");
        });
    }

    // Parked pills stack rather than piling up on one another - the same ParkedBadgeGeometry block
    // the main view and the Workbooks board pane use, so all three read the same way.
    [Fact]
    public void Two_parked_pills_do_not_land_on_top_of_each_other()
    {
        WithOverlay(new[] { Pill(0, 0, 1, isParked: true), Pill(0, 0, 2, isParked: true) }, placed =>
        {
            Assert.Equal(2, placed.Count);
            Assert.NotEqual(placed[0].Center, placed[1].Center);
        });
    }

    // Every parked pill must stay ON the image. The block is clamped rather than allowed to run off
    // the edge, which matters far more here than on the main view: a thumbnail is a fraction of the
    // size, so even a handful of pills can outgrow the corner.
    [Fact]
    public void Many_parked_pills_all_stay_inside_the_image()
    {
        var pills = Enumerable.Range(1, 12).Select(id => Pill(0, 0, id, isParked: true)).ToArray();

        WithOverlay(pills, placed =>
        {
            Assert.Equal(12, placed.Count);

            foreach (var (pill, center, _) in placed)
            {
                Assert.True(
                    center.X >= 0 && center.X <= OverlaySize,
                    $"parked pill #{pill.EntryId} is off the image horizontally, at X={center.X}");
                Assert.True(
                    center.Y >= 0 && center.Y <= OverlaySize,
                    $"parked pill #{pill.EntryId} is off the image vertically, at Y={center.Y}");
            }
        });
    }

    // A thumbnail letterboxes its image (Stretch="Uniform" in a fixed-size cell), so the control is
    // wider than what is on screen. The parked block must sit against the IMAGE's right edge, not
    // the control's - parking against the control would put the pills out in the empty margin
    // beside the board, floating next to nothing.
    [Fact]
    public void Parked_pills_hug_the_images_edge_not_the_controls_when_letterboxed()
    {
        // A 400x400 bitmap in a 400x200 control: the image is 200 wide, centred, so it occupies
        // x = 100..300 and there is a 100px empty margin on each side.
        WithOverlay(
            new[] { Pill(0, 0, 1, isParked: true) },
            placed =>
            {
                var (_, center, _) = Assert.Single(placed);

                Assert.True(center.X <= 300, $"parked pill at X={center.X} is past the image's right edge at 300");
                Assert.True(center.X >= 250, $"parked pill at X={center.X} is not hugging the image's right edge at 300");
            },
            width: 400,
            height: 200);
    }

    // A pill's colour and id come straight from the caller and must survive both branches - the
    // parked path builds its own FormattedText, and getting the id from the wrong entry there would
    // mislabel the pill without moving it.
    [Fact]
    public void A_parked_pill_keeps_its_own_id_and_colour()
    {
        var pills = new[]
        {
            new ThumbnailWorklogPillsOverlay.Pill(new Point(0, 0), Colors.SteelBlue, 7, IsParked: true),
            new ThumbnailWorklogPillsOverlay.Pill(new Point(0, 0), Colors.Goldenrod, 12, IsParked: true)
        };

        WithOverlay(pills, placed =>
        {
            Assert.Equal(2, placed.Count);

            var seven = placed.Single(p => p.Pill.EntryId == 7);
            var twelve = placed.Single(p => p.Pill.EntryId == 12);

            Assert.Equal(Colors.SteelBlue, seven.Pill.Color);
            Assert.Equal(Colors.Goldenrod, twelve.Pill.Color);

            // FormattedText exposes no way to read its string back, so the label is checked by its
            // laid-out WIDTH instead: "#12" is a digit longer than "#7" and therefore wider. That is
            // enough to catch the failure this guards - a parked pill built from the wrong entry's
            // id, which would give both the same text and so the same width.
            Assert.True(
                twelve.Text.Width > seven.Text.Width,
                $"#12 laid out {twelve.Text.Width}px wide and #7 laid out {seven.Text.Width}px - the two pills are not carrying their own ids");
        });
    }

    // A zero-sized bitmap has no content rect to map into, so nothing is placed rather than a
    // division by zero or a pill at NaN.
    [Fact]
    public void A_bitmap_with_no_size_lays_out_nothing()
    {
        WithOverlay(
            new[] { Pill(0, 0, 1, isParked: true) },
            placed => Assert.Empty(placed),
            bitmapSize: new PixelSize(0, 0));
    }

    // The result comes back in the SAME ORDER as Pills, even though parked pills cannot be
    // positioned until all of them are known and so are laid out in a second phase.
    //
    // Rendering does not care - it just draws each in turn. A caller correlating this list with its
    // own by index would, and the obvious next feature here is click-to-open on a thumbnail pill,
    // exactly as the board pane already has: that caller would map a click on the third pill to the
    // second entry. Appending the parked ones at the end is the version this fails against.
    [Fact]
    public void The_laid_out_pills_come_back_in_the_order_they_were_given()
    {
        var pills = new[]
        {
            Pill(50, 50, 1, isParked: false),
            Pill(60, 60, 2, isParked: true),
            Pill(70, 70, 3, isParked: false),
            Pill(80, 80, 4, isParked: true)
        };

        WithOverlay(pills, placed =>
        {
            Assert.Equal(4, placed.Count);
            Assert.Equal(new[] { 1, 2, 3, 4 }, placed.Select(p => p.Pill.EntryId));

            // Every slot really was filled - a placeholder would come back with a null Text and
            // throw in Render, so this is not merely a tidiness check.
            Assert.All(placed, p => Assert.NotNull(p.Text));

            // And the parked ones are still genuinely parked, not just in order: #2 and #4 must be
            // over in the top-right block rather than at their own markers.
            var parkedTwo = placed.Single(p => p.Pill.EntryId == 2);
            Assert.True(
                parkedTwo.Center.X > OverlaySize / 2 && parkedTwo.Center.Y < OverlaySize / 2,
                $"#2 was placed at {parkedTwo.Center} rather than in the top-right corner");
        });
    }
}
