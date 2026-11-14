using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// BadgeCloud - "KiCad: Files"'s lists of badges, drawn by ONE control (2026-11-13, measured: ~1,150 badges on the Open128-2
// were a Border each). Moved here from KiCadDataPanelTests on 2026-11-14 (code review: one test file per class).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BadgeCloudTests
{
    // ###########################################################################################
    // A NET ALREADY USED IS DRAWN FILLED GREEN AND SAYS SO UNDER ITSELF (owner report, 2026-10-29), now that the list draws
    // its badges itself (BadgeCloud, 2026-11-13): a used net's badge is of the Used kind, an unused one plain, and the
    // pointer on the used one gives the tooltip - moved to stand under that badge, never higher than the control's bottom
    // - while the pointer on an unused one, or off them all, gives none.
    // ###########################################################################################
    [Fact]
    public void A_used_net_is_drawn_green_and_says_so_under_itself_while_the_pointer_is_on_it()
    {
        UiTest.Run(() =>
        {
            var cloud = new BadgeCloud
            {
                ItemsSource = new[]
                {
                    new KiCadNetNameBadge { Name = "A1", IsUsed = true },
                    new KiCadNetNameBadge { Name = "A10" },
                    new KiCadNetNameBadge { Name = "+5V", IsUsed = true },
                },
            };

            var window = new Window { Content = new StackPanel { Width = 90, Children = { cloud } }, Width = 400, Height = 300 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            var badges = cloud.BadgesForTests;
            Assert.Equal([BadgeCloudKind.Used, BadgeCloudKind.Plain, BadgeCloudKind.Used], badges.Select(badge => badge.Kind));
            Assert.Equal(["A1", "A10", "+5V"], badges.Select(badge => badge.Text));

            // Wrapped in the 90 wide: the third badge on a line of its own.
            Assert.True(badges[2].Place.Y > badges[0].Place.Y);

            cloud.PointAt(badges[2].Place.Center);
            Assert.Equal("Already used in \"KiCad: Important signals\"", ToolTip.GetTip(cloud));
            Assert.Equal(badges[2].Place.Center.X - (cloud.Bounds.Width / 2), ToolTip.GetHorizontalOffset(cloud), precision: 3);
            Assert.Equal(badges[2].Place.Bottom - cloud.Bounds.Height, ToolTip.GetVerticalOffset(cloud), precision: 3);
            Assert.True(ToolTip.GetVerticalOffset(cloud) <= 0);

            cloud.PointAt(badges[1].Place.Center);
            Assert.Null(ToolTip.GetTip(cloud));

            cloud.PointAt(badges[0].Place.Center);
            Assert.NotNull(ToolTip.GetTip(cloud));

            cloud.PointAt(null);
            Assert.Null(ToolTip.GetTip(cloud));

            window.Close();
        });
    }

    // ###########################################################################################
    // The badges follow a change of the colors over them, not a change of theme alone (code review, 2026-11-13): the
    // "User preference" theme, applied or reloaded, keeps the light variant and changes only the resources - and the
    // badges, their text color drawn into them as they are built, kept the colors before.
    // ###########################################################################################
    [Fact]
    public void A_badge_cloud_follows_a_change_of_the_colors_over_it()
    {
        UiTest.Run(() =>
        {
            var cloud = new BadgeCloud { ItemsSource = new[] { new KiCadNetNameBadge { Name = "A1" } } };

            var window = new Window { Content = cloud, Width = 400, Height = 300 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            window.Resources["Fg"] = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Colors.Red);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(Avalonia.Media.Colors.Red, Assert.IsAssignableFrom<Avalonia.Media.ISolidColorBrush>(cloud.BadgesForTests.Single().TextBrush).Color);

            window.Close();
        });
    }

    // Badges made again let go of what the one under the pointer said (code review, 2026-11-13): an unused net standing
    // where a used one stood kept its "Already used" tip until the pointer moved.
    [Fact]
    public void Badges_made_again_let_go_of_the_tip_the_badge_under_the_pointer_gave()
    {
        UiTest.Run(() =>
        {
            var cloud = new BadgeCloud { ItemsSource = new[] { new KiCadNetNameBadge { Name = "A1", IsUsed = true } } };

            var window = new Window { Content = cloud, Width = 400, Height = 300 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            cloud.PointAt(cloud.BadgesForTests[0].Place.Center);
            Assert.NotNull(ToolTip.GetTip(cloud));

            cloud.ItemsSource = new[] { new KiCadNetNameBadge { Name = "A1" } };
            Dispatcher.UIThread.RunJobs();

            Assert.Null(ToolTip.GetTip(cloud));

            window.Close();
        });
    }
}
