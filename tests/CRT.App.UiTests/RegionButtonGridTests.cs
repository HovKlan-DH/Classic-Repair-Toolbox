using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE REGION BUTTONS IN THEIR COLORS (RegionButtonGrid, App.axaml's Button.RegionToggle styles), shown in a real window so the
// app's own styles decide:
//
//   - owner report, 2026-11-11: "in the left-side panel, then the text color is not the same as in "Board: Regions"" - a
//     region of the board's own shows its name in a TextBlock of its own, which App.axaml's style for every TextBlock painted
//     in the window's text color, black on the green of the picked button. It now reads the button's text color.
//   - owner request, 2026-11-11: a region the board gives colors of its own (RegionColors) is lit in them - its background and
//     its text - and one it gives none keeps CRT's.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class RegionButtonGridTests
{
    private static Color BackgroundOf(Button button) =>
        ((ISolidColorBrush)button.GetVisualDescendants().OfType<ContentPresenter>().First().Background!).Color;

    private static Color TextOf(Button button) =>
        ((ISolidColorBrush)((TextBlock)button.Content!).Foreground!).Color;

    [Fact]
    public void A_picked_regions_name_is_in_its_buttons_text_color_and_in_the_boards_own_colors_where_it_gives_any()
    {
        UiTest.Run(() =>
        {
            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,*,*,*,*,*"), RowDefinitions = new RowDefinitions("Auto,Auto") };
            var pal = new Button { Content = "PAL" };
            var ntsc = new Button { Content = "NTSC" };
            grid.Children.Add(pal);
            grid.Children.Add(ntsc);

            var own = new List<Button>();
            Button Make()
            {
                var made = new Button();
                made.Content = RegionButtonGrid.NameBlock(made);
                return made;
            }

            List<Button> buttons = RegionButtonGrid.ButtonsFor(
                ["EU", "US"],
                pal,
                ntsc,
                own,
                Make,
                grid,
                region => region == "EU" ? ("#1F3A93", "#FFFF00") : null);

            foreach (Button button in buttons)
            {
                button.Classes.Set("active", true);
            }

            var window = new Window { Content = grid, Width = 400, Height = 200 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Button eu = buttons[0];
            Button us = buttons[1];

            // The board's own: its background and its text.
            Assert.Contains("BoardColors", eu.Classes);
            Assert.Equal(Color.Parse("#1F3A93"), BackgroundOf(eu));
            Assert.Equal(Color.Parse("#FFFF00"), TextOf(eu));

            // CRT's own for the second region of the board's own - its text white on its blue, as its button says, never the
            // window's text color.
            Assert.DoesNotContain("BoardColors", us.Classes);
            Assert.True(Application.Current!.TryGetResource("Schematics_Region_Own2_Fg", ThemeVariant.Light, out object? fg));
            Assert.Equal(((ISolidColorBrush)fg!).Color, TextOf(us));
            Assert.Equal(((ISolidColorBrush)us.Foreground!).Color, TextOf(us));

            // Unpicked, the name follows the button's plain text color.
            us.Classes.Set("active", false);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(((ISolidColorBrush)us.Foreground!).Color, TextOf(us));

            window.Close();
        });
    }
}
