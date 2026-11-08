using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A LABEL OF ITS OWN IN A GREEN OR RED BUTTON IS WRITTEN AS THE BUTTON'S TEXT IS (owner report, 2026-11-07: the table's
// "Replace all" read black on green - "make sure to have same text colors (everywhere) for buttons (matching its
// background color)"). App.axaml gives every TextBlock the page's color, which outranked what a TextBlock in a
// button's content would take from its button; its style for Button.Accept and Button.Danger hands it the color of the
// button's presenter instead. A tooltip's text keeps the tooltip's color.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ButtonLabelColorTests
{
    private static Color ColorOf(IBrush? brush) => Assert.IsAssignableFrom<ISolidColorBrush>(brush).Color;

    [Theory]
    [InlineData("Accept")]
    [InlineData("Danger")]
    public void A_TextBlock_in_a_green_or_red_buttons_content_takes_the_buttons_text_color(string kind)
    {
        UiTest.Run(() =>
        {
            var alone = new TextBlock { Text = "Replace all" };
            var inPanel = new TextBlock { Text = "Replace all" };
            var tip = new TextBlock { Text = "A tip" };

            var first = new Button { Content = alone, Classes = { kind } };
            var second = new Button { Content = new StackPanel { Orientation = Orientation.Horizontal, Children = { inPanel } }, Classes = { kind } };
            ToolTip.SetTip(second, tip);

            var page = new TextBlock { Text = "Page text" };
            var window = new Window { Content = new StackPanel { Children = { page, first, second } }, Width = 600, Height = 400 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            Color buttonText = ColorOf(first.Foreground);

            Assert.NotEqual(ColorOf(page.Foreground), buttonText);
            Assert.Equal(buttonText, ColorOf(alone.Foreground));
            Assert.Equal(buttonText, ColorOf(inPanel.Foreground));

            ToolTip.SetIsOpen(second, true);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(ColorOf(page.Foreground), ColorOf(tip.Foreground));

            ToolTip.SetIsOpen(second, false);
            window.Close();
        });
    }
}
