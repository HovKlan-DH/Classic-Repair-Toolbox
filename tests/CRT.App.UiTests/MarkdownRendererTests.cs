using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Styling;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The changelog drawn as GitHub draws it (MarkdownRenderer, on the About tab): heading sizes and
// lines, the rule, list indentation and bullets per level, the space between blocks, bold, italic,
// struck-through and code pieces, links, alerts and quotes, and the colors in both themes.
//
// What GitHubMarkdown reads is GitHubMarkdownTests'; these take its blocks as given.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MarkdownRendererTests
{
    // -----------------------------------------------------------------------------------------
    // Headings and rules
    // -----------------------------------------------------------------------------------------

    // The text's own size, bold (owner request, 2026-11-06: the version and date "way too large. Please
    // have that normal font size and then just have it bold" - GitHub's 2 and 1.5 times until then), and a
    // line under the first two levels.
    [Fact]
    public void Headings_are_the_texts_size_and_bold_and_the_first_two_are_underlined()
    {
        UiTest.Run(() =>
        {
            var controls = Render("# One\n## Two\n### Three");

            var one = Assert.IsType<Border>(controls[0]);
            Assert.Equal(new Thickness(0, 0, 0, 1), one.BorderThickness);
            var oneText = Assert.IsType<TextBlock>(one.Child);
            Assert.Equal("One", oneText.Text);
            Assert.Equal(14, oneText.FontSize);
            Assert.Equal(FontWeight.Bold, oneText.FontWeight);

            var two = Assert.IsType<Border>(controls[1]);
            Assert.Equal(14, Assert.IsType<TextBlock>(two.Child).FontSize);

            var three = Assert.IsType<TextBlock>(controls[2]);
            Assert.Equal(14, three.FontSize);
            Assert.Equal(FontWeight.Bold, three.FontWeight);
            Assert.Contains("ChangelogHeading", three.Classes);
        });
    }

    // ###########################################################################################
    // A version's heading is plain with only its number bold (owner request, 2026-11-06: "Version <b>3.0.0</b>,
    // 2026-October-11") - BoldAsMarked, from ChangelogVersionHeadings - at the heading's size, with its line.
    // A bold piece of any other heading is as heavy as the heading round it.
    // ###########################################################################################
    [Fact]
    public void A_heading_bold_as_marked_is_plain_with_its_bold_piece_bold()
    {
        UiTest.Run(() =>
        {
            var renderer = new MarkdownRenderer(14, _ => { });
            var heading = ChangelogVersionHeadings.Reworded(Assert.IsType<MarkdownHeading>(Assert.Single(GitHubMarkdown.Parse("# 3.0.0, 2026-October-11"))));

            var border = Assert.IsType<Border>(renderer.Render([heading])[0]);
            var text = Assert.IsType<TextBlock>(border.Child);
            Assert.Equal(FontWeight.Normal, text.FontWeight);
            Assert.Contains("ChangelogHeading", text.Classes);

            var runs = text.Inlines!.OfType<Run>().ToList();
            Assert.Equal(["Version ", "3.0.0", ", 2026-October-11"], runs.Select(run => run.Text));
            Assert.Equal(FontWeight.Bold, runs[1].FontWeight);
            Assert.NotEqual(FontWeight.Bold, runs[0].FontWeight);
            Assert.NotEqual(FontWeight.Bold, runs[2].FontWeight);

            var plain = Assert.IsType<TextBlock>(Assert.IsType<Border>(Render("# Known **issues**")[0]).Child);
            Assert.Equal(FontWeight.Bold, plain.FontWeight);
            Assert.Equal(FontWeight.Bold, plain.Inlines!.OfType<Run>().Single(run => run.Text == "issues").FontWeight);
        });
    }

    // ###########################################################################################
    // A version's heading has ONE line under it (owner request, 2026-11-06: "there are two ... dividers
    // showing - just keep the first one (the thinnest one)"): the changelog's "----" right under each
    // version is left out, the heading's own thin line its divider. A rule anywhere else stays.
    // ###########################################################################################
    [Fact]
    public void The_rule_right_under_a_heading_is_left_out_and_one_elsewhere_stays()
    {
        UiTest.Run(() =>
        {
            var controls = Render("# 4.0.0, 2026-November-?\n------------------------\n- **Application**\n\ntext\n\n---\n\nmore");

            Assert.IsType<Border>(controls[0]);
            Assert.IsType<Grid>(controls[1]);
            Assert.Single(controls, control => control is Border { Child: null, Height: 4 });
            Assert.Equal(5, controls.Count);
        });
    }

    [Fact]
    public void A_rule_is_a_bar_four_pixels_high()
    {
        UiTest.Run(() =>
        {
            var rule = Assert.IsType<Border>(Assert.Single(Render("---")));

            Assert.Equal(4, rule.Height);
            Assert.Null(rule.Child);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Lists
    // -----------------------------------------------------------------------------------------

    // Each level 28px (GitHub's 2em) further in; a disc, a circle, then a square for every level
    // below - the text wrapping in a column of its own, never under the bullet.
    [Fact]
    public void Each_level_of_a_list_is_indented_one_step_further_with_its_own_bullet()
    {
        UiTest.Run(() =>
        {
            var controls = Render("- a\n  - b\n    - c\n      - d");
            string[] bullets = { "\u2022", "\u25E6", "\u25AA", "\u25AA" };

            for (int depth = 0; depth < 4; depth++)
            {
                var item = Assert.IsType<Grid>(controls[depth]);
                Assert.Equal(depth * 28, item.Margin.Left);

                var bullet = Assert.IsType<TextBlock>(item.Children[0]);
                Assert.Equal(bullets[depth], bullet.Text);

                var text = Assert.IsType<TextBlock>(item.Children[1]);
                Assert.Equal(1, Grid.GetColumn(text));
                Assert.Equal(TextWrapping.Wrap, text.TextWrapping);
                Assert.Equal(((char)('a' + depth)).ToString(), text.Text);
            }
        });
    }

    [Fact]
    public void An_ordered_item_shows_its_number()
    {
        UiTest.Run(() =>
        {
            var controls = Render("1. a\n1. b");

            Assert.Equal("1.", Assert.IsType<TextBlock>(Assert.IsType<Grid>(controls[0]).Children[0]).Text);
            Assert.Equal("2.", Assert.IsType<TextBlock>(Assert.IsType<Grid>(controls[1]).Children[0]).Text);
        });
    }

    // GitHub's margins, collapsed as a browser collapses them: 21px above a heading and round a rule, 3px
    // between items, 14px under a list, a paragraph or a heading, nothing above the first block - the rule
    // right under a heading left out (owner request, 2026-11-06), so 14px from the heading to the list.
    [Fact]
    public void The_space_between_blocks_is_GitHubs_margins_collapsed()
    {
        UiTest.Run(() =>
        {
            var controls = Render("# H\n---\n- a\n- b\n\nText\n# Next\n\nMore\n\n---\n\nEnd");

            Assert.Equal(
                new double[] { 0, 14, 3, 14, 21, 14, 21, 21 },
                controls.Select(control => control.Margin.Top).ToArray());
        });
    }

    // A list shows the blocks one by one, building each only as it comes into view - so the space
    // above each is decided for the whole run first, and a block built on its own still has it.
    [Fact]
    public void Each_block_carries_the_space_above_it_when_built_on_its_own()
    {
        var items = MarkdownRenderer.Arrange(GitHubMarkdown.Parse("# H\n---\n- a\n- b"));

        Assert.Equal(new double[] { 0, 14, 3 }, items.Select(item => item.Gap).ToArray());

        UiTest.Run(() =>
        {
            var renderer = new MarkdownRenderer(14, _ => { });

            Assert.Equal(3, renderer.Build(items[2]).Margin.Top);
            Assert.Equal(new Thickness(0, 14, 0, 0), renderer.Build(items[1]).Margin);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Text
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void Bold_italic_struck_and_code_pieces_are_runs_in_their_style()
    {
        UiTest.Run(() =>
        {
            var text = Assert.IsType<TextBlock>(Assert.Single(Render("**b** *i* ~~s~~ `c`")));
            Assert.Null(text.Text);

            var runs = text.Inlines!.OfType<Run>().ToList();

            Assert.Equal(FontWeight.SemiBold, runs.Single(run => run.Text == "b").FontWeight);
            Assert.Equal(FontStyle.Italic, runs.Single(run => run.Text == "i").FontStyle);
            Assert.Contains(runs.Single(run => run.Text == "s").TextDecorations!, decoration => decoration.Location == TextDecorationLocation.Strikethrough);

            // Code: a monospace font, a size smaller, padded inside its ground by a no-break space.
            var code = runs.Single(run => run.Text == "\u00A0c\u00A0");
            Assert.Contains("Consolas", code.FontFamily.ToString());
            Assert.Equal(14 * 0.85, code.FontSize, 3);
        });
    }

    // Most lines are just words - those stay plain Text, which lays out cheaper.
    [Fact]
    public void Plain_words_stay_plain_text()
    {
        UiTest.Run(() =>
        {
            var text = Assert.IsType<TextBlock>(Assert.Single(Render("plain words")));

            Assert.Equal("plain words", text.Text);
            Assert.True(text.Inlines == null || text.Inlines.Count == 0);
        });
    }

    // A click is looked up by where in the text it lands, so where each link lies is what matters.
    [Fact]
    public void A_link_is_underlined_and_found_by_where_it_lies_in_the_text()
    {
        UiTest.Run(() =>
        {
            var text = Assert.IsType<TextBlock>(Assert.Single(Render("Read [the docs](https://x.example/a) now")));

            var link = Assert.Single(MarkdownRenderer.LinksOf(text));
            Assert.Equal(new MarkdownLinkRange(5, 13, "https://x.example/a"), link);

            var links = MarkdownRenderer.LinksOf(text);
            Assert.Null(MarkdownRenderer.LinkAt(links, 4));
            Assert.Equal("https://x.example/a", MarkdownRenderer.LinkAt(links, 5));
            Assert.Equal("https://x.example/a", MarkdownRenderer.LinkAt(links, 12));
            Assert.Null(MarkdownRenderer.LinkAt(links, 13));

            var run = text.Inlines!.OfType<Run>().Single(r => r.Text == "the docs");
            Assert.Contains(run.TextDecorations!, decoration => decoration.Location == TextDecorationLocation.Underline);
        });
    }

    // A click on the link's words opens it, and only it - a click on the words beside it opens
    // nothing.
    [Fact]
    public void A_click_on_a_link_opens_its_address_and_a_click_beside_it_nothing()
    {
        UiTest.Run(() =>
        {
            var opened = new List<string>();
            var host = new StackPanel();
            var window = new Window { Content = host, Width = 600, Height = 200 };

            foreach (var control in new MarkdownRenderer(14, opened.Add).Render(GitHubMarkdown.Parse("Read [the docs](https://x.example/a) now")))
            {
                host.Children.Add(control);
            }

            window.Show();

            try
            {
                var text = (TextBlock)host.Children[0];

                Click(window, text, 1);
                Assert.Empty(opened);

                Click(window, text, 8);
                Assert.Equal(new[] { "https://x.example/a" }, opened);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A link after a code piece lies past the code's padding too - the click must still find it.
    [Fact]
    public void A_link_after_code_lies_where_the_padded_code_ends()
    {
        UiTest.Run(() =>
        {
            var text = Assert.IsType<TextBlock>(Assert.Single(Render("`ab` [x](https://x.example)")));

            // "\u00A0ab\u00A0" is 4 characters, then " ", then the link.
            Assert.Equal(new MarkdownLinkRange(5, 6, "https://x.example"), Assert.Single(MarkdownRenderer.LinksOf(text)));
        });
    }

    // "BUILDING.md" works on GitHub, relative to the repository; CRT has nothing it is relative to,
    // and ExternalTargetLauncher would look for it in the data folder - so it is no link here.
    [Fact]
    public void A_link_CRT_cannot_open_by_itself_is_plain_text()
    {
        UiTest.Run(() =>
        {
            var text = Assert.IsType<TextBlock>(Assert.Single(Render("[build notes](BUILDING.md)")));

            Assert.Equal("build notes", text.Text);
            Assert.Empty(MarkdownRenderer.LinksOf(text));
        });
    }

    [Theory]
    [InlineData("https://github.com/HovKlan-DH/Classic-Repair-Toolbox", true)]
    [InlineData("http://example.com", true)]
    [InlineData("mailto:anna@example.org", true)]
    [InlineData("BUILDING.md", false)]
    [InlineData("ftp://example.com/file", false)]
    [InlineData("file:///etc/passwd", false)]
    [InlineData("javascript:alert(1)", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void Only_an_absolute_web_or_mail_address_is_a_link(string? url, bool openable)
    {
        Assert.Equal(openable, MarkdownRenderer.IsOpenable(url));
    }

    // -----------------------------------------------------------------------------------------
    // Alerts, quotes and code blocks
    // -----------------------------------------------------------------------------------------

    // GitHub's alert: a bar down its left in the alert's color, then its icon and title over what
    // it says.
    [Fact]
    public void An_alert_has_a_colored_bar_with_its_icon_and_title_over_its_text()
    {
        UiTest.Run(() =>
        {
            var alert = Assert.IsType<Border>(Assert.Single(Render("> [!CAUTION]\n> Careful.")));
            Assert.Contains("ChangelogAlertCaution", alert.Classes);
            Assert.Equal(new Thickness(4, 0, 0, 0), alert.BorderThickness);

            var content = Assert.IsType<StackPanel>(alert.Child);
            var titleRow = Assert.IsType<StackPanel>(content.Children[0]);
            var icon = Assert.IsType<TextBlock>(titleRow.Children[0]);
            var title = Assert.IsType<TextBlock>(titleRow.Children[1]);

            Assert.Equal("\uF06A", icon.Text);
            Assert.Equal("Caution", title.Text);
            Assert.Equal("Careful.", Assert.IsType<TextBlock>(content.Children[1]).Text);
        });
    }

    [Fact]
    public void Every_alert_has_GitHubs_title()
    {
        Assert.Equal(
            new[] { "Note", "Tip", "Important", "Warning", "Caution" },
            Enum.GetValues<MarkdownAlertKind>().Select(MarkdownRenderer.AlertTitle).ToArray());

        Assert.Equal("Changelog_Alert_Caution", MarkdownRenderer.AlertKey(MarkdownAlertKind.Caution));
    }

    [Fact]
    public void A_plain_quote_has_a_bar_and_no_title()
    {
        UiTest.Run(() =>
        {
            var quote = Assert.IsType<Border>(Assert.Single(Render("> Said elsewhere.")));
            Assert.Contains("ChangelogQuote", quote.Classes);

            var content = Assert.IsType<StackPanel>(quote.Child);
            Assert.Equal("Said elsewhere.", Assert.IsType<TextBlock>(Assert.Single(content.Children)).Text);
        });
    }

    [Fact]
    public void A_code_block_is_monospace_on_a_ground_of_its_own()
    {
        UiTest.Run(() =>
        {
            var box = Assert.IsType<Border>(Assert.Single(Render("```\ndotnet test\n```")));
            var text = Assert.IsType<TextBlock>(box.Child);

            Assert.Equal("dotnet test", text.Text);
            Assert.Contains("Consolas", text.FontFamily.ToString());
        });
    }

    // -----------------------------------------------------------------------------------------
    // Colors
    // -----------------------------------------------------------------------------------------

    public static IEnumerable<object[]> ThemeKeys() => MarkdownRenderer.ThemeKeys.Select(key => new object[] { key });

    // App.axaml lists Light and Dark a few hundred lines apart; a key in only one leaves the other
    // theme drawing nothing there.
    [Theory]
    [MemberData(nameof(ThemeKeys))]
    public void Every_changelog_color_is_defined_in_both_themes(string key)
    {
        UiTest.Run(() =>
        {
            Assert.IsAssignableFrom<ISolidColorBrush>(Resolve(key, ThemeVariant.Light));
            Assert.IsAssignableFrom<ISolidColorBrush>(Resolve(key, ThemeVariant.Dark));
        });
    }

    [Theory]
    [MemberData(nameof(ThemeKeys))]
    public void Every_changelog_color_differs_between_light_and_dark(string key)
    {
        UiTest.Run(() =>
        {
            var light = (ISolidColorBrush)Resolve(key, ThemeVariant.Light)!;
            var dark = (ISolidColorBrush)Resolve(key, ThemeVariant.Dark)!;

            Assert.NotEqual(light.Color, dark.Color);
        });
    }

    // The colors are bound, not read once - shown, an alert's title has its theme's color.
    [Fact]
    public void A_shown_alert_takes_its_color_from_the_theme()
    {
        UiTest.Run(() =>
        {
            var host = new StackPanel();
            var window = new Window { Content = host, Width = 600, Height = 400 };
            var renderer = new MarkdownRenderer(14, _ => { });

            foreach (var control in renderer.Render(GitHubMarkdown.Parse("> [!TIP]\n> Try it.")))
            {
                host.Children.Add(control);
            }

            window.Show();

            try
            {
                var alert = (Border)host.Children[0];
                var title = (TextBlock)((StackPanel)((StackPanel)alert.Child!).Children[0]).Children[1];
                var expected = ThemeResources.ResolveColor(MarkdownRenderer.AlertKey(MarkdownAlertKind.Tip), Colors.Transparent);

                Assert.Equal(expected, Assert.IsAssignableFrom<ISolidColorBrush>(title.Foreground).Color);
                Assert.Equal(expected, Assert.IsAssignableFrom<ISolidColorBrush>(alert.BorderBrush).Color);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // -----------------------------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------------------------

    private static IReadOnlyList<Control> Render(string markdown) =>
        new MarkdownRenderer(14, _ => { }).Render(GitHubMarkdown.Parse(markdown));

    private static object? Resolve(string key, ThemeVariant variant) =>
        Application.Current!.TryGetResource(key, variant, out var resource) ? resource : null;

    // A left click in the middle of the character at index in the block's text.
    private static void Click(Window window, TextBlock text, int index)
    {
        var bounds = text.TextLayout.HitTestTextPosition(index);
        var point = text.TranslatePoint(bounds.Center, window)!.Value;

        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
    }
}
