using Avalonia.Controls;
using Avalonia.Media;
using Avalonia;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The two pieces of UI that surround "Add worklog" on the Schematics tab: the canvas that PARKS
// the pills of entries with no marked area, and the khaki mode hint telling the user to drag one
// out.
//
// What used to sit between them - a small "New fault" quick card asking for a title, description,
// category, state and the component checklist - is gone: drawing an area now opens the full
// WorklogEntryEditorWindow directly, so the fields those tests pinned down live in that window and
// are covered by its own tests. The two areas kept here are NOT part of that card and outlived it.
[Collection("HeadlessUi")]
public class WorklogEntryModeTests
{
    // ------------------------------------------------------------- parked-pill canvas

    // The pills of entries whose "Show marked area" is unticked are parked in the schematic
    // panel's top-right corner, on their OWN canvas - not the badge canvas beside it, which
    // carries the view matrix and would pan and zoom them with the board.
    [Fact]
    public void The_parked_pill_canvas_is_separate_from_the_anchored_badge_canvas()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            var parked = tab.FindControl<Canvas>("SchematicsWorklogParkedBadgeCanvas")!;
            var anchored = tab.FindControl<Canvas>("SchematicsWorklogEntriesBadgeCanvas")!;

            Assert.NotSame(anchored, parked);

            // The anchored canvas carries a transform it is given the view matrix through; the
            // parked one must not, or its pills would move with the board.
            Assert.NotNull(anchored.RenderTransform);
            Assert.Null(parked.RenderTransform);
        });
    }

    // In Avalonia a null Background is not hit-testable while Transparent IS - verified with a
    // probe, not assumed. A background here would make the canvas swallow every press across the
    // whole schematic panel and kill panning, so its absence is load-bearing rather than an
    // oversight waiting to be "fixed".
    [Fact]
    public void The_parked_pill_canvas_has_no_background_so_it_cannot_swallow_board_clicks()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            Assert.Null(tab.FindControl<Canvas>("SchematicsWorklogParkedBadgeCanvas")!.Background);
        });
    }

    // The pills park BELOW the "Netlist names" panel in z-order, since they step aside for it
    // rather than covering it.
    [Fact]
    public void The_parked_pills_sit_below_the_netlist_panel_in_z_order()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            var parked = tab.FindControl<Canvas>("SchematicsWorklogParkedBadgeCanvas")!;
            var panel = tab.FindControl<Border>("KiCadNetConnectionsPanel")!;

            Assert.True(parked.ZIndex < panel.ZIndex, $"parked pills at z{parked.ZIndex} are not below the panel at z{panel.ZIndex}");
        });
    }

    // ------------------------------------------------------------- the mode hint

    // Clicking "Add worklog" leaves the user with a crosshair and no instruction, so the mode
    // shows a khaki label in the tab-header row saying what to do with it. The wording is asserted
    // for its key phrases rather than verbatim, so rewording it is not a test failure.
    [Fact]
    public void The_worklog_mode_hint_says_to_mark_an_area_and_why()
    {
        Assert.Contains("mark an area", CRT.Main.WorklogAreaModeHint, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("components in scope", CRT.Main.WorklogAreaModeHint, StringComparison.OrdinalIgnoreCase);
    }

    // It reads as a full sentence, not a sentence fragment, so it ends with a period.
    [Fact]
    public void The_worklog_mode_hint_ends_with_a_period()
    {
        Assert.EndsWith(".", CRT.Main.WorklogAreaModeHint, StringComparison.Ordinal);
    }

    // The hint is a label over the tab-header row, not something the user has to interact with,
    // so it must never take a click - least of all the very click that dismisses it.
    //
    // Main itself is not constructible in a test (its constructor initialises every tab and reads
    // the real data root), so the markup is read from the compiled AXAML rather than from a live
    // window. That is enough to hold the two properties that matter here.
    [Fact]
    public void The_mode_hint_is_not_hit_testable_and_starts_hidden()
    {
        string markup = ReadMainWindowMarkup();

        int hintIndex = markup.IndexOf("x:Name=\"ModeHintBorder\"", StringComparison.Ordinal);
        Assert.True(hintIndex >= 0, "ModeHintBorder is not in Main.axaml");

        // The Panel's own attribute block, up to the end of its opening tag.
        string element = markup[hintIndex..markup.IndexOf('>', hintIndex)];

        Assert.Contains("IsVisible=\"False\"", element, StringComparison.Ordinal);
        Assert.Contains("IsHitTestVisible=\"False\"", element, StringComparison.Ordinal);
    }

    // The hint floats over the WHOLE window now - it is a direct RootGrid child pinned to the
    // upper-left corner, above the LeftPanel sidebar as well as the tab area - rather than being
    // confined to the row the data-sync icon sits in. It still carries a ZIndex well above every
    // other overlay in the window as a safety net, even though being RootGrid's last child already
    // guarantees it draws on top.
    [Fact]
    public void The_mode_hint_sits_above_the_data_sync_icon()
    {
        string markup = ReadMainWindowMarkup();

        int hintZ = ReadZIndex(markup, "ModeHintBorder");
        int syncZ = ReadZIndex(markup, "DataSyncStatusIconBorder");

        Assert.True(hintZ > syncZ, $"the hint at ZIndex {hintZ} does not draw above the sync icon at {syncZ}");
    }

    // The hint's text must WRAP inside its box rather than running past the window edge.
    //
    // This is not a style preference - it is a structural trap. A horizontal StackPanel measures
    // its children with INFINITE width, so a TextWrapping="Wrap" TextBlock inside one never wraps
    // however small the surrounding Border's MaxWidth is. Measured: inside a StackPanel the real
    // hint text laid out 1092px wide on a single line, overflowing a 560px box; as the Border's
    // direct child it lays out 540px across two lines.
    //
    // The layout is rebuilt here rather than read from Main.axaml, because the defect is in how
    // the controls MEASURE each other and only a real layout pass can show it.
    [Fact]
    public void The_mode_hint_text_wraps_inside_its_box_instead_of_overflowing()
    {
        UiTest.Run(() =>
        {
            const double maxWidth = 560;

            var text = new TextBlock
            {
                Text = CRT.Main.WorklogAreaModeHint,
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };

            var border = new Border
            {
                MaxWidth = maxWidth,
                Padding = new Thickness(10, 4),
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                Child = text,
            };

            var window = new Window { Width = 1200, Height = 300, Content = new Panel { Children = { border } } };

            try
            {
                window.Show();
                window.Measure(new Size(1200, 300));
                window.Arrange(new Rect(0, 0, 1200, 300));
                Dispatcher.UIThread.RunJobs();

                Assert.True(
                    text.Bounds.Width <= maxWidth,
                    $"the hint text laid out {text.Bounds.Width}px wide, past its {maxWidth}px box - it is not wrapping");

                // Two lines at this width, so the assertion above is not passing merely because the
                // text happened to be short.
                Assert.True(
                    text.Bounds.Height > 20,
                    $"the hint text is only {text.Bounds.Height}px tall - it did not wrap onto a second line, so this test is not exercising the case it claims");
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The hint carries the same Font Awesome "circle-question" glyph
    // (ConfigurationHelpIconTests.HelpGlyph, U+F059) the Configuration tab's help buttons use,
    // ahead of a bold "Hint:" lead-in on its OWN line, separated from the sentence below it by a
    // BLANK line (two LineBreaks, not one) for vertical breathing room. The icon is its own Run
    // with FontAwesomeRegular rather than folded into the "Hint:" Run's Text, because a single Run
    // can only carry one FontFamily and the regular UI font has no glyph at U+F059 - and the
    // sentence itself must still be a plain Run with no weight of its own, since only the lead-in
    // is bold.
    [Fact]
    public void The_mode_hint_has_an_icon_and_a_bold_lead_in_with_a_blank_line_after_it()
    {
        string markup = ReadMainWindowMarkup();

        int start = markup.IndexOf("x:Name=\"ModeHintBorder\"", StringComparison.Ordinal);
        Assert.True(start >= 0, "ModeHintBorder is not in Main.axaml");

        int end = markup.IndexOf("</Panel>", start, StringComparison.Ordinal);
        string block = markup[start..end];

        Assert.Contains("Text=\"&#xf059;\" FontFamily=\"{StaticResource FontAwesomeRegular}\"", block, StringComparison.Ordinal);
        Assert.Contains("Text=\"  Hint:\" FontWeight=\"Bold\" /><LineBreak /><LineBreak /><Run", block, StringComparison.Ordinal);

        int hintRunIndex = block.IndexOf("x:Name=\"ModeHintText\"", StringComparison.Ordinal);
        Assert.True(hintRunIndex >= 0, "ModeHintText is not in the mode-hint block");
        string hintRunElement = block[hintRunIndex..block.IndexOf('>', hintRunIndex)];
        Assert.DoesNotContain("FontWeight", hintRunElement, StringComparison.Ordinal);
    }

    // The hint's outline is DASHED rather than solid, so it reads as distinct from the app's
    // ordinary panel borders. Border has no dashed-edge option in Avalonia, so the outline is a
    // Rectangle with StrokeDashArray layered under the text - the same pattern
    // WorklogEntryEditorWindow.axaml uses for its drag placeholder. It is also IndianRed (via
    // Main_TabUnderline_Selected, the same accent the selected-tab underline and selected-workbook
    // card use) and 2px - one step heavier than the app's ordinary 1px borders - so it reads as
    // clearly distinct rather than just another panel outline.
    [Fact]
    public void The_mode_hint_border_is_dashed_and_indian_red_at_2px()
    {
        string markup = ReadMainWindowMarkup();

        int start = markup.IndexOf("x:Name=\"ModeHintBorder\"", StringComparison.Ordinal);
        Assert.True(start >= 0, "ModeHintBorder is not in Main.axaml");

        int end = markup.IndexOf("</Panel>", start, StringComparison.Ordinal);
        string block = markup[start..end];

        Assert.Contains("StrokeDashArray=", block, StringComparison.Ordinal);
        Assert.Contains("Stroke=\"{DynamicResource Main_TabUnderline_Selected}\"", block, StringComparison.Ordinal);
        Assert.Contains("StrokeThickness=\"2\"", block, StringComparison.Ordinal);
    }

    // It floats above the LEFT PANEL too, not just the tab area - a direct RootGrid child spanning
    // all three columns, rather than nested inside RightPanel where it used to live.
    [Fact]
    public void The_mode_hint_spans_the_whole_window_including_the_left_panel()
    {
        string markup = ReadMainWindowMarkup();

        int start = markup.IndexOf("x:Name=\"ModeHintBorder\"", StringComparison.Ordinal);
        Assert.True(start >= 0, "ModeHintBorder is not in Main.axaml");

        string element = markup[start..markup.IndexOf('>', start)];

        Assert.Contains("Grid.ColumnSpan=\"3\"", element, StringComparison.Ordinal);
        Assert.Contains("MaxWidth=\"350\"", element, StringComparison.Ordinal);
    }

    // 15px of padding inside the dashed box, on all sides - the text must never touch the border.
    [Fact]
    public void The_mode_hint_text_has_15px_padding()
    {
        string markup = ReadMainWindowMarkup();

        int start = markup.IndexOf("x:Name=\"ModeHintBorder\"", StringComparison.Ordinal);
        Assert.True(start >= 0, "ModeHintBorder is not in Main.axaml");

        int end = markup.IndexOf("</Panel>", start, StringComparison.Ordinal);
        string block = markup[start..end];

        int textBlockIndex = block.IndexOf("<TextBlock", StringComparison.Ordinal);
        Assert.True(textBlockIndex >= 0, "no TextBlock in the mode-hint block");
        string textBlockElement = block[textBlockIndex..block.IndexOf('>', textBlockIndex)];

        Assert.Contains("Margin=\"15\"", textBlockElement, StringComparison.Ordinal);
    }

    private static int ReadZIndex(string markup, string controlName)
    {
        int index = markup.IndexOf($"x:Name=\"{controlName}\"", StringComparison.Ordinal);
        Assert.True(index >= 0, $"{controlName} is not in Main.axaml");

        string element = markup[index..markup.IndexOf('>', index)];

        var match = System.Text.RegularExpressions.Regex.Match(element, @"ZIndex=""(\d+)""");
        Assert.True(match.Success, $"{controlName} declares no ZIndex");

        return int.Parse(match.Groups[1].Value);
    }

    private static string ReadMainWindowMarkup()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory != null)
        {
            string candidate = Path.Combine(directory.FullName, "Main", "Main.axaml");
            if (File.Exists(candidate))
            {
                return File.ReadAllText(candidate);
            }
            directory = directory.Parent;
        }

        throw new FileNotFoundException($"Could not find Main/Main.axaml above {AppContext.BaseDirectory}");
    }

}
