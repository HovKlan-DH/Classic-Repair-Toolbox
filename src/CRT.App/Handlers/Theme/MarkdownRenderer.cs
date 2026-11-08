using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Theming
{
    // ###########################################################################################
    // DRAWS WHAT GitHubMarkdown READ, THE WAY GITHUB DRAWS IT (owner request, 2026-11-06: the
    // changelog on the About tab, "style it alike GitHub, so it gets indented and bold/italic etc.").
    //
    // GitHub's own sizes, read off its stylesheet and scaled from its 16px text to CRT's 14px - but for
    // the headings, which are the text's own size and bold (owner request, 2026-11-06), the first two with
    // a thin line under them, and a thematic break right under a heading left out, the heading's line
    // being the divider; any other thematic break a 4px bar; text on lines 1.5 high; a list indented 2em a
    // level, its bullets a disc, a circle and then a square; inline code in a monospace font on a
    // faint ground; an alert a colored bar down its left with its icon and title in that color.
    // Space between blocks is GitHub's margins collapsed as a browser collapses them - the larger of
    // the one above's bottom and the one below's top.
    //
    // Every color is a theme resource (Changelog_* in App.axaml, Link_Fg for links), BOUND rather than
    // read once, so a theme switched in the Configuration tab recolors the changelog already shown -
    // each to the element's OWN resources, never a long-lived host's: the About tab builds a block as
    // it scrolls into view and lets it go again (Arrange/Build), and a binding to the tab would keep
    // every block ever built alive with it.
    //
    // A link opens through the action handed in - the About tab hands ExternalTargetLauncher's, the
    // only sanctioned way to open one - and only an absolute http, https or mailto address is a link:
    // a relative one ("BUILDING.md", fine on GitHub) has nothing here to be relative to, and is shown
    // as plain text.
    // ###########################################################################################
    internal sealed class MarkdownRenderer
    {
        public const string CodeBackgroundKey = "Changelog_Code_Bg";
        public const string RuleKey = "Changelog_Rule";
        public const string HeadingLineKey = "Changelog_HeadingLine";
        public const string QuoteBarKey = "Changelog_Quote_Bar";
        public const string QuoteTextKey = "Changelog_Quote_Fg";
        public const string LinkKey = "Link_Fg";

        // Every key the changelog draws with - MarkdownRendererTests holds each to both themes.
        public static readonly IReadOnlyList<string> ThemeKeys = new[]
        {
            CodeBackgroundKey, RuleKey, HeadingLineKey, QuoteBarKey, QuoteTextKey, LinkKey,
        }.Concat(Enum.GetValues<MarkdownAlertKind>().Select(AlertKey)).ToList();

        // A list's bullets by depth, as a browser draws nested lists: disc, circle, then square for
        // every level below that. Inter, CRT's font, has all three, so they sit on the text's own
        // baseline.
        public static readonly IReadOnlyList<string> Bullets = new[] { "\u2022", "\u25E6", "\u25AA" };

        private const string MonospaceFamily = "Cascadia Mono,Consolas,Menlo,monospace";

        // GitHub's px at 16px text, times 14/16.
        private const double ListIndent = 28;
        private const double BulletGap = 6;
        private const double ItemGap = 3;
        private const double BlockBottom = 14;
        private const double HeadingTop = 21;
        private const double RuleMargin = 21;
        private const double RuleHeight = 4;
        private const double QuoteBarWidth = 4;
        private const double AlertTitleGap = 10;

        private static readonly Cursor HandCursor = new(StandardCursorType.Hand);

        // The links of a text, by where they lie in it - what a click is looked up in.
        private static readonly AttachedProperty<IReadOnlyList<MarkdownLinkRange>?> LinksProperty =
            AvaloniaProperty.RegisterAttached<TextBlock, IReadOnlyList<MarkdownLinkRange>?>("MarkdownLinks", typeof(MarkdownRenderer));

        private readonly double thisFontSize;
        private readonly Action<string> thisOpenLink;

        // fontSize: the text's size, everything else is scaled from it. openLink: what a click on a
        // link does.
        public MarkdownRenderer(double fontSize, Action<string> openLink)
        {
            this.thisFontSize = fontSize > 0 ? fontSize : 14;
            this.thisOpenLink = openLink;
        }

        public static string AlertKey(MarkdownAlertKind kind) => $"Changelog_Alert_{kind}";

        // GitHub's title for each alert, and the Font Awesome glyph closest to its octicon (all five in
        // the Solid face, none drawn past the font's ascent - FontAwesomeAssetTests).
        public static string AlertTitle(MarkdownAlertKind kind) => kind.ToString();

        public static string AlertIcon(MarkdownAlertKind kind) => kind switch
        {
            MarkdownAlertKind.Note => "\uF05A",      // circle-info
            MarkdownAlertKind.Tip => "\uF0EB",       // lightbulb
            MarkdownAlertKind.Important => "\uF075", // comment
            MarkdownAlertKind.Warning => "\uF071",   // triangle-exclamation
            _ => "\uF06A",                           // circle-exclamation, for Caution
        };

        // The links a text block shows, by where they lie in its text - empty for one with none.
        public static IReadOnlyList<MarkdownLinkRange> LinksOf(TextBlock block) =>
            block.GetValue(LinksProperty) ?? Array.Empty<MarkdownLinkRange>();

        // The link at a place in a text, or null.
        public static string? LinkAt(IReadOnlyList<MarkdownLinkRange> links, int index)
        {
            foreach (var link in links)
            {
                if (index >= link.Start && index < link.End)
                {
                    return link.Url;
                }
            }

            return null;
        }

        // Only an address CRT can open on its own is a link: http, https or mailto, absolute.
        public static bool IsOpenable(string? url) =>
            !string.IsNullOrWhiteSpace(url) &&
            Uri.TryCreate(url.Trim(), UriKind.Absolute, out Uri? uri) &&
            (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeMailto);

        // ###########################################################################################
        // The blocks with the space above each worked out - what a list shows one by one, building
        // each with Build only as it comes into view. The space depends on the block above, so it is
        // decided here, once, for the whole run of them.
        // ###########################################################################################
        public static IReadOnlyList<MarkdownBlockItem> Arrange(IReadOnlyList<MarkdownBlock> blocks)
        {
            var items = new List<MarkdownBlockItem>(blocks.Count);
            MarkdownBlock? previous = null;

            foreach (var block in blocks)
            {
                // A heading's line is its divider: the rule the changelog puts right under each version is
                // a second, thicker one, and is left out (owner request, 2026-11-06: "there are two ...
                // dividers showing - just keep the first one (the thinnest one)").
                if (block is MarkdownRule && previous is MarkdownHeading)
                {
                    continue;
                }

                items.Add(new MarkdownBlockItem(block, Gap(previous, block)));
                previous = block;
            }

            return items;
        }

        // One block's control, the space above it as its top margin.
        public Control Build(MarkdownBlockItem item)
        {
            Control control = this.RenderBlock(item.Block);
            var margin = control.Margin;
            control.Margin = new Thickness(margin.Left, item.Gap, margin.Right, margin.Bottom);
            return control;
        }

        // Every block's control at once, in order - for blocks few enough to build together, as an
        // alert's inside.
        public IReadOnlyList<Control> Render(IReadOnlyList<MarkdownBlock> blocks) =>
            Arrange(blocks).Select(this.Build).ToList();

        // The space above a block: GitHub's margins, collapsed - the larger of the block above's
        // bottom margin and this one's top.
        internal static double Gap(MarkdownBlock? previous, MarkdownBlock next)
        {
            if (previous == null)
            {
                return 0;
            }

            if (previous is MarkdownListItem && next is MarkdownListItem)
            {
                return ItemGap;
            }

            double bottom = previous switch
            {
                MarkdownRule => RuleMargin,
                _ => BlockBottom,
            };

            double top = next switch
            {
                MarkdownHeading => HeadingTop,
                MarkdownRule => RuleMargin,
                _ => 0,
            };

            return Math.Max(bottom, top);
        }

        private Control RenderBlock(MarkdownBlock block) => block switch
        {
            MarkdownHeading heading => this.RenderHeading(heading),
            MarkdownRule => this.RenderRule(),
            MarkdownListItem item => this.RenderListItem(item),
            MarkdownParagraph paragraph => Indented(this.Text(paragraph.Spans, this.thisFontSize, 1.5), paragraph.Indent),
            MarkdownQuote quote => this.RenderQuote(quote),
            MarkdownCode code => this.RenderCode(code),
            _ => new Panel(),
        };

        // ###########################################################################################
        // A heading at the text's own size, bold (owner request, 2026-11-06: the version and date "way too
        // large. Please have that normal font size and then just have it bold") - GitHub's 2x and 1.5x went.
        // One BoldAsMarked is plain, its bold pieces bold: a version's, "Version **3.0.0**, 2026-October-11"
        // (ChangelogVersionHeadings). The first two levels keep their thin line under them.
        // ###########################################################################################
        private Control RenderHeading(MarkdownHeading heading)
        {
            int level = Math.Clamp(heading.Level, 1, 6);
            double size = this.thisFontSize;

            var text = this.Text(heading.Spans, size, 1.5, FontWeight.Bold);
            text.FontWeight = heading.BoldAsMarked ? FontWeight.Normal : FontWeight.Bold;
            text.Classes.Add("ChangelogHeading");

            if (level == 6)
            {
                BindBrush(text, TextBlock.ForegroundProperty, QuoteTextKey);
            }

            if (level > 2)
            {
                return text;
            }

            var underlined = new Border
            {
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(0, 0, 0, Math.Round(size * 0.3)),
                Child = text,
            };

            BindBrush(underlined, Border.BorderBrushProperty, HeadingLineKey);
            return underlined;
        }

        private Control RenderRule()
        {
            var rule = new Border { Height = RuleHeight };
            BindBrush(rule, Border.BackgroundProperty, RuleKey);
            return rule;
        }

        // The bullet in a column of the list's indent, right against the text; the text wraps under
        // itself, never under the bullet.
        private Control RenderListItem(MarkdownListItem item)
        {
            var grid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions($"{ListIndent},*"),
                Margin = new Thickness(item.Depth * ListIndent, 0, 0, 0),
            };

            // The same size and line height as the text beside it, so the two first lines sit on one
            // baseline.
            var bullet = new TextBlock
            {
                Text = item.Number is int number ? $"{number}." : Bullets[Math.Min(item.Depth, Bullets.Count - 1)],
                FontSize = this.thisFontSize,
                LineHeight = Math.Round(this.thisFontSize * 1.5),
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 0, BulletGap, 0),
            };

            bullet.Classes.Add("ChangelogBullet");

            var text = this.Text(item.Spans, this.thisFontSize, 1.5);
            Grid.SetColumn(text, 1);

            grid.Children.Add(bullet);
            grid.Children.Add(text);
            return grid;
        }

        // A quote: a bar down its left, its text muted. An alert: the bar in the alert's color, and
        // its icon and title over what it says.
        private Control RenderQuote(MarkdownQuote quote)
        {
            var content = new StackPanel();

            if (quote.Alert is MarkdownAlertKind kind)
            {
                var icon = new TextBlock
                {
                    Text = AlertIcon(kind),
                    FontFamily = ThemeResources.ResolveFontAwesomeSolid(),
                    FontSize = this.thisFontSize,
                    VerticalAlignment = VerticalAlignment.Center,
                    // The lightbulb reaches a little below the font's descent.
                    Padding = new Thickness(0, 0, 0, 1),
                };

                var title = new TextBlock
                {
                    Text = AlertTitle(kind),
                    FontSize = this.thisFontSize,
                    FontWeight = FontWeight.Medium,
                    VerticalAlignment = VerticalAlignment.Center,
                };

                title.Classes.Add("ChangelogAlertTitle");
                BindBrush(icon, TextBlock.ForegroundProperty, AlertKey(kind));
                BindBrush(title, TextBlock.ForegroundProperty, AlertKey(kind));

                content.Children.Add(new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Children = { icon, title },
                });
            }
            else
            {
                BindBrush(content, TextElement.ForegroundProperty, QuoteTextKey);
            }

            var inner = this.Render(quote.Blocks);
            for (int i = 0; i < inner.Count; i++)
            {
                if (i == 0 && quote.Alert != null)
                {
                    var margin = inner[i].Margin;
                    inner[i].Margin = new Thickness(margin.Left, AlertTitleGap, margin.Right, margin.Bottom);
                }

                content.Children.Add(inner[i]);
            }

            var bar = new Border
            {
                BorderThickness = new Thickness(QuoteBarWidth, 0, 0, 0),
                Padding = quote.Alert != null ? new Thickness(14, 7) : new Thickness(14, 0),
                Child = content,
            };

            bar.Classes.Add(quote.Alert is MarkdownAlertKind alert ? $"ChangelogAlert{alert}" : "ChangelogQuote");
            BindBrush(bar, Border.BorderBrushProperty, quote.Alert is MarkdownAlertKind barKind ? AlertKey(barKind) : QuoteBarKey);

            return Indented(bar, quote.Indent);
        }

        private Control RenderCode(MarkdownCode code)
        {
            var text = new TextBlock
            {
                Text = code.Text,
                FontFamily = new FontFamily(MonospaceFamily),
                FontSize = this.thisFontSize * 0.85,
                LineHeight = this.thisFontSize * 0.85 * 1.45,
                TextWrapping = TextWrapping.Wrap,
            };

            var box = new Border
            {
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(14),
                Child = text,
            };

            BindBrush(box, Border.BackgroundProperty, CodeBackgroundKey);
            return Indented(box, code.Indent);
        }

        private static Control Indented(Control control, int indent)
        {
            control.Margin = new Thickness(indent * ListIndent, 0, 0, 0);
            return control;
        }

        // ###########################################################################################
        // A text of styled pieces. With nothing styled it is plain Text, which lays out cheaper - most
        // of the changelog's lines are just that.
        // ###########################################################################################
        // boldWeight: how heavy a bold piece is - SemiBold in running text, as GitHub's 600; Bold in a heading.
        private TextBlock Text(IReadOnlyList<MarkdownSpan> spans, double fontSize, double lineHeight, FontWeight? boldWeight = null)
        {
            var block = new TextBlock
            {
                FontSize = fontSize,
                LineHeight = Math.Round(fontSize * lineHeight),
                TextWrapping = TextWrapping.Wrap,
            };

            if (spans.All(span => span.Style == MarkdownStyle.None && !IsOpenable(span.Url)))
            {
                block.Text = string.Concat(spans.Select(span => span.Text));
                return block;
            }

            block.Inlines ??= new InlineCollection();
            var links = new List<MarkdownLinkRange>();
            int position = 0;

            foreach (var span in spans)
            {
                var run = this.RunOf(span, fontSize, boldWeight ?? FontWeight.SemiBold);
                block.Inlines.Add(run);

                int length = run.Text?.Length ?? 0;
                if (IsOpenable(span.Url))
                {
                    links.Add(new MarkdownLinkRange(position, position + length, span.Url!.Trim()));
                }

                position += length;
            }

            if (links.Count > 0)
            {
                this.WireLinks(block, links);
            }

            return block;
        }

        private Run RunOf(MarkdownSpan span, double fontSize, FontWeight boldWeight)
        {
            var style = span.Style;
            bool isCode = (style & MarkdownStyle.Code) != 0;

            // Code is set off by a no-break space each side, inside its ground - GitHub pads it, and
            // the padding never wraps away from the code.
            var run = new Run(isCode ? $"\u00A0{span.Text}\u00A0" : span.Text);

            if ((style & MarkdownStyle.Bold) != 0)
            {
                run.FontWeight = boldWeight;
            }

            if ((style & MarkdownStyle.Italic) != 0)
            {
                run.FontStyle = FontStyle.Italic;
            }

            if (isCode)
            {
                run.FontFamily = new FontFamily(MonospaceFamily);
                run.FontSize = fontSize * 0.85;
                BindBrush(run, TextElement.BackgroundProperty, CodeBackgroundKey);
            }

            bool isLink = IsOpenable(span.Url);
            bool isStruck = (style & MarkdownStyle.Strikethrough) != 0;

            if (isLink)
            {
                BindBrush(run, TextElement.ForegroundProperty, LinkKey);
            }

            if (isLink || isStruck)
            {
                var decorations = new TextDecorationCollection();
                if (isLink)
                {
                    decorations.Add(new TextDecoration { Location = TextDecorationLocation.Underline });
                }

                if (isStruck)
                {
                    decorations.Add(new TextDecoration { Location = TextDecorationLocation.Strikethrough });
                }

                run.TextDecorations = decorations;
            }

            return run;
        }

        // Bound to the element's own resources - a run's are its text block's - so it follows the
        // theme of wherever it is shown, and holds on to nothing outside itself.
        private static void BindBrush(StyledElement target, AvaloniaProperty property, string key)
        {
            target.Bind(property, target.GetResourceObservable(key));
        }

        // ###########################################################################################
        // A click on a link opens it; the pointer is a hand only over one. A text block has no
        // pointer events per run, so the character under the pointer is looked up - by
        // TextLinkRenderer's rule for which line the pointer is on, as a wrapped line's own test is
        // wrong past the first.
        // ###########################################################################################
        private void WireLinks(TextBlock block, IReadOnlyList<MarkdownLinkRange> links)
        {
            block.SetValue(LinksProperty, links);

            // A text block with no background is not hit-testable.
            block.Background ??= Brushes.Transparent;

            block.PointerPressed += (_, e) =>
            {
                if (!e.GetCurrentPoint(block).Properties.IsLeftButtonPressed)
                {
                    return;
                }

                string? url = UrlUnder(block, e.GetPosition(block));
                if (url != null)
                {
                    e.Handled = true;
                    this.thisOpenLink(url);
                }
            };

            block.PointerMoved += (_, e) =>
            {
                Cursor? wanted = UrlUnder(block, e.GetPosition(block)) != null ? HandCursor : null;
                if (!ReferenceEquals(block.Cursor, wanted))
                {
                    block.Cursor = wanted;
                }
            };
        }

        private static string? UrlUnder(TextBlock block, Point point)
        {
            var layout = block.TextLayout;
            if (layout == null || !TextLinkRenderer.IsPointOverAnyLine(layout, point))
            {
                return null;
            }

            return LinkAt(LinksOf(block), layout.HitTestPoint(point).TextPosition);
        }
    }

    // Where a link lies in a text block's text, End past its last character.
    internal sealed record MarkdownLinkRange(int Start, int End, string Url);

    // A block as a list shows it: the block, and the space above it (MarkdownRenderer.Arrange).
    internal sealed record MarkdownBlockItem(MarkdownBlock Block, double Gap);
}
