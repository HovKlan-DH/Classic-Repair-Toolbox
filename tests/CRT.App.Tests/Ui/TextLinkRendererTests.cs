using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using Avalonia.Input;
using Handlers.DataHandling;
using Handlers.Theming;
using System.Reflection;

namespace ClassicRepairToolbox.Tests.Ui;

// Turning a user-typed repair note into a TextBlock whose links are clickable - the workbook Note,
// the worklog Description, and the Work done / Comment / Photo comment / File comment rows.
//
// WHICH runs are links is TextLinkFinder's job and is tested directly, on the Handlers side. What
// is tested here is the RENDERING, and above all the two ways it can lose text on screen:
//
//   1. A TextBlock carrying BOTH Text and Inlines renders the Text and silently ignores the
//      Inlines. So a linked block must have Text == null, and an unlinked one must have Inlines
//      empty - getting either backwards shows the wrong thing with no error.
//   2. The runs must concatenate back to the original string. Links and search highlighting cut
//      the text at INDEPENDENT places (a search term routinely lands inside a URL), so the two
//      splits are merged rather than applied one after the other; an off-by-one there drops or
//      doubles characters in a field the user typed by hand.
[Collection("HeadlessUi")]
public class TextLinkRendererTests
{
    // What the block actually shows, however it is carrying it. A highlighted or linked block has
    // Text == null and its content in Inlines, so a reader that only looks at Text sees it as blank.
    private static string VisibleText(TextBlock block) =>
        block.Text ?? string.Concat(block.Inlines?.OfType<Run>().Select(run => run.Text) ?? Array.Empty<string>());

    private static void WithBlock(string? text, Action<TextBlock> body)
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();
            TextLinkRenderer.Apply(block, text);
            body(block);
        });
    }

    private static void WithSegments(
        string text,
        string searchQuery,
        Action<TextBlock> body)
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();
            var query = WorklogSearchQuery.Parse(searchQuery);
            TextLinkRenderer.ApplySegments(block, text, query.SplitIntoSegments(text));
            body(block);
        });
    }

    // The overwhelmingly common case - a note with no URL in it. It stays a plain single-Text
    // TextBlock, which is measurably cheaper to lay out than a split into Inlines, and it must NOT
    // pick up a Hand cursor promising a click that does nothing.
    [Fact]
    public void Text_with_no_link_stays_a_plain_text_block()
    {
        WithBlock("Replaced U7 and reflowed CN2.", block =>
        {
            Assert.Equal("Replaced U7 and reflowed CN2.", block.Text);
            Assert.True(block.Inlines == null || block.Inlines.Count == 0);
            Assert.Null(block.Cursor);
        });
    }

    [Fact]
    public void Blank_text_renders_as_empty_rather_than_throwing()
    {
        WithBlock(null, block => Assert.Equal(string.Empty, VisibleText(block)));
        WithBlock(string.Empty, block => Assert.Equal(string.Empty, VisibleText(block)));
    }

    // A linked block MUST clear Text: a TextBlock holding both renders the Text and ignores the
    // Inlines entirely, so leaving it set would show the string with no link styling at all and
    // nothing would report an error.
    [Fact]
    public void A_block_with_a_link_moves_its_content_into_inlines()
    {
        WithBlock("See https://example.com for more", block =>
        {
            Assert.Null(block.Text);
            Assert.NotNull(block.Inlines);
            Assert.Equal("See https://example.com for more", VisibleText(block));
        });
    }

    // The link run is underlined and takes the themed link colour; the prose around it is left
    // alone, so the block does not read as one long link.
    [Fact]
    public void Only_the_link_run_is_underlined()
    {
        WithBlock("See https://example.com for more", block =>
        {
            var runs = block.Inlines!.OfType<Run>().ToList();

            var linked = runs.Where(r => r.TextDecorations != null).ToList();
            var plain = runs.Where(r => r.TextDecorations == null).ToList();

            Assert.Equal("https://example.com", Assert.Single(linked).Text);
            Assert.NotEmpty(plain);
            Assert.Equal("See  for more", string.Concat(plain.Select(r => r.Text)));
        });
    }

    [Fact]
    public void Several_links_in_one_text_are_all_marked_and_the_text_survives()
    {
        const string text = "Refs https://a.example and www.b.example done";

        WithBlock(text, block =>
        {
            Assert.Equal(2, block.Inlines!.OfType<Run>().Count(r => r.TextDecorations != null));
            Assert.Equal(text, VisibleText(block));
        });
    }

    // Re-rendering the same block - which happens on every refresh, and on every container a
    // template recycles - must not leave the previous pass's runs underneath the new value.
    [Fact]
    public void Re_rendering_a_block_replaces_its_previous_content()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();

            TextLinkRenderer.Apply(block, "See https://example.com now");
            Assert.Equal("See https://example.com now", VisibleText(block));

            TextLinkRenderer.Apply(block, "Plain text now");

            Assert.Equal("Plain text now", VisibleText(block));
            Assert.Equal("Plain text now", block.Text);
            Assert.True(block.Inlines == null || block.Inlines.Count == 0);

            // A block that has gone from linked to unlinked must lose its Hand cursor too, or it
            // keeps promising a click on prose that is no longer clickable.
            Assert.Null(block.Cursor);
        });
    }

    // ------------------------------------------------------------------------- cursor handling

    // Several of these blocks sit inside a row that sets its OWN cursor - the board pane's pills
    // carry a Hand, the editor's photo and file rows a resize cursor on their drag handle. Clearing
    // the cursor to null when the block stops being linked (or when the pointer moves off a link
    // run) silently drops that row's cursor the first time its prose is hovered, and nothing on
    // screen says why the handle stopped indicating it can be dragged. So the cursor the block
    // arrived with is recorded and put back, rather than being assumed to have been null.
    [Fact]
    public void A_blocks_own_cursor_survives_becoming_and_ceasing_to_be_linked()
    {
        UiTest.Run(() =>
        {
            var rowCursor = new Cursor(StandardCursorType.SizeNorthSouth);
            var block = new TextBlock { Cursor = rowCursor };

            TextLinkRenderer.Apply(block, "See https://example.com now");

            // Going back to link-free prose must restore the row's cursor, not null it.
            TextLinkRenderer.Apply(block, "No links here at all");

            Assert.Same(rowCursor, block.Cursor);
        });
    }

    // The same rule on the pointer-move path: moving off a link run inside a block that HAS a link
    // restores the block's own cursor rather than clearing it. Driven through the real handler, so
    // this fails against the version that assigned null.
    [Fact]
    public void Moving_off_a_link_restores_the_blocks_own_cursor()
    {
        UiTest.Run(() =>
        {
            var rowCursor = new Cursor(StandardCursorType.SizeNorthSouth);
            var block = new TextBlock { Cursor = rowCursor };

            TextLinkRenderer.Apply(block, "See https://example.com now");

            // The block is linked, so the Hand mechanism is live - but the pointer is not over a
            // link run here (nothing has been laid out, so no index resolves), which is exactly the
            // "off a link" branch.
            InvokePointerMoved(block, new Point(0, 0));

            Assert.Same(rowCursor, block.Cursor);
        });
    }

    // A block that never set a cursor of its own must still end up with none - the restore must not
    // invent one, and re-rendering an already-linked block must not capture the Hand this class
    // itself put there as that block's "original".
    [Fact]
    public void A_block_with_no_cursor_of_its_own_ends_up_with_none()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();

            TextLinkRenderer.Apply(block, "See https://example.com now");
            TextLinkRenderer.Apply(block, "Still https://example.com linked");
            TextLinkRenderer.Apply(block, "Plain prose now");

            Assert.Null(block.Cursor);
        });
    }

    // Calls the private PointerMoved handler directly. The alternative is a real pointer device on
    // a laid-out window, which is a great deal of machinery for a branch that is one assignment.
    private static void InvokePointerMoved(TextBlock block, Point point)
    {
        var method = typeof(TextLinkRenderer).GetMethod(
            "OnBlockPointerMoved", BindingFlags.NonPublic | BindingFlags.Static)!;

        var args = new Avalonia.Input.PointerEventArgs(
            InputElement.PointerMovedEvent,
            block,
            new Avalonia.Input.Pointer(0, Avalonia.Input.PointerType.Mouse, isPrimary: true),
            block,
            point,
            0,
            new Avalonia.Input.PointerPointProperties(
                Avalonia.Input.RawInputModifiers.None, Avalonia.Input.PointerUpdateKind.Other),
            Avalonia.Input.KeyModifiers.None);

        method.Invoke(null, new object?[] { block, args });
    }

    // ------------------------------------------------------------- links + search highlighting

    // The merge walks the segments by LENGTH against the text's own offsets, so it assumes they
    // concatenate back to exactly that text. Every caller passes SplitIntoSegments of the same
    // string, so it holds - but if it ever did not, the segment cursor would drift and the
    // highlight wash would land on characters that never matched, silently: the runs still rebuild
    // the string and nothing throws. So the invariant is CHECKED, and a mismatch degrades to
    // links-only rather than to wrong highlighting.
    [Fact]
    public void Segments_that_do_not_match_the_text_fall_back_to_no_highlighting()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();

            // Deliberately segments of a DIFFERENT string - what a caller passing the wrong text
            // (or a query whose splitting normalised something) would hand over.
            var wrongSegments = new List<(string Text, bool IsMatch)>
            {
                ("Replaced ", false),
                ("U7", true),
                (" today", false)
            };

            TextLinkRenderer.ApplySegments(block, "See https://example.com now", wrongSegments);

            // The text is still whole and correct - the non-negotiable part.
            Assert.Equal("See https://example.com now", VisibleText(block));

            // And nothing is washed as a search hit, rather than an arbitrary run being washed.
            Assert.DoesNotContain(block.Inlines!.OfType<Run>(), run => run.Background != null);

            // The link marking survives: it is derived from the text itself, not from the segments.
            Assert.Contains(block.Inlines!.OfType<Run>(), run => run.TextDecorations != null);
        });
    }

    // Segments that are a PREFIX of the text - the same length check from the other side, since a
    // short list leaves the tail unaccounted for rather than mismatching a character.
    [Fact]
    public void Segments_that_stop_short_of_the_text_fall_back_to_no_highlighting()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();

            var shortSegments = new List<(string Text, bool IsMatch)>
            {
                ("Replaced ", false),
                ("U7", true)
            };

            TextLinkRenderer.ApplySegments(block, "Replaced U7 and reflowed CN2.", shortSegments);

            Assert.Equal("Replaced U7 and reflowed CN2.", VisibleText(block));
            Assert.True(block.Inlines == null || block.Inlines.Count == 0);
            Assert.Equal("Replaced U7 and reflowed CN2.", block.Text);
        });
    }

    // Both markings on one block, cut at DIFFERENT places. This is the case that made the two
    // splits have to be merged rather than applied in sequence: "example" is inside the URL, so a
    // naive second pass would have to re-split a run that was already split.
    [Fact]
    public void A_search_hit_inside_a_link_keeps_the_whole_text_intact()
    {
        const string text = "See https://example.com for more";

        WithSegments(text, "example", block =>
        {
            Assert.Equal(text, VisibleText(block));

            // Both markings are present: something is underlined, and something is washed.
            var runs = block.Inlines!.OfType<Run>().ToList();
            Assert.Contains(runs, r => r.TextDecorations != null);
            Assert.Contains(runs, r => r.Background != null);
        });
    }

    // A search hit OUTSIDE the link marks the prose and leaves the link's own styling alone.
    [Fact]
    public void A_search_hit_outside_a_link_marks_only_the_prose()
    {
        const string text = "Replaced the cap, see https://example.com";

        WithSegments(text, "cap", block =>
        {
            Assert.Equal(text, VisibleText(block));

            var washed = block.Inlines!.OfType<Run>().Where(r => r.Background != null).ToList();
            Assert.Equal("cap", Assert.Single(washed).Text);
        });
    }

    // Highlighting still works on text with no link in it at all - the merge must not need a link
    // to be present in order to emit the search runs.
    [Fact]
    public void Search_highlighting_works_with_no_link_present()
    {
        const string text = "Replaced the cap on the board";

        WithSegments(text, "cap", block =>
        {
            Assert.Equal(text, VisibleText(block));

            var washed = block.Inlines!.OfType<Run>().Where(r => r.Background != null).ToList();
            Assert.Equal("cap", Assert.Single(washed).Text);
        });
    }

    // An empty query is not a filter and marks nothing, so the block takes the cheap plain path.
    [Fact]
    public void An_empty_query_leaves_plain_text_plain()
    {
        WithSegments("Replaced the cap", string.Empty, block =>
        {
            Assert.Equal("Replaced the cap", block.Text);
            Assert.True(block.Inlines == null || block.Inlines.Count == 0);
        });
    }

    // Every run must be non-empty. An empty Run renders as nothing but still costs a layout pass,
    // and one appearing is the signature of a boundary computed one character out.
    [Theory]
    [InlineData("https://example.com is the source", "example")]
    [InlineData("See https://example.com", "https")]
    [InlineData("www.example.com", "www")]
    [InlineData("a https://example.com b", "a")]
    public void The_merged_runs_are_never_empty_and_always_rebuild_the_text(string text, string query)
    {
        WithSegments(text, query, block =>
        {
            Assert.Equal(text, VisibleText(block));

            foreach (var run in block.Inlines!.OfType<Run>())
            {
                Assert.False(string.IsNullOrEmpty(run.Text), "an empty run was emitted - a boundary is off by one");
            }
        });
    }

    // ------------------------------------------------------------- the LinkText attached property

    // The markup-facing form the worklog editor's DataTemplates use, since a templated row has no
    // code-behind moment holding both the block and its text. Setting it must render exactly as
    // Apply does - not set Text alongside the Inlines, which would silently hide them.
    [Fact]
    public void Setting_LinkText_renders_the_same_as_Apply()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();
            TextLinkRenderer.SetLinkText(block, "See https://example.com now");

            Assert.Null(block.Text);
            Assert.Equal("See https://example.com now", VisibleText(block));
            Assert.Contains(block.Inlines!.OfType<Run>(), r => r.TextDecorations != null);
        });
    }

    // A template recycles its containers, so the same block is handed a different row's text as the
    // list is rebuilt. The property has to re-render on every change, not only the first.
    [Fact]
    public void Changing_LinkText_re_renders_the_block()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();

            TextLinkRenderer.SetLinkText(block, "See https://example.com");
            Assert.Equal("See https://example.com", VisibleText(block));

            TextLinkRenderer.SetLinkText(block, "A different row with no link");

            Assert.Equal("A different row with no link", VisibleText(block));
            Assert.Equal("A different row with no link", block.Text);
        });
    }

    // ------------------------------------------------------------- click resolution (TryResolveUrlAt)

    // TryResolveUrlAt is private - the same reflection approach ExternalTargetLauncherTests uses for
    // its own accept-path predicates (see that file's header comment for the reasoning): actually
    // OPENING a link calls Process.Start via ExternalTargetLauncher, which rule 6 in CLAUDE.md rules
    // out for a test, so what CAN be pinned down headlessly is the resolution step just before it -
    // does clicking at a given point find the right URL.
    //
    // Needs a real Window (not just UiTest.Run's dispatcher) so Measure/Arrange actually run and
    // block.TextLayout reflects real wrapped geometry - a bare unattached TextBlock's TextLayout is
    // null or stale.
    private static string? ResolveUrlAt(TextBlock block, Point point)
    {
        var method = typeof(TextLinkRenderer).GetMethod("TryResolveUrlAt", BindingFlags.NonPublic | BindingFlags.Static);
        return (string?)method!.Invoke(null, new object[] { block, point });
    }

    private static void WithLaidOutBlock(string text, double width, Action<TextBlock> body)
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock { TextWrapping = TextWrapping.Wrap };
            TextLinkRenderer.Apply(block, text);

            var window = new Window { Width = width, SizeToContent = SizeToContent.Height, Content = block };
            try
            {
                window.Show();
                window.Measure(new Size(width, double.PositiveInfinity));
                window.Arrange(new Rect(0, 0, width, window.DesiredSize.Height));
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();

                body(block);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Finds a point squarely inside the given (single-line) block's TEXT, by hit-testing along X
    // until HitTestPoint reports a TextPosition within [charStart, charEnd) - font metrics vary by
    // machine, so a guessed pixel offset is not reliable, but walking to the character index is.
    private static Point FindPointForCharacterRange(TextBlock block, int charStart, int charEnd)
    {
        var layout = block.TextLayout!;
        double y = layout.TextLines[0].Height / 2;

        for (double x = 0; x <= layout.Width; x += 1)
        {
            var hit = layout.HitTestPoint(new Point(x, y));
            if (hit.TextPosition >= charStart && hit.TextPosition < charEnd)
            {
                return new Point(x, y);
            }
        }

        throw new InvalidOperationException($"No point found for character range [{charStart},{charEnd})");
    }

    // A single-line note - the common case - resolves a click on its link.
    [Fact]
    public void A_click_on_a_single_line_links_own_run_resolves_its_url()
    {
        const string text = "See https://example.com for more";
        int urlStart = text.IndexOf("https://", StringComparison.Ordinal);

        WithLaidOutBlock(text, 400, block =>
        {
            var point = FindPointForCharacterRange(block, urlStart, urlStart + "https://example.com".Length);
            Assert.Equal("https://example.com", ResolveUrlAt(block, point));
        });
    }

    // A click over plain prose - not over any link run - resolves nothing.
    [Fact]
    public void A_click_over_plain_prose_resolves_no_url()
    {
        const string text = "See https://example.com for more";

        WithLaidOutBlock(text, 400, block =>
        {
            // "See " is plain prose before the link starts.
            var point = FindPointForCharacterRange(block, 0, 1);
            Assert.Null(ResolveUrlAt(block, point));
        });
    }

    // THE REGRESSION: a note long enough to WRAP past its first line, with the link on a LATER
    // line. Reported as "the note shows a link but does not react on click" - the link rendered
    // (styled, underlined) but every click on it was silently ignored.
    //
    // Root cause: Avalonia's TextLayout.HitTestPoint reports IsInside=false for every line except
    // the FIRST, even for a point squarely inside a later line's own text - verified directly
    // against the real API before writing the fix (a throwaway diagnostic test dumped
    // HitTestPoint's IsInside/TextPosition across a 5-line wrapped block: TextPosition kept
    // advancing correctly to the right per-line start index, 30/58/85/116, while IsInside stayed
    // false past y=15). TryResolveUrlAt no longer trusts IsInside - it works out containment itself
    // from each TextLine's own Height/Width - so this pins the fix down at the level a click
    // actually happens: forcing the block to a narrow width so the note wraps, then resolving a
    // point on the SECOND line.
    [Fact]
    public void A_click_on_a_link_that_wraps_past_the_first_line_still_resolves_its_url()
    {
        // Long enough, and the window narrow enough, to force at least two lines with the link
        // itself landing after the wrap - not on line one.
        const string text =
            "Bought at auction, no picture at all when powered on, see https://example.com/repair-notes for the full writeup and photos of the board";

        WithLaidOutBlock(text, 220, block =>
        {
            var layout = block.TextLayout!;
            Assert.True(layout.TextLines.Count > 1, "the test text must actually wrap to more than one line");

            // Find where the link run actually sits by walking the lines for the first one whose
            // source range overlaps the URL's character span, rather than hardcoding a guessed
            // pixel point - the exact wrap points depend on the font metrics of whatever machine
            // runs this.
            int urlStart = text.IndexOf("https://", StringComparison.Ordinal);
            double y = 0;
            double? hitY = null;
            foreach (var line in layout.TextLines)
            {
                if (urlStart >= line.FirstTextSourceIndex && urlStart < line.FirstTextSourceIndex + line.Length)
                {
                    hitY = y + (line.Height / 2);
                    break;
                }
                y += line.Height;
            }

            Assert.NotNull(hitY);
            Assert.True(hitY > 0, "the link must land on a line AFTER the first for this to actually test the wrap case");

            var url = ResolveUrlAt(block, new Point(2, hitY!.Value));
            Assert.Equal("https://example.com/repair-notes", url);
        });
    }

    // A click past the RIGHT EDGE of a wrapped line's own text - in the empty space beyond where
    // that line's content actually ends - must not resolve a link even when the line happens to end
    // in one, matching what IsInside protected against on a single-line block.
    [Fact]
    public void A_click_past_a_wrapped_lines_own_width_resolves_no_url()
    {
        const string text = "Photos and full writeup are at https://example.com/x";

        WithLaidOutBlock(text, 200, block =>
        {
            var layout = block.TextLayout!;
            var lastLine = layout.TextLines[^1];
            double y = layout.Height - (lastLine.Height / 2);

            // Comfortably to the right of the block's own laid-out width, which is what the line's
            // WidthIncludingTrailingWhitespace is bounded by - never inside any run.
            var url = ResolveUrlAt(block, new Point(block.Bounds.Width + 50, y));
            Assert.Null(url);
        });
    }
    // ---------------------------------------------------------------------------------------------
    // CARRYING THE WORKBOOKS FILTER INTO THE FULL EDITOR.
    //
    // The Workbooks tab publishes its "Find a previous repair" query through
    // TextLinkRenderer.SetActiveHighlightQuery, and every block this class renders marks it. That is
    // what makes a match visible inside the worklog editor, whose rows are DataTemplates over
    // bindings and so have no code-behind moment where a caller could hand segments in.
    //
    // Without it the filter would find a worklog by a comment left months ago, and then show that
    // comment unmarked once opened - leaving the user to re-read the entry hunting for the match.
    //
    // EVERY TEST HERE MUST RESET THE QUERY in a finally block: it is process-wide display state, so
    // a leaked value would mark runs in unrelated tests.
    // ---------------------------------------------------------------------------------------------

    private static void WithActiveQuery(string queryText, Action body)
    {
        TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse(queryText), queryText);

        try
        {
            body();
        }
        finally
        {
            TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse(null), null);
        }
    }

    // The core of the feature: a block rendered through the LinkText property - which is what the
    // editor's Work done, Comments, Photo and File rows use - marks the active query's matches.
    [Fact]
    public void A_templated_row_highlights_the_active_search_query()
    {
        UiTest.Run(() =>
        {
            WithActiveQuery("socket", () =>
            {
                var block = new TextBlock();
                TextLinkRenderer.SetLinkText(block, "Replaced the CPU socket today");

                // Highlighted text moves into Inlines, so Text is null - see the note on
                // WorkbooksSearchTests.VisibleText.
                Assert.Null(block.Text);

                var runs = block.Inlines!.OfType<Run>().ToList();
                Assert.Equal("Replaced the CPU socket today", string.Concat(runs.Select(r => r.Text)));

                var marked = runs.Where(r => r.Background != null).Select(r => r.Text).ToList();
                Assert.Equal(new[] { "socket" }, marked);
            });
        });
    }

    // With no search running the row keeps the cheap plain-Text path - the overwhelmingly common
    // case, and the one the Inlines split would cost layout time for nothing.
    [Fact]
    public void A_templated_row_with_no_active_search_stays_plain_text()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();
            TextLinkRenderer.SetLinkText(block, "Replaced the CPU socket today");

            Assert.Equal("Replaced the CPU socket today", block.Text);
            Assert.True(block.Inlines == null || block.Inlines.Count == 0);
        });
    }

    // Links and search highlighting have to COMPOSE, not overwrite one another: a search term
    // routinely lands inside a URL. Both markings must survive on the same text.
    [Fact]
    public void A_row_carrying_both_a_link_and_a_match_keeps_both_markings()
    {
        UiTest.Run(() =>
        {
            WithActiveQuery("example", () =>
            {
                var block = new TextBlock();
                TextLinkRenderer.SetLinkText(block, "See https://example.com for the pinout");

                var runs = block.Inlines!.OfType<Run>().ToList();

                Assert.Equal("See https://example.com for the pinout", string.Concat(runs.Select(r => r.Text)));
                Assert.Contains(runs, r => r.Background != null);              // the search hit
                Assert.Contains(runs, r => r.TextDecorations != null);          // the link
            });
        });
    }

    // The highlight-only property, for the editor's link-row headlines: those are searched, but the
    // whole row already opens the URL, so the headline must not offer a second clickable target.
    [Fact]
    public void The_highlight_only_property_marks_matches_without_linkifying()
    {
        UiTest.Run(() =>
        {
            WithActiveQuery("pinout", () =>
            {
                var block = new TextBlock();
                TextLinkRenderer.SetHighlightText(block, "https://example.com pinout sheet");

                var runs = block.Inlines!.OfType<Run>().ToList();

                Assert.Equal("https://example.com pinout sheet", string.Concat(runs.Select(r => r.Text)));
                Assert.Contains(runs, r => r.Background != null);
                Assert.DoesNotContain(runs, r => r.TextDecorations != null);
            });
        });
    }

    // Highlighting splits into Inlines, which disables TextTrimming - so with no search running the
    // block must fall back to plain Text and keep its ellipsis. That is what confines the trade-off
    // to the rows the user is actively hunting through.
    [Fact]
    public void The_highlight_only_property_keeps_plain_text_when_nothing_is_searched()
    {
        UiTest.Run(() =>
        {
            var block = new TextBlock();
            TextLinkRenderer.SetHighlightText(block, "A very long link headline");

            Assert.Equal("A very long link headline", block.Text);
            Assert.True(block.Inlines == null || block.Inlines.Count == 0);
        });
    }

    // Setting the same query twice reports no change, so a refresh pass whose query has not moved
    // does not force everything on screen to re-render. WorklogSearchQuery has no value equality of
    // its own, so this is compared on the raw text - a bug worth pinning, since comparing the parsed
    // objects would report a change on every single pass.
    [Fact]
    public void Setting_the_same_query_text_reports_no_change()
    {
        try
        {
            Assert.True(TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse("cpu"), "cpu"));

            // A DIFFERENT parsed object built from the same text - the shape the tab produces on
            // every refresh, since it re-parses the box each pass.
            Assert.False(TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse("cpu"), "cpu"));

            Assert.True(TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse("pla"), "pla"));
        }
        finally
        {
            TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse(null), null);
        }
    }

    // Clearing the filter must stop the marking everywhere, not just in the tab.
    [Fact]
    public void Clearing_the_active_query_stops_highlighting()
    {
        UiTest.Run(() =>
        {
            WithActiveQuery("socket", () =>
            {
                var block = new TextBlock();
                TextLinkRenderer.SetLinkText(block, "CPU socket");
                Assert.Null(block.Text);
            });

            // Re-rendered after the query was reset by WithActiveQuery's finally.
            var after = new TextBlock();
            TextLinkRenderer.SetLinkText(after, "CPU socket");

            Assert.Equal("CPU socket", after.Text);
        });
    }

    // ###########################################################################################
    // A BLOCK ALREADY ON SCREEN PICKS UP A QUERY CHANGE.
    //
    // This is what the return value of SetActiveHighlightQuery used to be for - the caller was
    // supposed to re-render whatever was showing - and the one call site discarded it, so blocks
    // already rendered kept marking the previous query's terms. Nothing else could have fixed the
    // worst case: ShowDialog does not block the dispatcher, so a worklog editor can be open over
    // the tab while a debounced refresh changes the query, and its rows are DataTemplate output
    // that only re-renders when the BOUND VALUE changes - which a query change is not.
    //
    // So the re-render is done inside SetActiveHighlightQuery itself. This fails against the
    // version that only stored the new query.
    // ###########################################################################################
    [Fact]
    public void Changing_the_query_re_marks_a_block_that_is_already_rendered()
    {
        UiTest.Run(() =>
        {
            try
            {
                TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse("socket"), "socket");

                var block = new TextBlock();
                TextLinkRenderer.SetLinkText(block, "CPU socket replaced");

                // Marked for "socket": split into Inlines, so Text is null.
                Assert.Null(block.Text);

                // The query moves to a term this text does NOT contain. The block has to come back
                // to plain text on its own - nothing re-sets its bound value.
                TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse("bravo"), "bravo");

                Assert.Equal("CPU socket replaced", block.Text);

                // And back again, so the re-render is not a one-way "give up highlighting".
                TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse("CPU"), "CPU");

                Assert.Null(block.Text);
            }
            finally
            {
                TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse(null), null);
            }
        });
    }

    // ###########################################################################################
    // THE QUERY CAN BE DROPPED WITHOUT ONE, which is what stops it outliving the box it was typed
    // into. It is process-wide state written only by the Workbooks tab's refresh, so once that tab
    // stops refreshing - hidden by the "Enable Worklog" checkbox, or simply never returned to while
    // worklogs are opened from the Schematics tab's own badges - the last query would keep washing
    // runs for a search the user is no longer running and can no longer see the box for.
    //
    // The Workbooks tab calls this from its detach; re-attaching republishes from the box, so a
    // plain tab switch loses nothing.
    // ###########################################################################################
    [Fact]
    public void Clearing_the_shared_query_stops_marking_blocks_already_rendered()
    {
        UiTest.Run(() =>
        {
            try
            {
                TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse("socket"), "socket");

                var block = new TextBlock();
                TextLinkRenderer.SetLinkText(block, "CPU socket replaced");
                Assert.Null(block.Text);

                TextLinkRenderer.ClearActiveHighlightQuery();

                Assert.Equal("CPU socket replaced", block.Text);

                // And nothing rendered afterwards is marked either.
                var after = new TextBlock();
                TextLinkRenderer.SetLinkText(after, "CPU socket replaced");
                Assert.Equal("CPU socket replaced", after.Text);
            }
            finally
            {
                TextLinkRenderer.SetActiveHighlightQuery(WorklogSearchQuery.Parse(null), null);
            }
        });
    }
}
