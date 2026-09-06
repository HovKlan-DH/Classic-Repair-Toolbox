using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.TextFormatting;
using Avalonia.Reactive;
using CRT;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Theming
{
    // ###########################################################################################
    // Renders a user-typed free-text field into a TextBlock with any web links in it made
    // clickable - the workbook Note, the worklog Description, and the Work done / Comment / Photo
    // comment / File comment rows.
    //
    // Which runs are links is TextLinkFinder's job (pure, unit tested); this turns its spans into
    // Avalonia Runs, styles the link ones, and opens the target through ExternalTargetLauncher -
    // the only sanctioned way to open an external target in this app.
    //
    // WHY THE CLICK IS ON THE TEXTBLOCK, NOT ON THE RUN: an Avalonia Run is an Inline, not a
    // Control - it has no pointer events of its own. So one PointerPressed on the block hit-tests
    // the character under the pointer (TextLayout.HitTestPoint) and looks that index up against the
    // spans. The alternative - a WrapPanel of per-run Controls - reflows differently from a real
    // TextBlock and loses hyphenation and trimming, in blocks that already wrap inside narrow
    // panels.
    //
    // Text with no link in it takes a fast path and stays a plain single-Text TextBlock: splitting
    // into Inlines costs measurably more to lay out, and the overwhelming majority of repair notes
    // contain no URL at all. This mirrors what TabWorkbooks.ApplyHighlightedText already does for
    // search highlighting, and the two compose - see ApplySegments.
    // ###########################################################################################
    public static class TextLinkRenderer
    {
        // One Hand cursor for every linked block rather than one per block: Cursor is IDisposable
        // and holds an HCURSOR on Win32, and these blocks are rebuilt in bulk on every refresh.
        private static readonly Cursor HandCursor = new(StandardCursorType.Hand);

        // ###########################################################################################
        // The spans of the text currently rendered into a given block, so the click handler can map
        // a character index back to a URL. Attached to the block rather than kept in a dictionary
        // keyed by block: a ConditionalWeakTable would work too, but an attached property is what
        // Avalonia already collects along with the control.
        // ###########################################################################################
        private static readonly AttachedProperty<IReadOnlyList<TextLinkSpan>?> LinkSpansProperty =
            AvaloniaProperty.RegisterAttached<TextBlock, IReadOnlyList<TextLinkSpan>?>(
                "LinkSpans", typeof(TextLinkRenderer));

        // ###########################################################################################
        // The Cursor the block carried before this class started swapping it for a Hand - restored
        // whenever the pointer leaves a link run, and when the block stops being linked at all.
        //
        // Assigning null instead would be wrong, not merely lossy: several of these blocks sit
        // inside a row that sets its own cursor (the board pane's pills carry a Hand; the editor's
        // photo and file rows carry a resize cursor on the drag handle), and nulling it there
        // silently drops that row's cursor the first time its prose is hovered.
        //
        // Two properties rather than one nullable: the recorded value is very often null (most of
        // these blocks set no cursor of their own), so "nothing was set" has to stay distinguishable
        // from "nothing recorded yet" - otherwise every re-render would re-capture, and a Hand this
        // class had already assigned would become the block's "original".
        // ###########################################################################################
        private static readonly AttachedProperty<Cursor?> OriginalCursorProperty =
            AvaloniaProperty.RegisterAttached<TextBlock, Cursor?>(
                "OriginalCursor", typeof(TextLinkRenderer));

        private static readonly AttachedProperty<bool> HasOriginalCursorProperty =
            AvaloniaProperty.RegisterAttached<TextBlock, bool>(
                "HasOriginalCursor", typeof(TextLinkRenderer));

        // ###########################################################################################
        // The markup-facing form: `theme:TextLinkRenderer.LinkText="{Binding Comment}"` INSTEAD of
        // `Text="{Binding Comment}"` on a TextBlock, for the read-only prose rows that live in a
        // DataTemplate (the worklog editor's Work done, Comments, Photos and Files lists).
        //
        // An attached property rather than the Apply call the code-built blocks use, because those
        // rows are built by a template from a binding - there is no code-behind moment holding both
        // the block and its text. Setting Text as well would be a bug, not a redundancy: a TextBlock
        // carrying both Text and Inlines renders the Text and silently ignores the Inlines.
        // ###########################################################################################
        public static readonly AttachedProperty<string?> LinkTextProperty =
            AvaloniaProperty.RegisterAttached<TextBlock, string?>(
                "LinkText", typeof(TextLinkRenderer));

        public static void SetLinkText(TextBlock block, string? value) =>
            block.SetValue(LinkTextProperty, value);

        public static string? GetLinkText(TextBlock block) =>
            block.GetValue(LinkTextProperty);

        static TextLinkRenderer()
        {
            // Re-render whenever the bound value changes - a template recycles its containers, so the
            // same block is handed a different row's text as the list scrolls or is rebuilt.
            //
            // Subscribed on the property's own change stream rather than per instance: the stream is
            // static and outlives every block, so there is nothing to unsubscribe and a recycled
            // container cannot leak a handler.
            LinkTextProperty.Changed.Subscribe(new AnonymousObserver<AvaloniaPropertyChangedEventArgs<string?>>(args =>
            {
                if (args.Sender is TextBlock block)
                {
                    // Registered so a later query change can re-mark it - this block is
                    // DataTemplate output with no other route back to it. See TrackBlock.
                    TrackBlock(block);
                    Apply(block, args.NewValue.GetValueOrDefault());
                }
            }));

            // The highlight-only twin, for the same reason and with the same lifetime.
            HighlightTextProperty.Changed.Subscribe(new AnonymousObserver<AvaloniaPropertyChangedEventArgs<string?>>(args =>
            {
                if (args.Sender is TextBlock block)
                {
                    TrackBlock(block);
                    ApplyHighlightOnly(block, args.NewValue.GetValueOrDefault());
                }
            }));
        }

        // ###########################################################################################
        // THE SEARCH QUERY CURRENTLY BEING HIGHLIGHTED, app-wide.
        //
        // Set by the Workbooks tab whenever its "Find a previous repair" box changes, and consulted
        // by every block this class renders. That is what carries the filter's yellow highlighting
        // into the full worklog EDITOR, which is a separate window built from DataTemplates: a
        // search that matched a comment left three months ago used to filter the worklog into view
        // and then show that comment unmarked, leaving the user to re-read the whole entry looking
        // for what had matched. The fields the editor shows are exactly the ones
        // WorklogSearchIndex.ForEntry searches, so anywhere a match can be found, it can now be seen.
        //
        // A STATIC rather than a parameter threaded through the editor, deliberately: the rows that
        // need it are built by DataTemplates from bindings (see LinkTextProperty), so there is no
        // code-behind moment holding both a block and the query, and every alternative - a
        // DataContext wrapper per row, an attached property set per block - would have to be applied
        // at every one of those template sites and would be silently missing from any added later.
        // The cost of the static is that it is process-wide state; it is acceptable here because it
        // is display-only (it can never change what is SAVED) and it is written through exactly one
        // method. It is CLEARED when the Workbooks tab detaches (ClearActiveHighlightQuery), which
        // is what stops a query outliving the box it was typed into - the tab being hidden by the
        // Configuration checkbox, or a worklog opened from the Schematics tab, would otherwise leave
        // the last query washing runs for a search nobody is running any more.
        //
        // Default is the empty query, which every consumer treats as "no highlighting" and which is
        // what a test or a build with no Workbooks tab on screen sees.
        // ###########################################################################################
        private static WorklogSearchQuery _activeHighlightQuery = WorklogSearchQuery.Parse(null);

        // The raw text the active query was parsed from, kept only to tell whether a newly supplied
        // query differs from the current one. WorklogSearchQuery is a plain sealed class with no
        // value equality, so comparing the objects would compare REFERENCES - and the tab parses a
        // fresh one on every refresh pass, so that comparison would report a change every time.
        private static string _activeHighlightQueryText = string.Empty;

        // ###########################################################################################
        // Points the shared highlighting at a query, and reports whether it actually changed.
        //
        // A change RE-RENDERS the blocks already on screen itself (RerenderTrackedBlocks) - blocks do
        // not re-read this on their own, and the callers that would have to do it have no handle on
        // the DataTemplate rows that need it most. The returned bool is therefore only "did it
        // change", for a caller wanting to skip work of its own; ignoring it is safe.
        //
        // Takes the raw text alongside the parsed query because that text is the only thing the two
        // can be compared by - see _activeHighlightQueryText.
        // ###########################################################################################
        public static bool SetActiveHighlightQuery(WorklogSearchQuery query, string? queryText)
        {
            string text = queryText ?? string.Empty;

            if (string.Equals(_activeHighlightQueryText, text, StringComparison.Ordinal))
            {
                return false;
            }

            _activeHighlightQueryText = text;
            _activeHighlightQuery = query ?? WorklogSearchQuery.Parse(null);

            // The re-render is done HERE rather than left to the caller. It used to be the caller's
            // job, announced by this method's return value - and the one call site discarded it, so
            // blocks already on screen kept marking the previous query's terms. That is not a
            // theoretical gap: ShowDialog does not block the dispatcher, so a worklog editor can be
            // open over the tab while a debounced refresh changes the query, and its rows are
            // DataTemplate output that only re-renders when the BOUND VALUE changes - which a query
            // change is not. Nothing else could have fixed those rows; the caller has no handle on
            // them. The bool is still returned, but now only as "did it change", for a caller that
            // wants to skip work of its own.
            RerenderTrackedBlocks();

            return true;
        }

        // ###########################################################################################
        // Every block this class has rendered that is still alive, so a change of query can re-mark
        // the ones already on screen - above all the worklog editor's DataTemplate rows, which have
        // no other route back to them (see SetActiveHighlightQuery).
        //
        // WEAK references throughout: these blocks belong to windows and lists that come and go, and
        // a static list of strong references would keep every worklog row ever rendered alive for
        // the life of the process. Dead entries are swept on each pass, which is the only time this
        // list is walked at all - a query change, i.e. a debounced keystroke in one text box.
        //
        // Keyed on the block itself rather than on a per-block id, and re-registering the same block
        // is a no-op, so the recycled containers a DataTemplate hands back do not accumulate.
        // ###########################################################################################
        private static readonly List<WeakReference<TextBlock>> _trackedBlocks = new();

        private static void TrackBlock(TextBlock block)
        {
            for (int i = _trackedBlocks.Count - 1; i >= 0; i--)
            {
                if (!_trackedBlocks[i].TryGetTarget(out var existing))
                {
                    _trackedBlocks.RemoveAt(i);
                    continue;
                }

                if (ReferenceEquals(existing, block))
                {
                    return;
                }
            }

            _trackedBlocks.Add(new WeakReference<TextBlock>(block));
        }

        // ###########################################################################################
        // Re-applies whichever of the two attached properties a tracked block was rendered through,
        // so it picks up the new query. A block rendered by a direct Apply/ApplySegments call - the
        // Workbooks tab's own code-built blocks - carries neither property and is skipped: the tab
        // rebuilds those itself in the very same refresh that changed the query, so re-rendering
        // them here would only duplicate that work.
        // ###########################################################################################
        private static void RerenderTrackedBlocks()
        {
            for (int i = _trackedBlocks.Count - 1; i >= 0; i--)
            {
                if (!_trackedBlocks[i].TryGetTarget(out var block))
                {
                    _trackedBlocks.RemoveAt(i);
                    continue;
                }

                // Reading a control's property is only legal from the dispatcher that OWNS it, and
                // this is reached from a plain state setter that callers may use from anywhere.
                // Checked per block rather than once up front: a block belongs to its own dispatcher,
                // and the headless test session gives each test its own - so a block left tracked by
                // an earlier test is on a thread this one has no access to, and touching it throws.
                // Skipping is right rather than posting: a block this thread cannot reach is not on
                // this thread's screen either, and posting from a thread with no initialised
                // dispatcher stands the platform up on the wrong one and takes unrelated tests down.
                if (!block.Dispatcher.CheckAccess())
                {
                    continue;
                }

                string? linkText = block.GetValue(LinkTextProperty);
                if (linkText != null)
                {
                    Apply(block, linkText);
                    continue;
                }

                string? highlightText = block.GetValue(HighlightTextProperty);
                if (highlightText != null)
                {
                    ApplyHighlightOnly(block, highlightText);
                }
            }
        }

        // ###########################################################################################
        // Drops the shared query, so nothing is highlighted anywhere until a search sets one again.
        //
        // WHY THIS IS NEEDED: the query is process-wide and is written from exactly one place - the
        // Workbooks tab's refresh. Once that tab stops refreshing, the last query it published stays
        // active for the rest of the session. Untick "Enable Worklog" in Configuration and the tab is
        // hidden and never refreshes again; or simply open a worklog from the SCHEMATICS tab's own
        // badges without going back to Workbooks. Every block rendered afterwards - the description,
        // every comment, work-done and file-comment row in the editor - would keep washing every
        // occurrence of a term the user is no longer searching for and can no longer see the box for.
        //
        // Called from the Workbooks tab's detach, which covers both: a tab switch and the tab being
        // hidden outright both detach it, and its own re-attach republishes the query from the box.
        // ###########################################################################################
        public static void ClearActiveHighlightQuery()
        {
            SetActiveHighlightQuery(WorklogSearchQuery.Parse(null), null);
        }

        // ###########################################################################################
        // Splits text against the active query, or null when nothing is being highlighted.
        //
        // Null rather than "one non-matching segment" for the no-search case: ApplySegments treats
        // null as the plain-text fast path, which skips the Inlines split entirely - and the
        // overwhelming majority of renders happen with no search active.
        // ###########################################################################################
        private static IReadOnlyList<(string Text, bool IsMatch)>? SplitForActiveQuery(string? text)
        {
            if (_activeHighlightQuery.IsEmpty || string.IsNullOrEmpty(text))
            {
                return null;
            }

            return _activeHighlightQuery.SplitIntoSegments(text);
        }

        // ###########################################################################################
        // The markup-facing form for text that must carry SEARCH HIGHLIGHTING but NOT link
        // detection: `theme:TextLinkRenderer.HighlightText="{Binding Headline}"` instead of `Text`.
        //
        // Exists for the worklog editor's link-row headlines. Those are searched
        // (WorklogSearchIndex.ForEntry yields both a link's Headline and its Url), so a match there
        // has to be visible - but the row is ALREADY a link: the whole row is clickable and opens
        // the URL, so marking runs inside the headline as separately clickable would offer a second,
        // narrower target for the same action.
        //
        // Highlighting still splits into Inlines, so a block using this loses TextTrimming (which
        // only applies to a block's Text). That is why it is applied ONLY while a search is active -
        // with no query the block keeps its plain Text and its ellipsis, and the trimming is only
        // given up on the rows the user is actively hunting through.
        // ###########################################################################################
        public static readonly AttachedProperty<string?> HighlightTextProperty =
            AvaloniaProperty.RegisterAttached<TextBlock, string?>(
                "HighlightText", typeof(TextLinkRenderer));

        public static void SetHighlightText(TextBlock block, string? value) =>
            block.SetValue(HighlightTextProperty, value);

        public static string? GetHighlightText(TextBlock block) =>
            block.GetValue(HighlightTextProperty);

        // ###########################################################################################
        // Marks the active query's matches in a block, with no link detection - see
        // HighlightTextProperty for why the two are separated.
        //
        // Falls back to plain Text when nothing is being highlighted, which keeps TextTrimming
        // working for every render that is not part of an active search.
        // ###########################################################################################
        public static void ApplyHighlightOnly(TextBlock block, string? text)
        {
            if (block == null)
            {
                return;
            }

            var segments = SplitForActiveQuery(text);

            block.Inlines?.Clear();

            if (segments == null || !segments.Any(segment => segment.IsMatch))
            {
                block.Text = text ?? string.Empty;
                return;
            }

            // Text must go before Inlines are added: a TextBlock carrying both renders the Text and
            // silently ignores the Inlines.
            block.Text = null;

            var hitBackground = ThemeResources.ResolveBrush("Workbooks_SearchHit_Bg", Brushes.Yellow);
            var hitForeground = ThemeResources.ResolveBrush("Workbooks_SearchHit_Fg", Brushes.Black);

            foreach (var segment in segments)
            {
                var run = new Run(segment.Text);

                if (segment.IsMatch)
                {
                    run.Background = hitBackground;
                    run.Foreground = hitForeground;
                }

                block.Inlines!.Add(run);
            }
        }

        // ###########################################################################################
        // Sets a block's content, rendering any links in it as clickable runs AND marking whatever
        // the active search query matches.
        //
        // The search split is applied automatically here rather than being the caller's job: this is
        // what the DataTemplate rows go through (via LinkTextProperty), and those have no
        // code-behind moment at which a caller could supply segments. A caller that has already
        // computed its own segments calls ApplySegments directly instead - the Workbooks tab does,
        // because it holds the query itself and splits against the same text for its own blocks.
        // ###########################################################################################
        public static void Apply(TextBlock block, string? text)
        {
            ApplySegments(block, text, SplitForActiveQuery(text));
        }

        // ###########################################################################################
        // Sets a block's content from PRE-SPLIT segments - the search-highlighting split - with links
        // applied on top of them.
        //
        // The two splits are independent and can cut the text at different places (a search term can
        // land in the middle of a URL, and routinely does: searching "example" while a note contains
        // "https://example.com"). So the segment boundaries are MERGED rather than one being applied
        // after the other: every run carries both markings, and the text still comes out whole.
        //
        // Pass segments: null when there is no search active - the whole text is then one
        // non-matching segment, which is the plain Apply case.
        // ###########################################################################################
        public static void ApplySegments(
            TextBlock block,
            string? text,
            IReadOnlyList<(string Text, bool IsMatch)>? segments)
        {
            if (block == null)
            {
                return;
            }

            block.Inlines?.Clear();
            block.SetValue(LinkSpansProperty, null);

            string value = text ?? string.Empty;

            // MergeRuns walks the segments by LENGTH against the text's own offsets, so segments
            // that do not concatenate back to exactly this text would put its search-hit flag on the
            // wrong characters - silently, since the runs still rebuild the string and nothing
            // throws. Every caller passes SplitIntoSegments(text) of the same text, so this holds;
            // it is checked rather than trusted because the failure is invisible. Dropping to
            // links-only is the right degradation: the text is still whole and correct, only the
            // highlight wash is missing.
            if (segments != null && !SegmentsCoverText(value, segments))
            {
                Logger.Warning(
                    $"Search-highlight segments do not match the text they were split from ([{value.Length}] characters); rendering without highlighting");
                segments = null;
            }

            var linkSpans = TextLinkFinder.FindSpans(value);

            bool hasLink = false;
            foreach (var span in linkSpans)
            {
                if (span.IsLink)
                {
                    hasLink = true;
                    break;
                }
            }

            bool hasHighlight = false;
            if (segments != null)
            {
                foreach (var segment in segments)
                {
                    if (segment.IsMatch)
                    {
                        hasHighlight = true;
                        break;
                    }
                }
            }

            if (!hasLink && !hasHighlight)
            {
                // Nothing to mark - the cheap path, and the one almost every block takes.
                block.Text = value;
                EnsureHandler(block, isLinked: false);
                return;
            }

            // Text must be cleared before Inlines are added: a TextBlock carrying both renders the
            // Text and ignores the Inlines entirely.
            block.Text = null;

            var linkForeground = ThemeResources.ResolveBrush("Link_Fg", Brushes.Blue);
            var hitBackground = ThemeResources.ResolveBrush("Workbooks_SearchHit_Bg", Brushes.Yellow);
            var hitForeground = ThemeResources.ResolveBrush("Workbooks_SearchHit_Fg", Brushes.Black);

            foreach (var (runText, isLink, isMatch) in MergeRuns(value, linkSpans, segments))
            {
                var run = new Run(runText);

                if (isMatch)
                {
                    run.Background = hitBackground;
                    run.Foreground = hitForeground;
                }

                if (isLink)
                {
                    run.TextDecorations = TextDecorations.Underline;

                    // A search hit inside a link keeps the hit's foreground: the wash is what tells
                    // the user WHY the record matched, and the underline already says it is a link.
                    if (!isMatch)
                    {
                        run.Foreground = linkForeground;
                    }
                }

                block.Inlines!.Add(run);
            }

            block.SetValue(LinkSpansProperty, hasLink ? linkSpans : null);
            EnsureHandler(block, hasLink);
        }

        // ###########################################################################################
        // Whether the segments concatenate to exactly the text they claim to be a split of - the
        // invariant MergeRuns depends on, checked without building the joined string.
        // ###########################################################################################
        private static bool SegmentsCoverText(string text, IReadOnlyList<(string Text, bool IsMatch)> segments)
        {
            int position = 0;

            foreach (var segment in segments)
            {
                string segmentText = segment.Text ?? string.Empty;

                if (position + segmentText.Length > text.Length ||
                    string.CompareOrdinal(text, position, segmentText, 0, segmentText.Length) != 0)
                {
                    return false;
                }

                position += segmentText.Length;
            }

            return position == text.Length;
        }

        // ###########################################################################################
        // Walks the text once, cutting a new run at every boundary EITHER split introduces, so each
        // run is uniformly a link or not and uniformly a search hit or not.
        //
        // Both inputs are ordered and cover the text completely, so this is a merge rather than a
        // search: the next boundary is whichever of the two comes first.
        // ###########################################################################################
        private static IEnumerable<(string Text, bool IsLink, bool IsMatch)> MergeRuns(
            string text,
            IReadOnlyList<TextLinkSpan> linkSpans,
            IReadOnlyList<(string Text, bool IsMatch)>? segments)
        {
            int position = 0;
            int linkIndex = 0;
            int segmentIndex = 0;
            int segmentStart = 0;

            while (position < text.Length)
            {
                // Advance past any span that ends at or before the cursor.
                while (linkIndex < linkSpans.Count && linkSpans[linkIndex].End <= position)
                {
                    linkIndex++;
                }

                while (segments != null &&
                       segmentIndex < segments.Count &&
                       segmentStart + segments[segmentIndex].Text.Length <= position)
                {
                    segmentStart += segments[segmentIndex].Text.Length;
                    segmentIndex++;
                }

                bool isLink = linkIndex < linkSpans.Count &&
                              linkSpans[linkIndex].Start <= position &&
                              linkSpans[linkIndex].IsLink;

                bool isMatch = segments != null &&
                               segmentIndex < segments.Count &&
                               segments[segmentIndex].IsMatch;

                int next = text.Length;

                if (linkIndex < linkSpans.Count)
                {
                    next = Math.Min(next, linkSpans[linkIndex].End);
                }

                if (segments != null && segmentIndex < segments.Count)
                {
                    next = Math.Min(next, segmentStart + segments[segmentIndex].Text.Length);
                }

                // Defensive: a zero-length step would spin forever. Neither producer emits an empty
                // span, but this loop is the one place where that would hang the UI thread rather
                // than showing wrong text, so it is not left to trust.
                if (next <= position)
                {
                    next = position + 1;
                }

                yield return (text[position..next], isLink, isMatch);
                position = next;
            }
        }

        // ###########################################################################################
        // Wires (or unwires) the block's click handling and cursor.
        //
        // Subscribing once and leaving it attached would be simpler, but these blocks are rebuilt in
        // bulk and a block can go from linked to unlinked when its text is edited - leaving a Hand
        // cursor on prose that is no longer clickable. Both halves are idempotent, so a block that
        // is refreshed repeatedly does not accumulate handlers.
        // ###########################################################################################
        private static void EnsureHandler(TextBlock block, bool isLinked)
        {
            block.PointerPressed -= OnBlockPointerPressed;
            block.PointerMoved -= OnBlockPointerMoved;

            if (!isLinked)
            {
                RestoreOriginalCursor(block);
                return;
            }

            // Recorded before the first Hand is ever assigned, and only once - a re-render of an
            // already-linked block must not capture the Hand this class itself put there.
            if (!block.GetValue(HasOriginalCursorProperty))
            {
                block.SetValue(OriginalCursorProperty, block.Cursor);
                block.SetValue(HasOriginalCursorProperty, true);
            }

            // A TextBlock with no background is not hit-testable in Avalonia - the same trap the
            // Links and Files rows in the worklog editor each carry a comment about.
            block.Background ??= Brushes.Transparent;

            block.PointerPressed += OnBlockPointerPressed;
            block.PointerMoved += OnBlockPointerMoved;
        }

        // ###########################################################################################
        // The pointer shows a Hand only while it is actually over a link run, not over the whole
        // block - a block is usually mostly prose, and a Hand over all of it would promise that
        // clicking anywhere does something.
        // ###########################################################################################
        private static void OnBlockPointerMoved(object? sender, PointerEventArgs e)
        {
            if (sender is not TextBlock block)
            {
                return;
            }

            bool overLink = TryResolveUrlAt(block, e.GetPosition(block)) != null;

            // Off a link, the block goes back to whatever cursor it had before this class touched
            // it - NOT to null, which would discard the row's own cursor. See OriginalCursorProperty.
            var wanted = overLink
                ? HandCursor
                : block.GetValue(OriginalCursorProperty);

            // Only assign on a change: this runs on every pointer-move frame.
            if (!ReferenceEquals(block.Cursor, wanted))
            {
                block.Cursor = wanted;
            }
        }

        // ###########################################################################################
        // Puts back the cursor the block carried before it was first linked, and forgets the record -
        // so a block that becomes linked again later captures its state afresh rather than restoring
        // a stale one.
        // ###########################################################################################
        private static void RestoreOriginalCursor(TextBlock block)
        {
            if (!block.GetValue(HasOriginalCursorProperty))
            {
                // Never linked, so nothing of this class's doing is on it to undo.
                return;
            }

            block.Cursor = block.GetValue(OriginalCursorProperty);
            block.SetValue(OriginalCursorProperty, null);
            block.SetValue(HasOriginalCursorProperty, false);
        }

        private static void OnBlockPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not TextBlock block)
            {
                return;
            }

            string? url = TryResolveUrlAt(block, e.GetPosition(block));
            if (url == null)
            {
                return;
            }

            // Handled so a click on a link does not also fire whatever the surrounding row does -
            // several of these blocks sit inside a card that opens an editor when clicked, and both
            // firing would open a window on top of the browser.
            e.Handled = true;

            ExternalTargetLauncher.TryOpen(url);
        }

        // ###########################################################################################
        // The URL of the link run under a point in the block's own coordinate space, or null when
        // the point is over plain text (or over the block's padding, past the end of the last line).
        //
        // TextLayout.HitTestPoint returns the NEAREST character for a point outside the text too, so
        // in principle IsInside should be checked rather than trusting the returned index - without
        // it, clicking the empty space to the right of a line that ends in a link would open it.
        //
        // BUT IsInside is unusable for a WRAPPED (multi-line) block: it comes back false for every
        // line except the first, even for a point squarely inside a later line's own text - verified
        // directly against TextLayout.HitTestPoint with a 5-line wrapped block, where every Y past
        // the first line's height returned IsInside=false while TextPosition kept advancing correctly
        // (30, 58, 85, 116 - the right start-of-line index each time). A single-line block (the
        // common case - most notes and descriptions are short) is unaffected, since there is only
        // one line for IsInside to be wrong about. Reported as a link in a multi-line workbook Note
        // rendering as a link but never reacting to a click - it only ever affected text long enough
        // to wrap past its first line.
        //
        // So containment is worked out here instead, per line, using ONLY properties TextLine itself
        // documents (Height, FirstTextSourceIndex, Length, Width/WidthIncludingTrailingWhitespace) -
        // TextPosition from HitTestPoint stays trustworthy throughout and is still what is used to
        // resolve the character index.
        // ###########################################################################################
        private static string? TryResolveUrlAt(TextBlock block, Point point)
        {
            var spans = block.GetValue(LinkSpansProperty);
            if (spans == null || spans.Count == 0)
            {
                return null;
            }

            var layout = block.TextLayout;
            if (layout == null)
            {
                return null;
            }

            if (!IsPointOverAnyLine(layout, point))
            {
                return null;
            }

            int index = layout.HitTestPoint(point).TextPosition;

            foreach (var span in spans)
            {
                if (span.IsLink && index >= span.Start && index < span.End)
                {
                    return span.Url;
                }
            }

            return null;
        }

        // ###########################################################################################
        // Whether point falls within some line's own drawn area: vertically inside that line's height
        // band (walking top-down, since TextLine carries no absolute Y of its own) and horizontally
        // no further right than the line's own width - trailing whitespace included, so a link that
        // ends a line stays clickable right up to where it visually ends, matching what IsInside did
        // for the first line before this replaced it.
        // ###########################################################################################
        private static bool IsPointOverAnyLine(TextLayout layout, Point point)
        {
            if (point.X < 0)
            {
                return false;
            }

            double y = 0;

            foreach (var line in layout.TextLines)
            {
                double top = y;
                double bottom = y + line.Height;

                if (point.Y >= top && point.Y < bottom)
                {
                    return point.X <= line.WidthIncludingTrailingWhitespace;
                }

                y = bottom;
            }

            return false;
        }
    }
}
