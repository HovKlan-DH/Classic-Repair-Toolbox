using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace Handlers.Theming
{
    // ###########################################################################################
    // The INFORMATIONAL (non-selectable) variants of the worklog status pill and category chip.
    //
    // These two shapes each have exactly two meanings in this app, and conflating them has been
    // reported in both directions:
    //
    //   SELECTABLE - only inside WorklogEntryEditorWindow, where clicking one CHOOSES it. The
    //                chosen one is FILLED with its colour and its label goes white; the others are
    //                grey-outlined. That pair of looks IS a selection state, and it stays where it
    //                is (ApplyStatePillVisualState / ApplyCategoryChipVisualState).
    //
    //   INFORMATIONAL - everywhere else: the worklog bar, a workbook card, the Workbooks tab's
    //                top-line, and an entry's detail card. Nothing there is selectable, so nothing
    //                may look "chosen". This class builds that variant, and it is the ONLY place
    //                its visual is decided.
    //
    // THE VISUAL: outlined - a Form_Bg fill, a 1px border in the thing's OWN colour (the state
    // colour for a status pill, the category colour for a category chip), glyph and label in that
    // same colour. The coloured rather than grey border is what makes these read as information at
    // a glance; 1px rather than the 2px three of these sites used is what makes them all read as
    // ONE pill rather than as several near-misses.
    //
    // WHY ONE BUILDER: before this, five sites drew these by hand - the worklog bar (2px, coloured),
    // the workbook card (2px, coloured), the Workbooks top-line (2px, coloured), and the entry
    // detail card's two (1px, GREY, label in the ordinary foreground). Each carried a comment
    // asserting they must all look the same, which nothing enforced and which was not in fact true
    // of any two of them. Now the assertion is the code.
    //
    // Builds controls but reads no tab state, so it lives here beside WorklogBadgeBuilder rather
    // than on any one UserControl - and so a headless test can build one without a window.
    // ###########################################################################################
    public static class WorklogInfoPillBuilder
    {
        // 1px, everywhere, on purpose. The three sites that used 2px were the ones the uniformity
        // complaint was actually about: beside the 1px pills on an entry card they read as a
        // different component rather than the same one.
        private static readonly Thickness InfoBorderThickness = new(1);

        // The pill (status) is fully rounded; the chip (category) is only softened. That difference
        // is deliberate and predates this class - it is how the two are told apart at a glance when
        // they sit side by side, and the editor's own selectable versions differ the same way.
        private static readonly CornerRadius PillCornerRadius = new(10);

        private static readonly CornerRadius ChipCornerRadius = new(3);

        // fa-regular note-sticky / fa-solid paint-roller / fa-solid triangle-exclamation - the same
        // codepoints WorklogEntryEditorWindow.axaml declares for its three selectable category
        // chips, spelled as hex so this file stays plain ASCII. Note is the one category whose
        // glyph comes from the Regular weight rather than Solid.
        private const int NoteCategoryCodepoint = 0xF15C;

        private const int CosmeticCategoryCodepoint = 0xF5D0;

        private const int IssueCategoryCodepoint = 0xF188;

        private static readonly Dictionary<string, (int Codepoint, bool IsRegular)> CategoryIconsByName =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Note"] = (NoteCategoryCodepoint, true),
                ["Cosmetic"] = (CosmeticCategoryCodepoint, false),
                ["Issue"] = (IssueCategoryCodepoint, false)
            };

        // ###########################################################################################
        // An Open/Closed status pill for a workbook or an entry, in the informational visual.
        //
        // fontSize varies by site (10 on a workbook card, 11 on an entry card, 12 in the worklog
        // bar) because the surrounding text does; the STYLE does not vary, which is the point.
        // ###########################################################################################
        public static Border BuildStatePill(string state, double fontSize = 11.0, int? count = null)
        {
            // A ZERO count is muted to a neutral grey rather than the state's own colour - asked
            // for directly, after the summary strip's five always-present pills (including the
            // zeroes, see BuildCountPills' own comment) blended together with no way to tell at a
            // glance which categories/states actually have entries. Only a COUNTED zero is muted;
            // an uncounted pill (a workbook card, an entry card) always names a real record's own
            // status and must never be grey.
            IBrush stateBrush = count == 0
                ? MutedPillBrush
                : new SolidColorBrush(ResolveStateColor(state));

            var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };

            // The count, when this pill is being used to report HOW MANY (the workbook summary
            // strip) rather than to state one record's own status. It leads the pill - "2 Open" -
            // because that is the order the sentence reads in, and it is bold because in the
            // summary the numbers are the content and the words merely label them.
            if (count.HasValue)
            {
                content.Children.Add(BuildCountLabel(count.Value, fontSize, stateBrush, isMuted: count == 0));
            }
            else
            {
                // NO glyph on a counted pill. A padlock between a number and its label reads as a
                // third piece of information rather than as decoration - "2 [lock] Open" invites
                // the question of what the lock is counting. On an uncounted pill the glyph is the
                // only thing distinguishing Open from Closed at a glance, so it stays.
                //
                // Resolved HERE rather than at the top of the method: it picks the padlock, which
                // only this branch draws, and a zero-count pill takes MutedPillBrush regardless of
                // state - computing it up there was dead work that also read as though the state
                // colour were still in play on a muted pill.
                bool isResolved = WorklogManager.IsResolvedState(state);

                content.Children.Add(BuildGlyph(
                    WorklogGlyphs.GlyphFor(isResolved),
                    ThemeResources.ResolveFontAwesomeSolid(),
                    WorklogGlyphs.CodepointFor(isResolved),
                    fontSize,
                    stateBrush));
            }

            content.Children.Add(BuildLabel(state, fontSize, stateBrush));

            return new Border
            {
                Background = ThemeResources.ResolveBrush("Form_Bg", Brushes.Transparent),
                BorderBrush = stateBrush,
                BorderThickness = InfoBorderThickness,
                CornerRadius = PillCornerRadius,
                Padding = new Thickness(10, 2),
                VerticalAlignment = VerticalAlignment.Center,
                Child = content
            };
        }

        // ###########################################################################################
        // A Note/Cosmetic/Issue category chip in the informational visual.
        //
        // Its border and text take the CATEGORY's own colour, for the same reason the status pill's
        // take the state's: this was asked for explicitly after the chip shipped grey-outlined
        // beside a colour-outlined pill, which made the two look like different kinds of control.
        // ###########################################################################################
        public static Border BuildCategoryChip(string category, double fontSize = 11.0, int? count = null)
        {
            var (codepoint, isRegular) = CategoryIconsByName.TryGetValue(category, out var icon)
                ? icon
                : (NoteCategoryCodepoint, true);

            // See BuildStatePill's own note on why a COUNTED zero is muted and an uncounted chip
            // never is.
            IBrush categoryBrush = count == 0
                ? MutedPillBrush
                : new SolidColorBrush(ResolveCategoryColor(category));

            var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 5 };

            // See BuildStatePill's own note: a leading bold count turns this from "what this entry
            // is" into "how many of these the workbook holds", which is what the summary needs.
            if (count.HasValue)
            {
                content.Children.Add(BuildCountLabel(count.Value, fontSize, categoryBrush, isMuted: count == 0));
            }
            else
            {
                // Dropped on a counted chip for the same reason as the status pill's padlock - see
                // its note. The category's own colour still identifies it without the icon.
                content.Children.Add(BuildGlyph(
                    char.ConvertFromUtf32(codepoint),
                    isRegular ? ThemeResources.ResolveFontAwesomeRegular() : ThemeResources.ResolveFontAwesomeSolid(),
                    codepoint,
                    fontSize,
                    categoryBrush));
            }

            content.Children.Add(BuildLabel(category, fontSize, categoryBrush));

            return new Border
            {
                Background = ThemeResources.ResolveBrush("Form_Bg", Brushes.Transparent),
                BorderBrush = categoryBrush,
                BorderThickness = InfoBorderThickness,
                CornerRadius = ChipCornerRadius,
                Padding = new Thickness(8, 3),
                VerticalAlignment = VerticalAlignment.Center,
                Child = content
            };
        }

        // ###########################################################################################
        // Restyles an EXISTING pill declared in markup, for the two sites that cannot simply be
        // handed a new Border: the worklog bar and the Workbooks top-line each keep one long-lived
        // pill in their AXAML and swap its text as the active workbook changes.
        //
        // Same visual as BuildStatePill, applied to controls someone else owns - so those two sites
        // cannot drift from the built ones without this method changing too.
        // ###########################################################################################
        public static void ApplyStatePillVisual(Border pill, TextBlock glyph, TextBlock label, string state)
        {
            bool isResolved = WorklogManager.IsResolvedState(state);
            var stateBrush = new SolidColorBrush(ResolveStateColor(state));

            pill.Background = ThemeResources.ResolveBrush("Form_Bg", Brushes.Transparent);
            pill.BorderBrush = stateBrush;
            pill.BorderThickness = InfoBorderThickness;
            pill.CornerRadius = PillCornerRadius;

            glyph.Text = WorklogGlyphs.GlyphFor(isResolved);
            glyph.Foreground = stateBrush;

            // Recomputed rather than set once at construction: the two padlocks overshoot their
            // font's declared ascent by different amounts, so a fixed padding is right for only one
            // of the two states - see FontAwesomeGlyphMetrics.
            glyph.Padding = FontAwesomeGlyphMetrics.GetTopOverflowThickness(
                WorklogGlyphs.CodepointFor(isResolved), glyph.FontSize);

            label.Text = state;
            label.Foreground = stateBrush;
            label.FontWeight = FontWeight.SemiBold;
        }

        // Closed is green, anything else - including an unrecognised future value - reads as
        // open/red, matching TabSchematics.Worklog.cs's own ResolveWorklogStateColor.
        public static Color ResolveStateColor(string state) =>
            ThemeResources.ResolveColor(
                WorklogManager.IsResolvedState(state) ? "Worklog_Status_Closed" : "Worklog_Status_Open",
                Colors.IndianRed);

        public static Color ResolveCategoryColor(string category) =>
            ThemeResources.ResolveColor($"Worklog_Category_{category}", Colors.IndianRed);

        // The neutral grey a COUNTED pill/chip takes when its count is zero - see BuildStatePill's
        // own note. Its OWN dedicated token, Workbooks_ZeroCount_Fg, rather than either of the
        // tab's body-text greys (Muted_Fg, then Faint_Fg, were each tried here and both reported as
        // "still looks black" on a small pill) - a zero count needs to read as washed-out/disabled,
        // a different and lighter register than muted PROSE needs to stay readable at. Resolved
        // fresh on every call rather than cached, matching every other brush here - ResolveBrush's
        // own two-step Application.Current + ThemeVariant lookup is what lets this follow a live
        // theme switch.
        private static IBrush MutedPillBrush =>
            ThemeResources.ResolveBrush("Workbooks_ZeroCount_Fg", Brushes.LightGray);

        private static TextBlock BuildGlyph(string text, FontFamily family, int codepoint, double fontSize, IBrush brush) => new()
        {
            Text = text,
            FontFamily = family,
            FontSize = fontSize,

            // See FontAwesomeGlyphMetrics: several of these glyphs are drawn taller than the font's
            // declared ascent, so without a reserved row their top pixel row is clipped. Computed
            // from the caller's own font size rather than hardcoded.
            Padding = FontAwesomeGlyphMetrics.GetTopOverflowThickness(codepoint, fontSize),

            Foreground = brush,
            VerticalAlignment = VerticalAlignment.Center
        };

        // Bold, unlike the label beside it: across the whole summary strip the numbers are what is
        // being reported and the words merely name them, so the numbers carry the weight.
        //
        // NOT bold when isMuted: a zero count is already de-emphasised by colour (MutedPillBrush),
        // and a BOLD zero in that same light grey still reads as solid/dark at this pill's small
        // size - reported as "the zero counts still look black". Regular weight lets the lighter
        // colour actually read as light.
        private static TextBlock BuildCountLabel(int count, double fontSize, IBrush brush, bool isMuted = false) => new()
        {
            Text = count.ToString(System.Globalization.CultureInfo.InvariantCulture),
            FontSize = fontSize,
            FontWeight = isMuted ? FontWeight.Normal : FontWeight.Bold,
            Foreground = brush,
            VerticalAlignment = VerticalAlignment.Center
        };

        private static TextBlock BuildLabel(string text, double fontSize, IBrush brush) => new()
        {
            Text = text,
            FontSize = fontSize,
            FontWeight = FontWeight.SemiBold,
            Foreground = brush,
            VerticalAlignment = VerticalAlignment.Center
        };
    }
}
