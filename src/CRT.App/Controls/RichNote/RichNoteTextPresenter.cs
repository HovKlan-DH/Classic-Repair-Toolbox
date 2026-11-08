using Avalonia;
using Avalonia.Controls.Presenters;
using Avalonia.Media;
using Avalonia.Media.TextFormatting;
using Avalonia.Utilities;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // THE TEXT OF A RichNoteTextBox, DRAWN IN ITS STYLES - bold, italic, underlined, struck through - in the very layout
    // the box's caret, selection and clicks are measured against (owner request, 2026-11-07: "see it directly in the
    // text"). Avalonia's TextPresenter lays its text out in one style; its CreateTextLayout is the one place a layout is
    // made, so this makes it with the styles as overrides of the runs. A bold word is wider, and the caret after it is
    // where the layout says - nothing drawn over the text to drift from it.
    //
    // Measured as the base measures (its constraint is private, so it is kept here as the base keeps it: set by Measure,
    // replaced by Arrange at a new width). While text is being composed (IME), the base draws it: the composition stands
    // between the styled characters and would shift them.
    // ###########################################################################################
    public class RichNoteTextPresenter : TextPresenter
    {
        private IReadOnlyList<RichNoteStyle> thisStyles = Array.Empty<RichNoteStyle>();

        private Size thisConstraint;

        // Each character's style, as the box holds them - drawn at once.
        public IReadOnlyList<RichNoteStyle> CharacterStyles
        {
            get => this.thisStyles;
            set
            {
                this.thisStyles = value ?? Array.Empty<RichNoteStyle>();
                this.InvalidateTextLayout();
            }
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            this.thisConstraint = availableSize;
            return base.MeasureOverride(availableSize);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            // The base lays its text out again at the width it is given (never narrower than its text); its constraint
            // is followed here first, so the layout made for it is made at that width.
            if (!AreClose(this.thisConstraint.Width, finalSize.Width))
            {
                this.thisConstraint = new Size(Math.Ceiling(finalSize.Width), double.PositiveInfinity);
            }

            return base.ArrangeOverride(finalSize);
        }

        protected override TextLayout CreateTextLayout()
        {
            string text = this.Text ?? string.Empty;

            if (!string.IsNullOrEmpty(this.PreeditText) || this.PasswordChar != '\0' || this.thisStyles.Count != text.Length || !HasStyle(this.thisStyles))
            {
                return base.CreateTextLayout();
            }

            var typeface = new Typeface(this.FontFamily, this.FontStyle, this.FontWeight, this.FontStretch);
            int selectionStart = Math.Min(this.SelectionStart, this.SelectionEnd);
            int selectionEnd = Math.Max(this.SelectionStart, this.SelectionEnd);
            IBrush? selectionForeground = this.ShowSelectionHighlight && selectionEnd > selectionStart ? this.SelectionForegroundBrush : null;

            var overrides = new List<ValueSpan<TextRunProperties>>();
            int start = 0;

            for (int index = 1; index <= text.Length; index++)
            {
                bool boundary = index == text.Length ||
                                this.thisStyles[index] != this.thisStyles[start] ||
                                (selectionForeground is not null && (index == selectionStart || index == selectionEnd));

                if (!boundary)
                {
                    continue;
                }

                RichNoteStyle style = this.thisStyles[start];
                bool selected = selectionForeground is not null && start >= selectionStart && start < selectionEnd;

                if (style != RichNoteStyle.None || selected)
                {
                    overrides.Add(new ValueSpan<TextRunProperties>(start, index - start, new GenericTextRunProperties(
                        StyledTypeface(typeface, style),
                        this.FontSize,
                        DecorationsOf(style),
                        selected ? selectionForeground : this.Foreground,
                        null,
                        BaselineAlignment.Baseline,
                        null,
                        this.FontFeatures)));
                }

                start = index;
            }

            double maxWidth = IsZero(this.thisConstraint.Width) ? double.PositiveInfinity : this.thisConstraint.Width;
            double maxHeight = IsZero(this.thisConstraint.Height) ? double.PositiveInfinity : this.thisConstraint.Height;

            return new TextLayout(
                text,
                typeface,
                this.FontSize,
                this.Foreground,
                this.TextAlignment,
                this.TextWrapping,
                null,
                null,
                this.FlowDirection,
                maxWidth,
                maxHeight,
                this.LineHeight,
                this.LetterSpacing,
                0,
                this.FontFeatures,
                overrides);
        }

        // As Avalonia's own MathUtilities compares (internal there): equal within a hair, and zero as near as makes none.
        private static bool AreClose(double left, double right) =>
            left == right || (double.IsFinite(left) && double.IsFinite(right) && Math.Abs(left - right) < 1.5E-7 * (Math.Abs(left) + Math.Abs(right) + 10.0));

        private static bool IsZero(double value) => Math.Abs(value) < 2.2204460492503131E-15 * 10.0;

        private static bool HasStyle(IReadOnlyList<RichNoteStyle> styles)
        {
            for (int index = 0; index < styles.Count; index++)
            {
                if (styles[index] != RichNoteStyle.None)
                {
                    return true;
                }
            }

            return false;
        }

        internal static Typeface StyledTypeface(Typeface typeface, RichNoteStyle style) => new(
            typeface.FontFamily,
            (style & RichNoteStyle.Italic) != 0 ? FontStyle.Italic : typeface.Style,
            (style & RichNoteStyle.Bold) != 0 ? FontWeight.Bold : typeface.Weight,
            typeface.Stretch);

        // The lines a style draws through or under its words - none for none.
        internal static TextDecorationCollection? DecorationsOf(RichNoteStyle style)
        {
            bool underline = (style & RichNoteStyle.Underline) != 0;
            bool strike = (style & RichNoteStyle.Strike) != 0;

            if (!underline && !strike)
            {
                return null;
            }

            var decorations = new TextDecorationCollection();

            if (underline)
            {
                foreach (TextDecoration line in TextDecorations.Underline)
                {
                    decorations.Add(line);
                }
            }

            if (strike)
            {
                foreach (TextDecoration line in TextDecorations.Strikethrough)
                {
                    decorations.Add(line);
                }
            }

            return decorations;
        }
    }
}
