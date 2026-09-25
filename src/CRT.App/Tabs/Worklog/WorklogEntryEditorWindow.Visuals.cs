using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Theming;
using Handlers.Geometry;
using System;

namespace CRT
{
    // ###########################################################################################
    // Category chip and state pill visual state (selected/unselected colouring), the schematic
    // location preview overlay, and the shared theme-brush/category-colour helpers they use. See
    // WorklogEntryEditorWindow.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class WorklogEntryEditorWindow
    {
        // ###########################################################################################
        // Resolves a theme brush by key, falling back when the resource cannot be found - same idiom
        // TabSchematics.ResolveThemeBrush uses, including its Application.Current fallback: this
        // window's ThemeVariant-keyed resources (Worklog_Category_*, Worklog_Status_Closed, etc.) live
        // in App.axaml's ResourceDictionary.ThemeDictionaries, and plain TryFindResource does not
        // always resolve a themed key by itself - without this second lookup every category chip and
        // state pill silently fell back to the caller's fallback color instead of its real one.
        // ###########################################################################################
        private IBrush ResolveThemeBrush(string key, IBrush fallback) =>
            ThemeResources.ResolveForControl(this, key, fallback);

        private Color ResolveCategoryColor(string category)
        {
            var brush = this.ResolveThemeBrush($"Worklog_Category_{category}", new SolidColorBrush(Colors.IndianRed));
            return brush is ISolidColorBrush solidBrush ? solidBrush.Color : Colors.IndianRed;
        }

        // ###########################################################################################
        // Draws the entry's marked-area rectangle over the (fully visible, unzoomed) schematic
        // preview image on the right - a static reference showing where on the board this entry
        // applies, not an interactive viewer.
        // ###########################################################################################
        private void RefreshLocationPreviewOverlay()
        {
            this.EditorLocationPreviewOverlayCanvas.Children.Clear();

            if (this.thisSchematicBitmap == null)
                return;

            var controlSize = this.EditorLocationPreviewGrid.Bounds.Size;
            if (controlSize.Width <= 0 || controlSize.Height <= 0)
                return;

            // Centered, not origin-anchored: EditorLocationPreviewImage is Stretch="Uniform" with no
            // alignment set, so Avalonia centres the content in the fixed-height preview box. Using
            // the origin-anchored GetImageContentRect drew the marker off by half the letterbox -
            // pointing at the wrong part of the board, and clipped away entirely on tall schematics.
            var contentRect = RectGeometry.GetCenteredImageContentRect(controlSize, this.thisSchematicBitmap.PixelSize);
            var pixelRect = new Rect(this.thisEntry.AreaX, this.thisEntry.AreaY, this.thisEntry.AreaWidth, this.thisEntry.AreaHeight);
            var localRect = RectGeometry.PixelToLocalRect(pixelRect, contentRect, this.thisSchematicBitmap.PixelSize);

            var color = this.ResolveCategoryColor(this.thisSelectedCategory);

            var marker = new Avalonia.Controls.Shapes.Rectangle
            {
                Width = Math.Max(1, localRect.Width),
                Height = Math.Max(1, localRect.Height),
                Fill = new SolidColorBrush(color, 0.18),
                Stroke = new SolidColorBrush(color, 1.0),
                StrokeThickness = 2,
                StrokeDashArray = new Avalonia.Collections.AvaloniaList<double> { 4, 3 }
            };

            Canvas.SetLeft(marker, localRect.X);
            Canvas.SetTop(marker, localRect.Y);
            this.EditorLocationPreviewOverlayCanvas.Children.Add(marker);
        }

        // ###########################################################################################
        // Clicking a category chip records the change as an automatic comment, so the entry carries
        // its own history rather than only its current category.
        //
        // Clicking the ALREADY-selected chip records nothing: it is not a change, and treating it as
        // one would let a user fill the comment list by clicking the same chip repeatedly.
        // ###########################################################################################
        private void OnEditorCategoryChipPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Border { Tag: string category })
                return;

            if (string.Equals(this.thisSelectedCategory, category, StringComparison.Ordinal))
                return;

            this.thisSelectedCategory = category;
            this.UpdateCategoryChipVisuals();
            this.RefreshLocationPreviewOverlay();
            this.MarkDirty();

            this.RecordAutomaticComment(WorklogManager.BuildCategoryChangedCommentText(category));
        }

        private void UpdateCategoryChipVisuals()
        {
            this.ApplyCategoryChipVisualState(this.EditorCategoryNoteChip, this.EditorCategoryNoteText, this.EditorCategoryNoteIcon, "Note");
            this.ApplyCategoryChipVisualState(this.EditorCategoryCosmeticChip, this.EditorCategoryCosmeticText, this.EditorCategoryCosmeticIcon, "Cosmetic");
            this.ApplyCategoryChipVisualState(this.EditorCategoryIssueChip, this.EditorCategoryIssueText, this.EditorCategoryIssueIcon, "Issue");

            this.EditorIdBadge.Background = new SolidColorBrush(this.ResolveCategoryColor(this.thisSelectedCategory));
        }

        // The icon takes the label's colour rather than a colour of its own - white on the selected
        // chip's filled background, the ordinary foreground otherwise. An icon left at one fixed
        // colour would either disappear into the fill or stay dark while its own label went white.
        private void ApplyCategoryChipVisualState(Border chip, TextBlock label, TextBlock icon, string category)
        {
            var categoryBrush = this.ResolveThemeBrush($"Worklog_Category_{category}", new SolidColorBrush(Colors.IndianRed));

            if (string.Equals(this.thisSelectedCategory, category, StringComparison.Ordinal))
            {
                chip.Background = categoryBrush;
                chip.BorderBrush = categoryBrush;
                chip.BorderThickness = new Thickness(2);
                chip.Opacity = 0.9;
                label.Foreground = Brushes.White;
                label.FontWeight = FontWeight.SemiBold;
                icon.Foreground = label.Foreground;
            }
            else
            {
                chip.Background = this.ResolveThemeBrush("Form_Bg", new SolidColorBrush(Color.Parse("#F5F5F5")));
                chip.BorderBrush = this.ResolveThemeBrush("Form_Border", new SolidColorBrush(Color.Parse("#CCCCCC")));
                chip.BorderThickness = new Thickness(1);
                chip.Opacity = 1.0;
                label.Foreground = this.ResolveThemeBrush("Schematics_Panels_Fg", Brushes.Black);
                label.FontWeight = FontWeight.Normal;
                icon.Foreground = label.Foreground;
            }
        }

        // Clicking the already-selected pill records nothing - see the category handler above.
        private void OnEditorStatePillPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Border { Tag: string state })
                return;

            if (string.Equals(this.thisSelectedState, state, StringComparison.Ordinal))
                return;

            this.thisSelectedState = state;
            this.UpdateStatePillVisuals();
            this.MarkDirty();

            this.RecordAutomaticComment(WorklogManager.BuildStateChangedCommentText(state));
        }

        // Reserves the top pixel row the padlocks need, computed from each control's own font size
        // rather than hardcoded in markup - see FontAwesomeGlyphMetrics for why the literal form is
        // a clipped icon waiting for a font-size change.
        private static void ApplyFontAwesomeOverflowPadding(params TextBlock[] icons)
        {
            foreach (var icon in icons)
            {
                icon.Padding = FontAwesomeGlyphMetrics.GetTopOverflowThicknessForText(icon.Text, icon.FontSize);
            }
        }

        private void UpdateStatePillVisuals()
        {
            ApplyFontAwesomeOverflowPadding(this.EditorStateOpenDot, this.EditorStateClosedDot);

            this.ApplyStatePillVisualState(this.EditorStateOpenPill, this.EditorStateOpenText, this.EditorStateOpenDot, "Open", "Worklog_Status_Open");
            this.ApplyStatePillVisualState(this.EditorStateClosedPill, this.EditorStateClosedText, this.EditorStateClosedDot, "Closed", "Worklog_Status_Closed");
        }

        // The SELECTED pill is filled with its state colour and its label goes white and bold -
        // the same treatment the category chips use. It was outline-only, which on the pale
        // Schematics_Panels_Bg left "selected" and "unselected" separated by little more than a
        // 1px border-width difference, and the selected pill was genuinely hard to pick out.
        //
        // The padlock keeps its state colour in the UNSELECTED pill (it is the state's identity, not
        // a selection cue) but turns white in the selected one, where the fill already carries the
        // colour and a coloured glyph on a same-coloured fill would simply vanish.
        private void ApplyStatePillVisualState(Border pill, TextBlock label, TextBlock icon, string state, string colorResourceKey)
        {
            var stateBrush = this.ResolveThemeBrush(colorResourceKey, new SolidColorBrush(Colors.IndianRed));

            if (string.Equals(this.thisSelectedState, state, StringComparison.Ordinal))
            {
                pill.Background = stateBrush;
                pill.BorderBrush = stateBrush;
                pill.BorderThickness = new Thickness(2);
                pill.Opacity = 0.9;
                icon.Foreground = Brushes.White;
                label.Foreground = Brushes.White;
                label.FontWeight = FontWeight.SemiBold;
            }
            else
            {
                pill.Background = this.ResolveThemeBrush("Form_Bg", new SolidColorBrush(Color.Parse("#F5F5F5")));
                pill.BorderBrush = this.ResolveThemeBrush("Form_Border", new SolidColorBrush(Color.Parse("#CCCCCC")));
                pill.BorderThickness = new Thickness(1);
                pill.Opacity = 1.0;
                icon.Foreground = stateBrush;
                label.Foreground = this.ResolveThemeBrush("Schematics_Panels_Fg", Brushes.Black);
                label.FontWeight = FontWeight.Normal;
            }
        }
    }
}
