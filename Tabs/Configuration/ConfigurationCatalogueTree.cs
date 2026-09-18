using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // Builds the Configuration tab's hardware/board/schematic visibility tree: a flat set of
    // nested StackPanels of CheckBox rows, generated in code rather than as a DataTemplate - same
    // reason the Workbooks tab builds its cards/pills in code (Handlers/Theme brush lookups need
    // Application.Current + ThemeVariant, which a template binding cannot express).
    //
    // There is no TreeView anywhere in this codebase, and this dataset (dozens of hardware, low
    // hundreds of boards/schematics total) has no need for virtualization, so a plain indented
    // StackPanel list is simplest.
    //
    // Every hardware/board row (anything with children) carries a collapse/expand chevron, the same
    // fa-solid chevron-right/chevron-down recipe as TabWorkbooks.Summary.cs's summary strip toggle -
    // same codepoints, same per-state overflow padding via FontAwesomeGlyphMetrics. Collapsed state
    // persists per row in UserSettings.CatalogueCollapsedKeys, keyed the same way as the checked
    // state; unlisted rows default to expanded.
    // ###########################################################################################
    internal static class ConfigurationCatalogueTree
    {
        private const double IndentPerLevel = 18.0;

        // Same fixed checkbox box size and font size as the Schematics tab's "Global settings"
        // panel (see TabSchematics.axaml's GlobalSettingsListPanel rows) - a Viewbox-wrapped
        // CheckBox with MinHeight/Margin/Padding zeroed out, otherwise the default Avalonia
        // CheckBox template's generous hit-target padding makes a long list far taller than it
        // needs to be.
        private const double CheckBoxBoxSize = 16.0;
        private const double RowLabelFontSize = 11.0;
        private const double RowSpacing = 6.0;

        // fa-solid chevron-right (collapsed) / chevron-down (expanded) - same glyphs and codepoints
        // as TabWorkbooks.Summary.cs's own toggle.
        private const int ChevronRightCodepoint = 0xF054;
        private const int ChevronDownCodepoint = 0xF078;
        private const double ChevronFontSize = 9.0;

        // The chevron sits in a fixed box of its own, same width for a row that has one and a row
        // that does not (a leaf schematic), so labels at the same depth still line up regardless of
        // whether their row is collapsible. This is the LAYOUT size (spacing to the checkbox next
        // to it); the actual click target is wider - see ChevronHitTestSize below.
        private const double ChevronBoxSize = 12.0;

        // The chevron's hit-test area is bigger than its layout box: a 12px glyph is an easy miss,
        // and a near-miss used to fall through to the row underneath and toggle the checkbox
        // instead - reported directly. Overlaying a wider, centred, transparent hit-test control
        // gives the chevron a generous click target without widening the space it actually reserves
        // in the row (which would push the checkbox and label further right for every row).
        private const double ChevronHitTestSize = 24.0;

        // ###########################################################################################
        // One row's controls, kept together so a checkbox change can walk to its descendant rows
        // (to enable/disable them), a chevron toggle can walk to its own children (to show/hide
        // them), and a rebuild can find/reuse nothing - the tree is rebuilt wholesale on data
        // changes and only enable/collapsed state is patched afterwards.
        // ###########################################################################################
        internal sealed class Row
        {
            public required string Key { get; init; }
            public required Control RowControl { get; init; }
            public required CheckBox CheckBox { get; init; }
            public required TextBlock LabelText { get; init; }
            public TextBlock? Chevron { get; init; }
            public required List<Row> Children { get; init; }

            // Test seam: invokes exactly the same logic the chevron's own PointerPressed handler
            // runs, without fabricating a headless pointer event on a control that is not attached
            // to a window. Null for a leaf row with no chevron.
            internal Action? ToggleCollapsedForTests { get; init; }
        }

        internal sealed class BuildResult
        {
            public required Control RootControl { get; init; }
            public required List<Row> HardwareRows { get; init; }
        }

        // ###########################################################################################
        // Builds the full tree. schematicsByBoardKey maps "HardwareName|BoardName" (the same
        // composite Main.GetCurrentBoardKey format) to that board's schematic names, already
        // resolved by the caller (loading board data is async; this method is not).
        //
        // onCheckedChanged is called with the row's key and its new checked state whenever a
        // checkbox is toggled by the user - the caller persists it via
        // UserSettings.SetCatalogueKeyChecked and re-applies enable state.
        //
        // onCollapsedChanged is the same idea for the chevron: called with the row's key and its new
        // collapsed state, so the caller can persist it via UserSettings.SetCatalogueKeyCollapsed.
        // Applying the resulting IsVisible change happens here, not in the caller, because collapsing
        // is purely a display concern - unlike a checkbox change it never affects what the rest of
        // the app shows, so it needs no re-filter callback into Main.
        // ###########################################################################################
        public static BuildResult Build(
            IEnumerable<HardwareBoardEntry> hardwareBoards,
            IReadOnlyDictionary<string, List<string>> schematicsByBoardKey,
            IReadOnlySet<string> uncheckedKeys,
            IReadOnlySet<string> collapsedKeys,
            Action<string, bool> onCheckedChanged,
            Action<string, bool> onCollapsedChanged)
        {
            var root = new StackPanel { Spacing = 0 };
            var hardwareRows = new List<Row>();

            // Deliberately NOT re-sorted alphabetically: GroupBy preserves each group's first-seen
            // order, which keeps the tree in the same order as DataManager.HardwareBoards - the
            // same source, in the same order, that Main.PopulateHardwareDropDown and
            // OnHardwareSelectionChanged build the Hardware/Board drop-downs from. A tree sorted
            // differently from those drop-downs would make the same hardware/board hard to find in
            // one place after finding it in the other.
            var boardsByHardware = hardwareBoards
                .Where(entry => !string.IsNullOrWhiteSpace(entry.HardwareName) && !string.IsNullOrWhiteSpace(entry.BoardName))
                .GroupBy(entry => entry.HardwareName, StringComparer.OrdinalIgnoreCase);

            foreach (var hardwareGroup in boardsByHardware)
            {
                string hardwareName = hardwareGroup.Key;
                string hardwareKey = CatalogueVisibility.BuildKey(hardwareName);

                // A hardware group always has at least one board (GroupBy would not have produced
                // the group otherwise), so it always gets a chevron.
                var hardwareRow = BuildRow(
                    hardwareKey, hardwareName, depth: 0, hasChildren: true,
                    uncheckedKeys, collapsedKeys, onCheckedChanged, onCollapsedChanged);
                root.Children.Add(hardwareRow.RowControl);

                foreach (var boardEntry in hardwareGroup)
                {
                    string boardName = boardEntry.BoardName;
                    string boardKey = CatalogueVisibility.BuildKey(hardwareName, boardName);

                    string boardLookupKey = $"{hardwareName}|{boardName}";
                    bool boardHasSchematics =
                        schematicsByBoardKey.TryGetValue(boardLookupKey, out var schematicNames) &&
                        schematicNames.Count > 0;

                    var boardRow = BuildRow(
                        boardKey, boardName, depth: 1, hasChildren: boardHasSchematics,
                        uncheckedKeys, collapsedKeys, onCheckedChanged, onCollapsedChanged);
                    root.Children.Add(boardRow.RowControl);
                    hardwareRow.Children.Add(boardRow);

                    if (boardHasSchematics)
                    {
                        foreach (string schematicName in schematicNames!)
                        {
                            string schematicKey = CatalogueVisibility.BuildKey(hardwareName, boardName, schematicName);

                            var schematicRow = BuildRow(
                                schematicKey, schematicName, depth: 2, hasChildren: false,
                                uncheckedKeys, collapsedKeys, onCheckedChanged, onCollapsedChanged);
                            root.Children.Add(schematicRow.RowControl);
                            boardRow.Children.Add(schematicRow);
                        }
                    }
                }

                hardwareRows.Add(hardwareRow);
            }

            ApplyEnabledState(hardwareRows, uncheckedKeys);
            ApplyCollapsedState(hardwareRows, collapsedKeys);
            ApplyAlertState(hardwareRows, uncheckedKeys);

            return new BuildResult { RootControl = root, HardwareRows = hardwareRows };
        }

        // ###########################################################################################
        // Walks every row and disables (IsEnabled = false, which also fades it - the built-in
        // CheckBox/TextBlock disabled opacity) any whose parent chain is unchecked, without touching
        // IsChecked - "disabled, not erased" from the feature request. A row's own checkbox stays
        // enabled regardless of its own checked state; only its DESCENDANTS become unreachable once
        // it is unchecked. Disabling the ROW control (not just the checkbox) is what actually blocks
        // the click, since the checkbox itself is IsHitTestVisible="False" - see BuildRow.
        // ###########################################################################################
        public static void ApplyEnabledState(IEnumerable<Row> hardwareRows, IReadOnlySet<string> uncheckedKeys)
        {
            foreach (var hardwareRow in hardwareRows)
            {
                bool hardwareChecked = !CatalogueVisibility.IsKeyListed(uncheckedKeys, hardwareRow.Key);

                foreach (var boardRow in hardwareRow.Children)
                {
                    boardRow.RowControl.IsEnabled = hardwareChecked;
                    bool boardChecked = hardwareChecked && !CatalogueVisibility.IsKeyListed(uncheckedKeys, boardRow.Key);

                    foreach (var schematicRow in boardRow.Children)
                    {
                        schematicRow.RowControl.IsEnabled = boardChecked;
                    }
                }
            }
        }

        // ###########################################################################################
        // Walks every row and marks its label IndianRed when that row's OWN checkbox is unchecked
        // AND the row is still active (enabled) - i.e. no ancestor is already unchecked. Asked for
        // directly, to draw the eye to an unchecked item without colouring an entire disabled
        // subtree red, which would be "too much distraction": an unchecked hardware shows red once,
        // and its now-disabled (greyed out) boards and schematics underneath do not each get their
        // own red on top of that fade - the disabled opacity already marks them as inert.
        //
        // Must run AFTER ApplyEnabledState, since it reads RowControl.IsEnabled rather than
        // recomputing the ancestor chain itself - the two would otherwise have to agree on the same
        // logic twice.
        // ###########################################################################################
        public static void ApplyAlertState(IEnumerable<Row> hardwareRows, IReadOnlySet<string> uncheckedKeys)
        {
            foreach (var hardwareRow in hardwareRows)
            {
                ApplyAlertStateForRow(hardwareRow, uncheckedKeys);
            }
        }

        private static void ApplyAlertStateForRow(Row row, IReadOnlySet<string> uncheckedKeys)
        {
            bool isUnchecked = CatalogueVisibility.IsKeyListed(uncheckedKeys, row.Key);
            bool isActive = row.RowControl.IsEnabled;

            if (isUnchecked && isActive)
            {
                row.LabelText.Foreground = Brushes.IndianRed;
            }
            else
            {
                row.LabelText.ClearValue(TextBlock.ForegroundProperty);
            }

            foreach (var child in row.Children)
            {
                ApplyAlertStateForRow(child, uncheckedKeys);
            }
        }

        // ###########################################################################################
        // Walks every row and shows/hides each one's CHILDREN (never the row itself) to match its
        // own collapsed state - collapsing a hardware hides its boards, which in turn hides their
        // schematics regardless of the boards' own collapsed state, since a hidden parent already
        // hides everything under it. Also points each chevron the right way.
        //
        // Top-level hardware rows are always themselves visible (nothing collapses a hardware row
        // out of the list), so the walk starts with parentVisible: true for each of them.
        // ###########################################################################################
        public static void ApplyCollapsedState(IEnumerable<Row> hardwareRows, IReadOnlySet<string> collapsedKeys)
        {
            foreach (var hardwareRow in hardwareRows)
            {
                ApplyCollapsedStateForRow(hardwareRow, collapsedKeys, parentVisible: true);
            }
        }

        // ###########################################################################################
        // parentVisible carries whether ROW ITSELF is currently shown - false when an ancestor
        // further up is collapsed. A row's children are visible only when both the row is visible
        // AND the row is not itself collapsed; that combined flag is then threaded down as the
        // children's OWN parentVisible, so a grandchild under a collapsed board stays hidden even
        // after that board's hardware is later expanded on its own.
        // ###########################################################################################
        private static void ApplyCollapsedStateForRow(Row row, IReadOnlySet<string> collapsedKeys, bool parentVisible)
        {
            bool isCollapsed = CatalogueVisibility.IsKeyListed(collapsedKeys, row.Key);
            SetChevronCollapsed(row.Chevron, isCollapsed);

            bool childrenVisible = parentVisible && !isCollapsed;

            foreach (var child in row.Children)
            {
                child.RowControl.IsVisible = childrenVisible;
                ApplyCollapsedStateForRow(child, collapsedKeys, childrenVisible);
            }
        }

        private static void SetChevronCollapsed(TextBlock? chevron, bool isCollapsed)
        {
            if (chevron == null)
            {
                return;
            }

            int codepoint = isCollapsed ? ChevronRightCodepoint : ChevronDownCodepoint;
            chevron.Text = char.ConvertFromUtf32(codepoint);

            // Recomputed per state like TabWorkbooks.Summary.cs's own chevron - harmless here since
            // neither codepoint currently overshoots its declared ascent, but it stays correct if
            // that ever changes rather than silently clipping.
            chevron.Padding = FontAwesomeGlyphMetrics.GetTopOverflowThickness(codepoint, ChevronFontSize);
        }

        // ###########################################################################################
        // Builds one condensed row: an optional collapse chevron, a small Viewbox-wrapped CheckBox,
        // and a label - the same visual recipe as TabSchematics.axaml's "Global settings" panel rows
        // for the checkbox (MinHeight/Margin/Padding zeroed, sized down via the Viewbox rather than a
        // font-size trick so the tick mark stays crisp), and the same chevron recipe as
        // TabWorkbooks.Summary.cs's summary toggle.
        //
        // The checkbox and label are IsHitTestVisible="False" and the row toggles the checkbox on
        // PointerPressed, so clicking the label is as good as clicking the box - same as the
        // Schematics panel. The chevron is a SEPARATE control with its own PointerPressed that marks
        // the event handled, so a click on it toggles collapse without also toggling the checkbox
        // underneath (the row's handler never sees a handled event).
        // ###########################################################################################
        private static Row BuildRow(
            string key,
            string label,
            int depth,
            bool hasChildren,
            IReadOnlySet<string> uncheckedKeys,
            IReadOnlySet<string> collapsedKeys,
            Action<string, bool> onCheckedChanged,
            Action<string, bool> onCollapsedChanged)
        {
            var checkBox = new CheckBox
            {
                MinHeight = 0,
                Margin = new Thickness(0),
                Padding = new Thickness(0),
                CornerRadius = new CornerRadius(2),
                IsChecked = !CatalogueVisibility.IsKeyListed(uncheckedKeys, key),
                IsHitTestVisible = false
            };

            var checkBoxBox = new Viewbox
            {
                Width = CheckBoxBoxSize,
                Height = CheckBoxBoxSize,
                IsHitTestVisible = false,
                Child = checkBox
            };

            var labelText = new TextBlock
            {
                Text = label,
                FontSize = RowLabelFontSize,
                // Depth 0 is the hardware row, bolded so it reads as the group heading for the
                // boards and schematics indented underneath it.
                FontWeight = depth == 0 ? FontWeight.Bold : FontWeight.Normal,
                VerticalAlignment = VerticalAlignment.Center,
                IsHitTestVisible = false
            };

            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = RowSpacing,
                Background = Brushes.Transparent,
                Cursor = new Cursor(StandardCursorType.Hand),
                Margin = new Thickness(depth * IndentPerLevel, 1, 0, 1)
            };

            TextBlock? chevron = null;
            Action? toggleCollapsed = null;

            // Always reserves the chevron's box width, even for a leaf row with none, so labels at
            // the same depth line up whether or not their row happens to be collapsible. This is a
            // plain Panel rather than the old Border-as-hit-target: the actual click target is now
            // the wider, overlaid ChevronHitTestSize Border added below, which is free to be bigger
            // than this slot without disturbing anything else in the row's layout.
            var chevronSlot = new Panel { Width = ChevronBoxSize, Height = ChevronBoxSize };

            if (hasChildren)
            {
                chevron = new TextBlock
                {
                    FontFamily = ThemeResources.ResolveFontAwesomeSolid(),
                    FontSize = ChevronFontSize,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    IsHitTestVisible = false
                };

                chevronSlot.Children.Add(chevron);

                toggleCollapsed = () => onCollapsedChanged(key, !CatalogueVisibility.IsKeyListed(collapsedKeys, key));

                // A wider, centred, transparent hit-test area layered over the slot via a negative
                // margin - ZIndex keeps it above the label/checkbox so it wins the hit test even
                // though it visually overlaps them, and it alone carries the click handler and the
                // Hand cursor, so the small visible chevron is not what determines the click target.
                double overhang = (ChevronHitTestSize - ChevronBoxSize) / 2.0;
                var chevronHitTestArea = new Border
                {
                    Width = ChevronHitTestSize,
                    Height = ChevronHitTestSize,
                    Margin = new Thickness(-overhang, -overhang, -overhang, -overhang),
                    Background = Brushes.Transparent,
                    Cursor = new Cursor(StandardCursorType.Hand),
                    ZIndex = 1
                };

                chevronHitTestArea.PointerPressed += (_, e) =>
                {
                    toggleCollapsed();
                    e.Handled = true;
                };

                chevronSlot.Children.Add(chevronHitTestArea);
            }

            row.Children.Add(chevronSlot);
            row.Children.Add(checkBoxBox);
            row.Children.Add(labelText);

            row.PointerPressed += (_, e) =>
            {
                if (!row.IsEnabled)
                {
                    return;
                }

                checkBox.IsChecked = !(checkBox.IsChecked == true);
                e.Handled = true;
            };

            checkBox.IsCheckedChanged += (_, _) => onCheckedChanged(key, checkBox.IsChecked == true);

            return new Row
            {
                Key = key,
                RowControl = row,
                CheckBox = checkBox,
                LabelText = labelText,
                Chevron = chevron,
                ToggleCollapsedForTests = toggleCollapsed,
                Children = new List<Row>()
            };
        }
    }
}
