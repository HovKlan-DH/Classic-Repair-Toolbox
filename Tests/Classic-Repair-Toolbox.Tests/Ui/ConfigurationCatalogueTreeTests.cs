using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Media;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Covers ConfigurationCatalogueTree - the Configuration tab's hardware/board/schematic
// visibility checkbox tree. Built entirely in code (no DataTemplate), so these construct real
// Avalonia CheckBox/StackPanel controls and must run headlessly through UiTest.Run.
// ###########################################################################################
[Collection("HeadlessUi")]
public class ConfigurationCatalogueTreeTests
{
    private static List<HardwareBoardEntry> TwoHardwareTwoBoards() =>
    [
        new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/data.xlsx" },
        new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "KU-14194HB", ExcelDataFile = "Commodore/C64/KU-14194HB/data.xlsx" },
        new HardwareBoardEntry { HardwareName = "Commodore 128", BoardName = "310378", ExcelDataFile = "Commodore/C128/310378/data.xlsx" },
    ];

    private static Dictionary<string, List<string>> SchematicsFor(TwoHardwareTwoBoardsSchematics schematics) =>
        new()
        {
            ["Commodore 64|250407"] = schematics.Board250407,
            ["Commodore 64|KU-14194HB"] = schematics.BoardKu14194Hb,
            ["Commodore 128|310378"] = schematics.Board310378,
        };

    private sealed class TwoHardwareTwoBoardsSchematics
    {
        public List<string> Board250407 { get; } = new() { "Top" };
        public List<string> BoardKu14194Hb { get; } = new() { "Top (replica)", "Bottom (replica)", "Schematics (replica)" };
        public List<string> Board310378 { get; } = new() { "Top (rev. 7)" };
    }

    // Convenience overload for tests that do not care about collapse state - defaults to nothing
    // collapsed and a no-op collapse callback.
    private static ConfigurationCatalogueTree.BuildResult BuildTree(
        IEnumerable<HardwareBoardEntry> hardwareBoards,
        IReadOnlyDictionary<string, List<string>> schematicsByBoardKey,
        IReadOnlySet<string> uncheckedKeys,
        System.Action<string, bool> onCheckedChanged,
        IReadOnlySet<string>? collapsedKeys = null,
        System.Action<string, bool>? onCollapsedChanged = null) =>
        ConfigurationCatalogueTree.Build(
            hardwareBoards,
            schematicsByBoardKey,
            uncheckedKeys,
            collapsedKeys ?? new HashSet<string>(),
            onCheckedChanged,
            onCollapsedChanged ?? ((_, _) => { }));

    [Fact]
    public void Building_the_tree_creates_one_row_per_hardware_board_and_schematic()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            // Two hardware rows: Commodore 64, Commodore 128.
            Assert.Equal(2, result.HardwareRows.Count);

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.Equal(2, c64.Children.Count); // 250407, KU-14194HB

            var ku14194 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|KU-14194HB");
            Assert.Equal(3, ku14194.Children.Count); // three schematics

            var c128 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 128");
            var board310378 = Assert.Single(c128.Children);
            Assert.Equal("Commodore 128|310378", board310378.Key);
            var schematic = Assert.Single(board310378.Children);
            Assert.Equal("Commodore 128|310378|Top (rev. 7)", schematic.Key);
        });
    }

    // The tree deliberately does NOT sort hardware or boards alphabetically - it follows the
    // same source order as Main.PopulateHardwareDropDown/OnHardwareSelectionChanged, which is
    // DataManager.HardwareBoards' own row order (the master Excel's order), so the same
    // hardware/board sits in the same relative position in both the tree and the drop-downs.
    // Deliberately fed hardware and boards OUT of alphabetical order, so a stray .OrderBy would
    // fail this.
    [Fact]
    public void Hardware_and_boards_keep_the_order_they_were_given_in_rather_than_sorting_alphabetically()
    {
        UiTest.Run(() =>
        {
            List<HardwareBoardEntry> outOfAlphabeticalOrder =
            [
                new HardwareBoardEntry { HardwareName = "Commodore 128", BoardName = "310378", ExcelDataFile = "Commodore/C128/310378/data.xlsx" },
                new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "KU-14194HB", ExcelDataFile = "Commodore/C64/KU-14194HB/data.xlsx" },
                new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/data.xlsx" },
            ];

            var result = BuildTree(
                outOfAlphabeticalOrder, new Dictionary<string, List<string>>(), new HashSet<string>(), (_, _) => { });

            Assert.Equal(
                ["Commodore 128", "Commodore 64"],
                result.HardwareRows.Select(r => r.Key));

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.Equal(
                ["Commodore 64|KU-14194HB", "Commodore 64|250407"],
                c64.Children.Select(b => b.Key));
        });
    }

    [Fact]
    public void Rows_are_fully_expanded_by_default_with_no_collapse_state()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            // Every row was added directly to the root panel's Children - nothing is collapsed
            // or hidden behind an expander, so every row's control is visible.
            var rootPanel = Assert.IsType<StackPanel>(result.RootControl);
            Assert.All(rootPanel.Children, child => Assert.True(child.IsVisible));

            // 2 hardware + 3 boards (250407, KU-14194HB, 310378) + 5 schematics
            // (1 + 3 + 1) = 10 rows total.
            Assert.Equal(10, rootPanel.Children.Count);
        });
    }

    [Fact]
    public void Only_hardware_rows_are_bold()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.Equal(FontWeight.Bold, c64.LabelText.FontWeight);

            foreach (var boardRow in c64.Children)
            {
                Assert.Equal(FontWeight.Normal, boardRow.LabelText.FontWeight);
                foreach (var schematicRow in boardRow.Children)
                {
                    Assert.Equal(FontWeight.Normal, schematicRow.LabelText.FontWeight);
                }
            }
        });
    }

    // -------------------------------------------------------------- alert (IndianRed) state

    [Fact]
    public void A_checked_row_is_not_IndianRed()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.NotEqual(Brushes.IndianRed, c64.LabelText.Foreground);
        });
    }

    // The core rule from the request: a row whose OWN checkbox is unchecked, while it is still
    // ACTIVE (no ancestor already unchecked), shows IndianRed.
    [Fact]
    public void An_unchecked_and_active_row_is_IndianRed()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var unchecked_ = new HashSet<string> { "Commodore 64|250407" };

            var result = BuildTree(hardwareBoards, schematics, unchecked_, (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            var board250407 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|250407");

            Assert.Equal(Brushes.IndianRed, board250407.LabelText.Foreground);

            // The hardware row above it is still checked and active - no colour.
            Assert.NotEqual(Brushes.IndianRed, c64.LabelText.Foreground);

            // A sibling board, also checked - no colour.
            var ku14194 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|KU-14194HB");
            Assert.NotEqual(Brushes.IndianRed, ku14194.LabelText.Foreground);
        });
    }

    // The "too much distraction" rule: once a parent is unchecked (and shows IndianRed itself),
    // its now-disabled children do NOT also turn red, even though they are still individually
    // "unchecked" by the ancestor-AND visibility rule - CatalogueVisibility.IsEffectivelyVisible
    // says they are hidden, but their OWN stored checked state, and this label colour, are
    // untouched. Only the row that is itself both unchecked AND still active gets the colour.
    [Fact]
    public void A_disabled_descendant_of_an_unchecked_parent_does_not_also_turn_red()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var unchecked_ = new HashSet<string> { "Commodore 64" };

            var result = BuildTree(hardwareBoards, schematics, unchecked_, (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.Equal(Brushes.IndianRed, c64.LabelText.Foreground);

            // Boards and schematics underneath are disabled (not active) and stay checked in
            // storage, so none of them gets the alert colour despite the hardware above being red.
            foreach (var boardRow in c64.Children)
            {
                Assert.False(boardRow.RowControl.IsEnabled);
                Assert.NotEqual(Brushes.IndianRed, boardRow.LabelText.Foreground);

                foreach (var schematicRow in boardRow.Children)
                {
                    Assert.NotEqual(Brushes.IndianRed, schematicRow.LabelText.Foreground);
                }
            }
        });
    }

    // Re-checking the parent makes it active again - if a CHILD was already individually
    // unchecked underneath it, that child now becomes both unchecked AND active, and only NOW
    // does it turn red - it was there in storage the whole time, just not shown while inert.
    [Fact]
    public void Rechecking_a_parent_reveals_the_red_on_a_child_that_was_already_unchecked_underneath_it()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var unchecked_ = new HashSet<string> { "Commodore 64", "Commodore 64|250407" };

            var result = BuildTree(hardwareBoards, schematics, unchecked_, (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            var board250407 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|250407");

            // Still disabled under its unchecked (red) hardware, so no red of its own yet.
            Assert.NotEqual(Brushes.IndianRed, board250407.LabelText.Foreground);

            unchecked_.Remove("Commodore 64");
            ConfigurationCatalogueTree.ApplyEnabledState(result.HardwareRows, unchecked_);
            ConfigurationCatalogueTree.ApplyAlertState(result.HardwareRows, unchecked_);

            Assert.NotEqual(Brushes.IndianRed, c64.LabelText.Foreground); // hardware itself is checked again
            Assert.True(board250407.RowControl.IsEnabled);
            Assert.Equal(Brushes.IndianRed, board250407.LabelText.Foreground); // now active AND unchecked
        });
    }

    [Fact]
    public void Every_checkbox_starts_checked_when_nothing_is_unchecked()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            foreach (var hardwareRow in result.HardwareRows)
            {
                Assert.True(hardwareRow.CheckBox.IsChecked);
                foreach (var boardRow in hardwareRow.Children)
                {
                    Assert.True(boardRow.CheckBox.IsChecked);
                    foreach (var schematicRow in boardRow.Children)
                    {
                        Assert.True(schematicRow.CheckBox.IsChecked);
                    }
                }
            }
        });
    }

    [Fact]
    public void An_unchecked_hardware_key_is_reflected_as_unchecked_on_build()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var unchecked_ = new HashSet<string> { "Commodore 128" };

            var result = BuildTree(
                hardwareBoards, schematics, unchecked_, (_, _) => { });

            var c128 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 128");
            Assert.False(c128.CheckBox.IsChecked);
        });
    }

    // The core "disabled, not erased" rule: unchecking a hardware disables its boards' and
    // schematics' checkboxes, but does not touch their own IsChecked value.
    [Fact]
    public void Unchecking_a_hardware_disables_its_boards_and_schematics_without_changing_their_own_checked_state()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var unchecked_ = new HashSet<string>();

            var result = BuildTree(
                hardwareBoards, schematics, unchecked_, (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");

            // All descendants start enabled and checked.
            foreach (var boardRow in c64.Children)
            {
                Assert.True(boardRow.RowControl.IsEnabled);
                foreach (var schematicRow in boardRow.Children)
                {
                    Assert.True(schematicRow.RowControl.IsEnabled);
                }
            }

            // Simulate unchecking Commodore 64 (as the callback would do).
            unchecked_.Add("Commodore 64");
            ConfigurationCatalogueTree.ApplyEnabledState(result.HardwareRows, unchecked_);

            foreach (var boardRow in c64.Children)
            {
                Assert.False(boardRow.RowControl.IsEnabled);
                Assert.True(boardRow.CheckBox.IsChecked); // stored state untouched

                foreach (var schematicRow in boardRow.Children)
                {
                    Assert.False(schematicRow.RowControl.IsEnabled);
                    Assert.True(schematicRow.CheckBox.IsChecked); // stored state untouched
                }
            }

            // The hardware row's own row is never disabled by its own uncheck.
            Assert.True(c64.RowControl.IsEnabled);

            // Siblings (Commodore 128) are unaffected.
            var c128 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 128");
            foreach (var boardRow in c128.Children)
            {
                Assert.True(boardRow.RowControl.IsEnabled);
            }
        });
    }

    [Fact]
    public void Rechecking_a_hardware_reenables_its_descendants()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var unchecked_ = new HashSet<string> { "Commodore 64" };

            var result = BuildTree(
                hardwareBoards, schematics, unchecked_, (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.All(c64.Children, boardRow => Assert.False(boardRow.RowControl.IsEnabled));

            unchecked_.Remove("Commodore 64");
            ConfigurationCatalogueTree.ApplyEnabledState(result.HardwareRows, unchecked_);

            Assert.All(c64.Children, boardRow => Assert.True(boardRow.RowControl.IsEnabled));
            Assert.All(c64.Children, boardRow =>
                Assert.All(boardRow.Children, schematicRow => Assert.True(schematicRow.RowControl.IsEnabled)));
        });
    }

    // Unchecking a BOARD (not the hardware) disables only that board's own schematics, leaving
    // sibling boards under the same hardware untouched.
    [Fact]
    public void Unchecking_a_board_disables_only_its_own_schematics()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var unchecked_ = new HashSet<string> { "Commodore 64|KU-14194HB" };

            var result = BuildTree(
                hardwareBoards, schematics, unchecked_, (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            var ku14194 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|KU-14194HB");
            var board250407 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|250407");

            Assert.All(ku14194.Children, schematicRow => Assert.False(schematicRow.RowControl.IsEnabled));
            Assert.All(board250407.Children, schematicRow => Assert.True(schematicRow.RowControl.IsEnabled));

            // The board row's own row stays enabled (only descendants of an unchecked row
            // are disabled).
            Assert.True(ku14194.RowControl.IsEnabled);
        });
    }

    [Fact]
    public void Toggling_a_checkbox_invokes_the_callback_with_its_key_and_new_state()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            string? lastKey = null;
            bool? lastChecked = null;

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(),
                (key, isChecked) =>
                {
                    lastKey = key;
                    lastChecked = isChecked;
                });

            var c128 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 128");
            c128.CheckBox.IsChecked = false;

            Assert.Equal("Commodore 128", lastKey);
            Assert.False(lastChecked);
        });
    }

    // -------------------------------------------------------------- collapse / expand

    [Fact]
    public void Hardware_and_board_rows_get_a_chevron_but_leaf_schematic_rows_do_not()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.NotNull(c64.Chevron);

            var board250407 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|250407");
            Assert.NotNull(board250407.Chevron); // has one schematic, so still collapsible

            var schematic = Assert.Single(board250407.Children);
            Assert.Null(schematic.Chevron);
        });
    }

    [Fact]
    public void Nothing_is_collapsed_by_default()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.All(c64.Children, boardRow => Assert.True(boardRow.RowControl.IsVisible));
            Assert.All(c64.Children, boardRow =>
                Assert.All(boardRow.Children, schematicRow => Assert.True(schematicRow.RowControl.IsVisible)));
        });
    }

    // Collapsing a hardware hides its boards, and transitively its schematics too - regardless
    // of whether those boards are themselves marked collapsed.
    [Fact]
    public void Collapsing_a_hardware_hides_its_boards_and_their_schematics()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var collapsed = new HashSet<string>();

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(), (_, _) => { }, collapsed);

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");

            collapsed.Add("Commodore 64");
            ConfigurationCatalogueTree.ApplyCollapsedState(result.HardwareRows, collapsed);

            Assert.All(c64.Children, boardRow => Assert.False(boardRow.RowControl.IsVisible));
            Assert.All(c64.Children, boardRow =>
                Assert.All(boardRow.Children, schematicRow => Assert.False(schematicRow.RowControl.IsVisible)));

            // Siblings (Commodore 128) are unaffected.
            var c128 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 128");
            Assert.All(c128.Children, boardRow => Assert.True(boardRow.RowControl.IsVisible));

            // The hardware row itself is never hidden by its own collapse - only its children are.
            Assert.True(c64.RowControl.IsVisible);
        });
    }

    [Fact]
    public void Expanding_a_hardware_again_reveals_its_boards_and_schematics()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var collapsed = new HashSet<string> { "Commodore 64" };

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(), (_, _) => { }, collapsed);

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.All(c64.Children, boardRow => Assert.False(boardRow.RowControl.IsVisible));

            collapsed.Remove("Commodore 64");
            ConfigurationCatalogueTree.ApplyCollapsedState(result.HardwareRows, collapsed);

            Assert.All(c64.Children, boardRow => Assert.True(boardRow.RowControl.IsVisible));
            Assert.All(c64.Children, boardRow =>
                Assert.All(boardRow.Children, schematicRow => Assert.True(schematicRow.RowControl.IsVisible)));
        });
    }

    // Collapsing a BOARD (not the hardware) hides only that board's own schematics, leaving
    // sibling boards under the same hardware untouched - the same scoping rule as unchecking.
    [Fact]
    public void Collapsing_a_board_hides_only_its_own_schematics()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var collapsed = new HashSet<string> { "Commodore 64|KU-14194HB" };

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(), (_, _) => { }, collapsed);

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            var ku14194 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|KU-14194HB");
            var board250407 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|250407");

            Assert.All(ku14194.Children, schematicRow => Assert.False(schematicRow.RowControl.IsVisible));
            Assert.All(board250407.Children, schematicRow => Assert.True(schematicRow.RowControl.IsVisible));

            // The board row itself stays visible - only its own children are hidden.
            Assert.True(ku14194.RowControl.IsVisible);
        });
    }

    [Fact]
    public void A_collapsed_hardware_key_shows_the_chevron_right_glyph_and_an_expanded_one_shows_chevron_down()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());
            var collapsed = new HashSet<string> { "Commodore 64" };

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(), (_, _) => { }, collapsed);

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            var c128 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 128");

            // fa-solid chevron-right (U+F054) vs chevron-down (U+F078).
            Assert.Equal("", c64.Chevron!.Text);
            Assert.Equal("", c128.Chevron!.Text);
        });
    }

    // Drives the exact same logic the chevron's PointerPressed handler runs (via
    // Row.ToggleCollapsedForTests), rather than fabricating a headless pointer event on a control
    // that was never attached to a window.
    [Fact]
    public void Toggling_the_chevron_invokes_the_collapse_callback_and_never_the_checked_callback()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            string? lastCollapsedKey = null;
            bool? lastCollapsed = null;
            bool checkedCallbackFired = false;

            var result = BuildTree(
                hardwareBoards, schematics, new HashSet<string>(),
                (_, _) => checkedCallbackFired = true,
                onCollapsedChanged: (key, isCollapsed) =>
                {
                    lastCollapsedKey = key;
                    lastCollapsed = isCollapsed;
                });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            Assert.NotNull(c64.ToggleCollapsedForTests);

            c64.ToggleCollapsedForTests!();

            Assert.Equal("Commodore 64", lastCollapsedKey);
            Assert.True(lastCollapsed);
            Assert.False(checkedCallbackFired);
        });
    }

    [Fact]
    public void A_leaf_schematic_row_has_no_collapse_toggle()
    {
        UiTest.Run(() =>
        {
            var hardwareBoards = TwoHardwareTwoBoards();
            var schematics = SchematicsFor(new TwoHardwareTwoBoardsSchematics());

            var result = BuildTree(hardwareBoards, schematics, new HashSet<string>(), (_, _) => { });

            var c64 = Assert.Single(result.HardwareRows, r => r.Key == "Commodore 64");
            var board250407 = Assert.Single(c64.Children, b => b.Key == "Commodore 64|250407");
            var schematic = Assert.Single(board250407.Children);

            Assert.Null(schematic.ToggleCollapsedForTests);
        });
    }

    // The catalogue panel's splitter width follows the same restore-and-clamp pattern as
    // TabWorkbooks' splitters - see UserSettings.ConfigurationCataloguePanelWidth.
    [Fact]
    public void The_catalogue_panel_width_restores_from_settings_and_clamps_out_of_range_values()
    {
        UiTest.Run(() =>
        {
            UserSettings.ConfigurationCataloguePanelWidth = 500.0;
            var tab = new TabConfiguration();
            Assert.Equal(500.0, tab.CataloguePanelColumnWidthForTests);

            // Below the floor clamps up. The floor is 200, matching the catalogue column's own
            // MinWidth in the markup - the two have to agree, or a restored width the clamp
            // considers legal is silently widened again by the layout.
            UserSettings.ConfigurationCataloguePanelWidth = 10.0;
            tab.ApplyCatalogueSplitterWidthForTests();
            Assert.Equal(200.0, tab.CataloguePanelColumnWidthForTests);

            // Above the ceiling clamps down.
            UserSettings.ConfigurationCataloguePanelWidth = 5000.0;
            tab.ApplyCatalogueSplitterWidthForTests();
            Assert.Equal(900.0, tab.CataloguePanelColumnWidthForTests);

            // Restore a sane default so this scalar does not leak an out-of-range value into
            // whichever test runs next against the same static UserSettings state.
            UserSettings.ConfigurationCataloguePanelWidth = 320.0;
        });
    }

    // Nothing on this tab scrolls horizontally, so the split's declared minimums are a hard
    // budget: if they add up to more than the tab can ever be given, the catalogue panel on the
    // right is simply clipped, with no way for the user to reach it. The budget is what Main
    // leaves this tab at the window's own MinWidth="800" - its 200 DIP sidebar, the 4 DIP
    // splitter beside it, and this tab's own 16 DIP margin on each side.
    //
    // Shipped once at 360 + 4 + 260 = 624 against a budget of about 564.
    [Fact]
    public void The_configuration_split_minimum_widths_fit_the_narrowest_the_tab_can_get()
    {
        UiTest.Run(() =>
        {
            const double windowMinimumWidth = 800.0;
            const double mainSidebarWidth = 200.0;
            const double mainSplitterWidth = 4.0;
            const double tabHorizontalMargin = 16.0 * 2;

            double availableWidth =
                windowMinimumWidth - mainSidebarWidth - mainSplitterWidth - tabHorizontalMargin;

            var tab = new TabConfiguration();
            double requiredWidth = tab.SplitMinimumWidthForTests;

            Assert.True(
                requiredWidth <= availableWidth,
                $"Configuration split needs [{requiredWidth}] DIP but only [{availableWidth}] is available at the window's minimum width");
        });
    }
}
