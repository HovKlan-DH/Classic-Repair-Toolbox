using System.Collections.Generic;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for CatalogueVisibility - the Configuration tab's hardware/board/
// schematic checkbox tree, and the AND-with-ancestors rule that decides what the rest of the
// app actually shows.
public class CatalogueVisibilityTests
{
    // -------------------------------------------------------------- BuildKey

    [Fact]
    public void BuildKey_joins_segments_with_a_pipe()
    {
        Assert.Equal("Commodore 64", CatalogueVisibility.BuildKey("Commodore 64"));
        Assert.Equal("Commodore 64|250407", CatalogueVisibility.BuildKey("Commodore 64", "250407"));
        Assert.Equal(
            "Commodore 64|250407|Top (replica)",
            CatalogueVisibility.BuildKey("Commodore 64", "250407", "Top (replica)"));
    }

    // -------------------------------------------------------------- Everything checked by default

    [Fact]
    public void Nothing_unchecked_means_everything_is_visible()
    {
        var unchecked_ = new HashSet<string>();

        Assert.True(CatalogueVisibility.IsHardwareVisible(unchecked_, "Commodore 64"));
        Assert.True(CatalogueVisibility.IsBoardVisible(unchecked_, "Commodore 64", "250407"));
        Assert.True(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top"));
    }

    // -------------------------------------------------------------- Own key unchecked

    [Fact]
    public void Unchecking_a_hardware_hides_only_that_hardware()
    {
        var unchecked_ = new HashSet<string> { "Commodore 128" };

        Assert.False(CatalogueVisibility.IsHardwareVisible(unchecked_, "Commodore 128"));
        Assert.True(CatalogueVisibility.IsHardwareVisible(unchecked_, "Commodore 64"));
    }

    [Fact]
    public void Unchecking_a_board_hides_only_that_board_and_its_schematics()
    {
        var unchecked_ = new HashSet<string> { "Commodore 64|250407" };

        Assert.True(CatalogueVisibility.IsHardwareVisible(unchecked_, "Commodore 64"));
        Assert.False(CatalogueVisibility.IsBoardVisible(unchecked_, "Commodore 64", "250407"));
        Assert.True(CatalogueVisibility.IsBoardVisible(unchecked_, "Commodore 64", "KU-14194HB"));
        Assert.False(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top"));
    }

    [Fact]
    public void Unchecking_a_schematic_hides_only_that_schematic()
    {
        var unchecked_ = new HashSet<string> { "Commodore 64|250407|Bottom (replica)" };

        Assert.True(CatalogueVisibility.IsBoardVisible(unchecked_, "Commodore 64", "250407"));
        Assert.False(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Bottom (replica)"));
        Assert.True(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top (replica)"));
    }

    // -------------------------------------------------------------- Ancestor AND rule

    // The core rule this class exists for: a schematic or board can still be individually
    // "checked" in storage (its own key is absent from the unchecked set) and yet be hidden,
    // because an ancestor is unchecked. This is what lets a parent hide a whole subtree without
    // ever touching the children's own stored keys.
    [Fact]
    public void An_unchecked_hardware_hides_every_board_and_schematic_beneath_it_even_though_their_own_keys_are_still_checked()
    {
        var unchecked_ = new HashSet<string> { "Commodore 128" };

        Assert.False(CatalogueVisibility.IsBoardVisible(unchecked_, "Commodore 128", "310378"));
        Assert.False(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 128", "310378", "Top (rev. 7)"));

        // The child keys themselves were never added to the unchecked set.
        Assert.DoesNotContain("Commodore 128|310378", unchecked_);
        Assert.DoesNotContain("Commodore 128|310378|Top (rev. 7)", unchecked_);
    }

    [Fact]
    public void An_unchecked_board_hides_its_schematics_even_though_their_own_keys_are_still_checked()
    {
        var unchecked_ = new HashSet<string> { "Commodore 64|250407" };

        Assert.False(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top (replica)"));
        Assert.DoesNotContain("Commodore 64|250407|Top (replica)", unchecked_);
    }

    // Re-checking a parent (removing its key from the unchecked set) must restore exactly what
    // was visible before, with no change needed to the children's own keys.
    [Fact]
    public void Rechecking_a_parent_restores_child_visibility_without_touching_child_keys()
    {
        var unchecked_ = new HashSet<string> { "Commodore 64|250407" };
        Assert.False(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top (replica)"));

        unchecked_.Remove("Commodore 64|250407");
        Assert.True(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top (replica)"));
    }

    [Fact]
    public void A_schematic_hidden_by_its_own_key_stays_hidden_even_after_its_parents_are_rechecked()
    {
        var unchecked_ = new HashSet<string>
        {
            "Commodore 64",
            "Commodore 64|250407",
            "Commodore 64|250407|Bottom (replica)"
        };

        unchecked_.Remove("Commodore 64");
        unchecked_.Remove("Commodore 64|250407");

        Assert.True(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top (replica)"));
        Assert.False(CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Bottom (replica)"));
    }

    // -------------------------------------------------------------- Case-insensitive keys

    // Every other hardware/board/schematic name comparison in the app is OrdinalIgnoreCase, and
    // these keys are built from those same names. A key stored before a synced data file recased
    // a name must keep hiding the item afterwards - the failure mode is silent, since the item
    // simply reappears and the stale entry is not visible anywhere for the user to notice.
    //
    // Asserted against a set with the DEFAULT ordinal comparer on purpose: that is what
    // System.Text.Json hands back for a deserialized HashSet, so the predicates cannot rely on
    // the caller's set having the right comparer.
    [Fact]
    public void Keys_are_matched_case_insensitively_even_in_an_ordinal_set()
    {
        var unchecked_ = new HashSet<string>
        {
            "commodore 64",
            "Commodore 128|310378",
            "Commodore 64|250407|top (replica)"
        };

        Assert.False(CatalogueVisibility.IsHardwareVisible(unchecked_, "Commodore 64"));
        Assert.False(CatalogueVisibility.IsBoardVisible(unchecked_, "COMMODORE 128", "310378"));
        Assert.False(
            CatalogueVisibility.IsSchematicVisible(unchecked_, "Commodore 64", "250407", "Top (Replica)"));

        // A genuinely different name is still visible - case-insensitivity must not become
        // "matches anything".
        Assert.True(CatalogueVisibility.IsHardwareVisible(unchecked_, "Amstrad CPC"));
    }

    [Fact]
    public void Keys_are_matched_case_insensitively_in_a_case_insensitive_set_too()
    {
        var unchecked_ = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "commodore 64" };

        Assert.False(CatalogueVisibility.IsHardwareVisible(unchecked_, "COMMODORE 64"));
        Assert.True(CatalogueVisibility.IsHardwareVisible(unchecked_, "Commodore 128"));
    }

    // -------------------------------------------------------------- Key shape

    // Main reads a changed key's DEPTH to decide what a checkbox toggle has to refresh: only a
    // schematic key can require the current board to reload, and only a hardware/board key can
    // change what the drop-downs contain.
    [Fact]
    public void A_keys_depth_is_read_from_its_separators()
    {
        Assert.True(CatalogueVisibility.IsHardwareKey("Commodore 64"));
        Assert.False(CatalogueVisibility.IsBoardKey("Commodore 64"));
        Assert.False(CatalogueVisibility.IsSchematicKey("Commodore 64"));

        Assert.True(CatalogueVisibility.IsBoardKey("Commodore 64|250407"));
        Assert.False(CatalogueVisibility.IsSchematicKey("Commodore 64|250407"));

        Assert.True(CatalogueVisibility.IsSchematicKey("Commodore 64|250407|Top"));
        Assert.False(CatalogueVisibility.IsHardwareKey("Commodore 64|250407|Top"));
    }

    [Fact]
    public void KeyNamesBoard_accepts_the_board_and_its_schematics_and_nothing_else()
    {
        var board = ("Commodore 64", "250407");

        Assert.True(CatalogueVisibility.KeyNamesBoard("Commodore 64|250407", board));
        Assert.True(CatalogueVisibility.KeyNamesBoard("Commodore 64|250407|Top", board));

        // Case-insensitive here too, for the same reason the visibility predicates are.
        Assert.True(CatalogueVisibility.KeyNamesBoard("commodore 64|250407|top", board));

        // A hardware key names no single board, so it cannot name this one.
        Assert.False(CatalogueVisibility.KeyNamesBoard("Commodore 64", board));

        // A different board under the same hardware, and a board whose name merely STARTS with
        // this one's - the separator is what makes the prefix test safe.
        Assert.False(CatalogueVisibility.KeyNamesBoard("Commodore 64|KU-14194HB", board));
        Assert.False(CatalogueVisibility.KeyNamesBoard("Commodore 64|2504071", board));
    }

    [Fact]
    public void KeyNamesBoard_matches_nothing_when_no_board_is_selected()
    {
        Assert.False(CatalogueVisibility.KeyNamesBoard("Commodore 64|250407", (string.Empty, string.Empty)));
        Assert.False(CatalogueVisibility.KeyNamesBoard("Commodore 64|250407", ("Commodore 64", string.Empty)));
    }

    // -------------------------------------------------------------- CurrentSelectionSurvives
    //
    // Main.ApplyCatalogueVisibility's actual bug report: unchecking/rechecking a hardware or board
    // UNRELATED to the one on screen still forced a full board reload (and, with the detached
    // thumbnails window open, closed and reopened it) because the drop-down repopulation reselected
    // the current item even though it never stopped being visible. These pin the rule that decides
    // whether that reselect-and-reload is actually warranted.

    [Fact]
    public void The_current_selection_survives_when_both_names_are_still_in_the_visible_lists()
    {
        var hardwareNames = new[] { "Commodore 64", "Commodore 128" };
        var boardNames = new[] { "250407", "KU-14194HB" };

        Assert.True(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, "Commodore 64", "250407"));
    }

    // The exact reported shape: hiding a DIFFERENT hardware shrinks the visible hardware list, but
    // the one actually on screen is still in it, so the selection survives and no reload is needed.
    [Fact]
    public void Hiding_an_unrelated_hardware_does_not_break_survival_of_the_selected_one()
    {
        // "Commodore 128" has just been unchecked and so is no longer in the visible list -
        // "Commodore 64" (the one on screen) still is.
        var hardwareNames = new[] { "Commodore 64" };
        var boardNames = new[] { "250407" };

        Assert.True(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, "Commodore 64", "250407"));
    }

    // The mirror for a board: hiding a SIBLING board under the same hardware must not break the
    // currently selected board's own survival.
    [Fact]
    public void Hiding_a_sibling_board_does_not_break_survival_of_the_selected_one()
    {
        var hardwareNames = new[] { "Commodore 64" };

        // "KU-14194HB" has just been unchecked; "250407" (the one on screen) still shows.
        var boardNames = new[] { "250407" };

        Assert.True(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, "Commodore 64", "250407"));
    }

    [Fact]
    public void The_selection_does_not_survive_when_the_current_hardware_itself_was_hidden()
    {
        var hardwareNames = new[] { "Commodore 128" };
        var boardNames = System.Array.Empty<string>();

        Assert.False(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, "Commodore 64", "250407"));
    }

    [Fact]
    public void The_selection_does_not_survive_when_the_current_board_itself_was_hidden()
    {
        var hardwareNames = new[] { "Commodore 64" };
        var boardNames = new[] { "KU-14194HB" };

        Assert.False(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, "Commodore 64", "250407"));
    }

    // No hardware selected at all never "survives" - there is nothing on screen for a checkbox
    // change to leave alone, so the caller should fall back to its ordinary reselect path.
    [Fact]
    public void Nothing_survives_when_no_hardware_is_currently_selected()
    {
        var hardwareNames = new[] { "Commodore 64" };
        var boardNames = new[] { "250407" };

        Assert.False(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, string.Empty, string.Empty));
    }

    // Hardware selected but no board yet (e.g. that hardware has no boards, or none is selected
    // yet) trivially survives on the board half - there is no board selection to invalidate.
    [Fact]
    public void A_selected_hardware_with_no_board_selected_survives_on_the_hardware_check_alone()
    {
        var hardwareNames = new[] { "Commodore 64" };
        var boardNames = System.Array.Empty<string>();

        Assert.True(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, "Commodore 64", string.Empty));
    }

    // Case-insensitive, matching every other name comparison in this class and the rest of the app.
    [Fact]
    public void Survival_is_case_insensitive()
    {
        var hardwareNames = new[] { "commodore 64" };
        var boardNames = new[] { "250407" };

        Assert.True(CatalogueVisibility.CurrentSelectionSurvives(
            hardwareNames, boardNames, "Commodore 64", "250407"));
    }
}
