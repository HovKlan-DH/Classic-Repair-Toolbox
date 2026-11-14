using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for ComponentListBuilder - the region filter, category filter and
// search that decide which components the main window lists.
//
// This was private static logic inside Main. The rules it encodes are invisible from the UI and
// easy to break: a blank region means "shared", search terms are ANDed across the whole composed
// display string, and a component with no usable text is dropped entirely.
public class ComponentListBuilderTests
{
    // ###########################################################################################
    // What stays selected as the region is switched (owner report, 2026-11-13: a component selected was let go of). A key
    // holds its row's region, so a component told apart per region is found again by its label - its row of the region now
    // shown; one of every region keeps its key; one with no row in the region now shown has nothing to select.
    // ###########################################################################################
    [Fact]
    public void A_selected_component_is_found_again_by_its_label_in_the_other_regions_list()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "C116", TechnicalNameOrValue = "10uF", Region = "PAL" },
                new ComponentEntry { BoardLabel = "C116", TechnicalNameOrValue = "47uF", Region = "NTSC" },
                new ComponentEntry { BoardLabel = "C1", TechnicalNameOrValue = "100nF" },
                new ComponentEntry { BoardLabel = "U30", TechnicalNameOrValue = "PAL only", Region = "PAL" },
                new ComponentEntry { BoardLabel = "C2", TechnicalNameOrValue = "1nF" },
            ],
        };

        List<ComponentListItem> pal = ComponentListBuilder.BuildComponentItems(board, "PAL");
        List<ComponentListItem> ntsc = ComponentListBuilder.BuildComponentItems(board, "NTSC");

        string KeyOf(List<ComponentListItem> rows, string label) => rows.Single(row => row.BoardLabel == label).SelectionKey;

        HashSet<string> selected = [KeyOf(pal, "C116"), KeyOf(pal, "C1"), KeyOf(pal, "U30")];

        List<int> rows = ComponentListBuilder.RowsToSelect(ntsc, selected);

        Assert.Equal(
            new[] { "C1", "C116" },
            rows.Select(index => ntsc[index].BoardLabel).OrderBy(label => label, StringComparer.Ordinal));
        Assert.Contains(ntsc.FindIndex(row => row.SelectionKey == KeyOf(ntsc, "C116")), rows);
    }

    // ###########################################################################################
    // The same rule finds a component whose NAMES changed since it was held (code review, 2026-11-09) - its key holds them -
    // and it is the one rule every restore goes by (code review, 2026-11-13: the region's stood beside the region switch
    // alone, and the selection held across a load lost a component told apart per region).
    // ###########################################################################################
    [Fact]
    public void A_selected_component_whose_names_changed_is_found_again_by_its_label()
    {
        var before = new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" }, new ComponentEntry { BoardLabel = "U9" }] };
        var after = new BoardData { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "Processor" }, new ComponentEntry { BoardLabel = "U9" }] };

        string held = ComponentListBuilder.BuildComponentItems(before, string.Empty)[0].SelectionKey;
        List<ComponentListItem> rows = ComponentListBuilder.BuildComponentItems(after, string.Empty);

        Assert.Equal([0], ComponentListBuilder.RowsToSelect(rows, [held]));
        Assert.Empty(ComponentListBuilder.RowsToSelect(rows, []));
    }

    // A row says what the draft's checks found on it (the Contribute tab's list), changing as it stands: a
    // row made again lets go of the selection.
    [Fact]
    public void A_row_says_its_problems_and_tells_when_they_change()
    {
        var item = new ComponentListItem { BoardLabel = "U8" };
        var told = new List<string?>();
        item.PropertyChanged += (_, change) => told.Add(change.PropertyName);

        Assert.False(item.HasErrors);
        Assert.False(item.HasWarnings);

        item.Problems = new BoardProblemCounts(1, 3);

        Assert.True(item.HasErrors);
        Assert.Equal("1", item.ErrorsText);
        Assert.Equal("1 error on this component", item.ErrorsTip);
        Assert.True(item.HasWarnings);
        Assert.Equal("3", item.WarningsText);
        Assert.Equal("3 warnings on this component", item.WarningsTip);
        Assert.Contains(nameof(ComponentListItem.HasErrors), told);
        Assert.Contains(nameof(ComponentListItem.WarningsText), told);

        // The same again says nothing: this runs for every row, after every save.
        told.Clear();
        item.Problems = new BoardProblemCounts(1, 3);
        Assert.Empty(told);
    }

    private static BoardData Board(params ComponentEntry[] components)
        => new() { Components = components.ToList() };

    private static ComponentEntry Component(
        string boardLabel = "", string friendly = "", string technical = "",
        string category = "", string region = "")
        => new()
        {
            BoardLabel = boardLabel,
            FriendlyName = friendly,
            TechnicalNameOrValue = technical,
            Category = category,
            Region = region
        };

    // -------------------------------------------------------------- BuildDistinctCategories

    [Fact]
    public void Categories_are_returned_once_each_in_first_seen_order()
    {
        var board = Board(
            Component(boardLabel: "U1", category: "IC"),
            Component(boardLabel: "R1", category: "Resistor"),
            Component(boardLabel: "U2", category: "IC"));

        Assert.Equal(new[] { "IC", "Resistor" }, ComponentListBuilder.BuildDistinctCategories(board));
    }

    // De-duplication is case-insensitive, but the FIRST spelling seen is the one kept.
    [Fact]
    public void Categories_differing_only_in_case_collapse_to_the_first_spelling()
    {
        var board = Board(
            Component(boardLabel: "U1", category: "IC"),
            Component(boardLabel: "U2", category: "ic"));

        Assert.Equal(new[] { "IC" }, ComponentListBuilder.BuildDistinctCategories(board));
    }

    [Fact]
    public void Blank_categories_are_skipped()
    {
        var board = Board(
            Component(boardLabel: "U1", category: ""),
            Component(boardLabel: "U2", category: "   "),
            Component(boardLabel: "U3", category: "IC"));

        Assert.Equal(new[] { "IC" }, ComponentListBuilder.BuildDistinctCategories(board));
    }

    [Fact]
    public void A_board_with_no_components_yields_no_categories()
    {
        Assert.Empty(ComponentListBuilder.BuildDistinctCategories(Board()));
    }

    // -------------------------------------------------------------- region filtering

    // The central rule: a component with NO region is shared and appears in every region, while
    // one that names a region appears only there.
    [Fact]
    public void A_component_without_a_region_is_shared_across_every_region()
    {
        var board = Board(Component(boardLabel: "U1"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL"));
        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "NTSC"));
    }

    [Fact]
    public void A_component_naming_a_region_appears_only_in_that_region()
    {
        var board = Board(Component(boardLabel: "U1", region: "PAL"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL"));
        Assert.Empty(ComponentListBuilder.BuildComponentItems(board, "NTSC"));
    }

    [Fact]
    public void Region_matching_ignores_case_and_surrounding_whitespace()
    {
        var board = Board(Component(boardLabel: "U1", region: "  pal  "));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL"));
    }

    // A whitespace-only region trims to empty, which means shared - not "a region called space".
    [Fact]
    public void A_whitespace_only_region_counts_as_shared()
    {
        var board = Board(Component(boardLabel: "U1", region: "   "));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "NTSC"));
    }

    // -------------------------------------------------------------- display text

    [Fact]
    public void Display_text_joins_board_label_friendly_name_and_technical_name()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA", technical: "906114-01"));

        var items = ComponentListBuilder.BuildComponentItems(board, "PAL");

        Assert.Equal("U1 | PLA | 906114-01", items.Single().DisplayText);
    }

    // Blank parts are skipped rather than leaving an empty slot, so the separator never appears
    // with nothing around it.
    [Fact]
    public void Blank_parts_are_omitted_from_the_display_text()
    {
        var board = Board(Component(boardLabel: "U1", technical: "906114-01"));

        Assert.Equal("U1 | 906114-01", ComponentListBuilder.BuildComponentItems(board, "PAL").Single().DisplayText);
    }

    // A component with nothing to show at all is dropped entirely - it would otherwise render as
    // a blank, unclickable row.
    [Fact]
    public void A_component_with_no_usable_text_is_dropped()
    {
        var board = Board(Component(category: "IC", region: "PAL"));

        Assert.Empty(ComponentListBuilder.BuildComponentItems(board, "PAL"));
    }

    [Fact]
    public void Display_text_parts_are_trimmed()
    {
        var board = Board(Component(boardLabel: "  U1  ", friendly: "  PLA  "));

        Assert.Equal("U1 | PLA", ComponentListBuilder.BuildComponentItems(board, "PAL").Single().DisplayText);
    }

    // -------------------------------------------------------------- search

    // Terms are ANDed, and each is matched against the WHOLE composed display string - so a term
    // hitting the label and another hitting the technical name both count.
    [Fact]
    public void Every_search_term_must_match_somewhere_in_the_display_text()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA", technical: "906114-01"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL", null, "U1 906114"));
        Assert.Empty(ComponentListBuilder.BuildComponentItems(board, "PAL", null, "U1 missing"));
    }

    [Fact]
    public void Search_is_case_insensitive()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL", null, "pla"));
    }

    // A blank or whitespace-only search is "no filter", not "match nothing".
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_search_term_does_not_filter(string search)
    {
        var board = Board(Component(boardLabel: "U1"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL", null, search));
    }

    // Runs of spaces between terms collapse rather than producing an empty term that matches
    // everything (or nothing).
    [Fact]
    public void Repeated_spaces_between_search_terms_are_ignored()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL", null, "U1    PLA"));
    }

    // The search runs against the composed string, which includes the " | " separator - so a term
    // containing the separator can match across two fields. Documented, not endorsed.
    [Fact]
    public void A_search_term_can_span_the_field_separator()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL", null, "U1 | PLA"));
    }

    // ###########################################################################################
    // The search's one rule, for a list already built (code review, 2026-10-30: the Contribute tab
    // built its list twice per refresh, once only to count): narrowing the whole list by MatchesSearch
    // gives what the search itself builds, whatever the words.
    // ###########################################################################################
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("u")]
    [InlineData("PLA")]
    [InlineData("U1    pla")]
    [InlineData("U1 | PLA")]
    [InlineData("nothing")]
    public void Narrowing_the_whole_list_by_the_search_rule_gives_what_the_search_builds(string search)
    {
        var board = Board(
            Component(boardLabel: "U1", friendly: "PLA", technical: "906114"),
            Component(boardLabel: "U2", friendly: "CPU", region: "NTSC"),
            Component(boardLabel: "U3", friendly: "VIC", region: "PAL"),
            Component(boardLabel: "R1", technical: "10k"));

        List<ComponentListItem> all = ComponentListBuilder.BuildComponentItems(board, "PAL");
        List<ComponentListItem> searched = ComponentListBuilder.BuildComponentItems(board, "PAL", null, search);

        Assert.Equal(
            searched.Select(item => item.SelectionKey),
            all.Where(item => ComponentListBuilder.MatchesSearch(item.DisplayText, search)).Select(item => item.SelectionKey));
    }

    // -------------------------------------------------------------- category filter

    [Fact]
    public void Only_components_in_the_category_filter_are_listed()
    {
        var board = Board(
            Component(boardLabel: "U1", category: "IC"),
            Component(boardLabel: "R1", category: "Resistor"));

        var items = ComponentListBuilder.BuildComponentItems(
            board, "PAL", new HashSet<string> { "IC" });

        Assert.Equal("U1", items.Single().BoardLabel);
    }

    // The filter set is built with the default comparer, so category matching here is
    // case-SENSITIVE even though BuildDistinctCategories de-duplicates case-insensitively.
    [Fact]
    public void The_category_filter_matches_case_sensitively()
    {
        var board = Board(Component(boardLabel: "U1", category: "IC"));

        Assert.Empty(ComponentListBuilder.BuildComponentItems(board, "PAL", new HashSet<string> { "ic" }));
    }

    // A null filter means "no category filtering", distinct from an empty set which excludes all.
    [Fact]
    public void A_null_category_filter_lets_everything_through_but_an_empty_set_excludes_all()
    {
        var board = Board(Component(boardLabel: "U1", category: "IC"));

        Assert.Single(ComponentListBuilder.BuildComponentItems(board, "PAL", null));
        Assert.Empty(ComponentListBuilder.BuildComponentItems(board, "PAL", new HashSet<string>()));
    }

    // -------------------------------------------------------------- selection key

    // The key joins four fields with a unit separator (U+001F) so it stays unique even when a
    // field legitimately contains a pipe or a space.
    [Fact]
    public void The_selection_key_joins_four_trimmed_fields_with_a_unit_separator()
    {
        var board = Board(Component(
            boardLabel: " U1 ", friendly: " PLA ", technical: " 906114-01 ", region: " PAL "));

        var item = ComponentListBuilder.BuildComponentItems(board, "PAL").Single();

        const char separator = '\u001F';
        Assert.Equal(
            string.Join(separator, "U1", "PLA", "906114-01", "PAL"),
            item.SelectionKey);
    }

    // The key is read apart where it is made (code review, 2026-11-13): a key built here gives back its label and names.
    [Fact]
    public void A_key_built_here_reads_back_as_its_label_and_names()
    {
        var item = ComponentListBuilder.BuildComponentItems(
            Board(Component(boardLabel: "U1", friendly: "PLA", technical: "906114-01", region: "PAL")), "PAL").Single();

        Assert.Equal("U1", ComponentListBuilder.LabelOf(item.SelectionKey));
        Assert.Equal(("PLA", "906114-01"), ComponentListBuilder.NamesOf(item.SelectionKey));
        Assert.Equal((string.Empty, string.Empty), ComponentListBuilder.NamesOf(null));
    }

    // The key includes the region, so the PAL and NTSC variants of one label stay distinguishable.
    [Fact]
    public void Two_region_variants_of_one_label_get_different_selection_keys()
    {
        var pal = ComponentListBuilder
            .BuildComponentItems(Board(Component(boardLabel: "U1", region: "PAL")), "PAL").Single();
        var ntsc = ComponentListBuilder
            .BuildComponentItems(Board(Component(boardLabel: "U1", region: "NTSC")), "NTSC").Single();

        Assert.NotEqual(pal.SelectionKey, ntsc.SelectionKey);
    }

    [Fact]
    public void ToString_returns_the_display_text_so_the_list_renders_it()
    {
        var item = ComponentListBuilder
            .BuildComponentItems(Board(Component(boardLabel: "U1", friendly: "PLA")), "PAL").Single();

        Assert.Equal(item.DisplayText, item.ToString());
    }

    // -------------------------------------------------------------- IsDrafted (NewContributeStrategy.md Phase 2, session 2b)

    [Fact]
    public void A_component_in_the_drafted_labels_set_is_marked_IsDrafted()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA"));

        var item = ComponentListBuilder
            .BuildComponentItems(board, "", draftedBoardLabels: new HashSet<string> { "U1" }).Single();

        Assert.True(item.IsDrafted);
    }

    [Fact]
    public void A_component_not_in_the_drafted_labels_set_is_not_marked()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA"));

        var item = ComponentListBuilder
            .BuildComponentItems(board, "", draftedBoardLabels: new HashSet<string> { "U2" }).Single();

        Assert.False(item.IsDrafted);
    }

    [Fact]
    public void With_no_drafted_labels_set_given_nothing_is_marked()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA"));

        var item = ComponentListBuilder.BuildComponentItems(board, "").Single();

        Assert.False(item.IsDrafted);
    }

    [Fact]
    public void IsDrafted_matches_the_drafted_labels_set_case_insensitively()
    {
        var board = Board(Component(boardLabel: "U1", friendly: "PLA"));

        var item = ComponentListBuilder
            .BuildComponentItems(board, "", draftedBoardLabels: new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "u1" }).Single();

        Assert.True(item.IsDrafted);
    }

    // -------------------------------------------------------------- IsSupportedKiCadRawFile

    [Theory]
    [InlineData("board.kicad_pcb", true)]
    [InlineData("board.kicad_pro", true)]
    [InlineData("board.kicad_sch", true)]
    [InlineData("BOARD.KICAD_PCB", true)]
    [InlineData("  board.kicad_pcb  ", true)]
    [InlineData("board.kicad_prl", false)]     // a real KiCad file, but not one we load
    [InlineData("board.txt", false)]
    [InlineData("board", false)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    public void Only_the_three_loadable_KiCad_extensions_are_supported(string path, bool expected)
    {
        Assert.Equal(expected, ComponentListBuilder.IsSupportedKiCadRawFile(path));
    }

    // -------------------------------------------------------------- BuildComponentsInScope

    // This backs the worklog entry card's "Mark components in scope" checklist: given the set of
    // board labels a drawn area touches, return the matching components in board-data order (the
    // same order Overview and the main Component list already use).
    [Fact]
    public void Only_components_whose_label_is_in_scope_are_returned()
    {
        var board = Board(
            Component(boardLabel: "C1", friendly: "Ceramic"),
            Component(boardLabel: "R1", friendly: "Resistor"));

        var rows = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "C1" });

        Assert.Equal("C1", Assert.Single(rows).BoardLabel);
    }

    // Order must follow boardData.Components, not the (unordered) scope set, or the list would
    // shuffle components relative to how Overview and the main Component list show them.
    [Fact]
    public void Matched_components_keep_board_data_order_regardless_of_scope_set_order()
    {
        var board = Board(
            Component(boardLabel: "C3"),
            Component(boardLabel: "C1"),
            Component(boardLabel: "C2"));

        var rows = ComponentListBuilder.BuildComponentsInScope(
            board, new HashSet<string> { "C1", "C2", "C3" });

        Assert.Equal(new[] { "C3", "C1", "C2" }, rows.Select(row => row.BoardLabel));
    }

    // A regionalized component occupies several rows in the board Excel - one per region - all
    // sharing a board label. The checklist is a list of PHYSICAL components to mark in scope, and
    // there is only one physical U1 on the board, so it must appear once however many rows
    // describe it. Before this, marking an area around a regionalized part listed it two or three
    // times with no way to tell the entries apart.
    [Fact]
    public void A_component_with_one_row_per_region_is_listed_only_once()
    {
        var board = Board(
            Component(boardLabel: "U1", friendly: "VIC-II", technical: "6569", region: "PAL"),
            Component(boardLabel: "U1", friendly: "VIC-II", technical: "6567", region: "NTSC"));

        var rows = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "U1" });

        Assert.Equal("U1", Assert.Single(rows).BoardLabel);
    }

    // The first row wins, matching how BuildDistinctCategories already resolves a collision - and
    // the scope set itself is keyed by board label alone, so no row is more "correct" than another.
    [Fact]
    public void The_first_row_supplies_the_display_name_for_a_regionalized_component()
    {
        var board = Board(
            Component(boardLabel: "U1", friendly: "VIC-II", technical: "6569", region: "PAL"),
            Component(boardLabel: "U1", friendly: "VIC-II", technical: "6567", region: "NTSC"));

        var rows = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "U1" });

        Assert.Equal("VIC-II | 6569", Assert.Single(rows).DisplayName);
    }

    // De-duplication uses the same case-insensitive comparison the rest of the board-label
    // handling uses, so "u1" and "U1" are one component rather than two.
    [Fact]
    public void Rows_whose_labels_differ_only_in_case_collapse_to_one_component()
    {
        var board = Board(
            Component(boardLabel: "U1", friendly: "VIC-II"),
            Component(boardLabel: "u1", friendly: "VIC-II"));

        var rows = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "U1" });

        Assert.Equal("U1", Assert.Single(rows).BoardLabel);
    }

    // Distinct components must still all appear - the de-duplication is per label, not a blanket
    // "one row per call".
    [Fact]
    public void De_duplication_does_not_drop_genuinely_different_components()
    {
        var board = Board(
            Component(boardLabel: "U1", friendly: "VIC-II", region: "PAL"),
            Component(boardLabel: "U1", friendly: "VIC-II", region: "NTSC"),
            Component(boardLabel: "C1", friendly: "Ceramic"));

        var rows = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "U1", "C1" });

        Assert.Equal(new[] { "U1", "C1" }, rows.Select(row => row.BoardLabel));
    }

    // A blank board label identifies no component. Before the de-duplication was added these were
    // listed (as blank rows); with it they would have collapsed into ONE blank row, silently
    // hiding the rest. Neither is useful, so they are skipped entirely.
    [Fact]
    public void Components_with_a_blank_board_label_are_skipped_entirely()
    {
        var board = Board(
            Component(boardLabel: "", friendly: "Mystery one"),
            Component(boardLabel: "  ", friendly: "Mystery two"),
            Component(boardLabel: "C1", friendly: "Ceramic"));

        var rows = ComponentListBuilder.BuildComponentsInScope(
            board, new HashSet<string> { "", "C1" });

        Assert.Equal("C1", Assert.Single(rows).BoardLabel);
    }

    [Fact]
    public void Board_label_matching_is_case_insensitive()
    {
        var board = Board(Component(boardLabel: "C1"));

        var rows = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "c1" });

        Assert.Single(rows);
    }

    [Fact]
    public void An_empty_scope_yields_no_rows()
    {
        var board = Board(Component(boardLabel: "C1"));

        Assert.Empty(ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string>()));
    }

    // The display name deliberately excludes the board label - the UI shows that separately, bold.
    [Fact]
    public void Display_name_joins_friendly_and_technical_name_but_not_the_board_label()
    {
        var board = Board(Component(boardLabel: "C1", friendly: "Ceramic", technical: "100pF 25V"));

        var row = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "C1" }).Single();

        Assert.Equal("Ceramic | 100pF 25V", row.DisplayName);
    }

    [Fact]
    public void Blank_display_name_parts_are_omitted()
    {
        var board = Board(Component(boardLabel: "C1", technical: "100pF 25V"));

        var row = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "C1" }).Single();

        Assert.Equal("100pF 25V", row.DisplayName);
    }

    [Fact]
    public void Board_label_and_display_name_parts_are_trimmed()
    {
        var board = Board(Component(boardLabel: "  C1  ", friendly: "  Ceramic  "));

        var row = ComponentListBuilder.BuildComponentsInScope(board, new HashSet<string> { "C1" }).Single();

        Assert.Equal("C1", row.BoardLabel);
        Assert.Equal("Ceramic", row.DisplayName);
    }

    // -------------------------------------------------------- NarrowSelectionToScope

    // The rule when a worklog area is resized: it can only ever REMOVE. Shrinking an area off a
    // component the user had marked leaves that component recorded against a fault whose area no
    // longer covers it, so it goes.
    [Fact]
    public void A_component_the_area_no_longer_touches_is_dropped()
    {
        var kept = ComponentListBuilder.NarrowSelectionToScope(
            new[] { "U7", "C34", "Y1" },
            new HashSet<string> { "U7", "Y1" });

        Assert.Equal(new[] { "U7", "Y1" }, kept);
    }

    // The other half, and the one that is easy to get wrong: a component the area now covers is
    // NOT auto-selected. Being inside the rectangle is not the same as the user deciding it is
    // relevant - auto-ticking would quietly add components nobody chose, more of them the wider the
    // area is dragged.
    [Fact]
    public void A_newly_covered_component_is_not_selected_automatically()
    {
        var kept = ComponentListBuilder.NarrowSelectionToScope(
            new[] { "U7" },
            new HashSet<string> { "U7", "C34", "R3" });

        Assert.Equal(new[] { "U7" }, kept);
    }

    // Both directions at once, which is what an ordinary drag produces: the area slides off some
    // components and onto others.
    [Fact]
    public void A_resize_that_both_gains_and_loses_components_only_removes()
    {
        var kept = ComponentListBuilder.NarrowSelectionToScope(
            new[] { "U7", "C34" },
            new HashSet<string> { "C34", "R3", "Q2" });

        Assert.Equal(new[] { "C34" }, kept);
    }

    // Matching is case-insensitive, as board-label comparisons are everywhere else in the app - a
    // selection stored as "u7" must survive an area whose rects key it as "U7".
    [Fact]
    public void Selection_matching_ignores_case()
    {
        var kept = ComponentListBuilder.NarrowSelectionToScope(
            new[] { "u7" },
            new HashSet<string> { "U7" });

        Assert.Equal(new[] { "u7" }, kept);
    }

    // The user's own ordering is preserved rather than rebuilt from the board, so a hand-curated
    // entry keeps its arrangement across a resize.
    [Fact]
    public void The_existing_selection_order_is_preserved()
    {
        var kept = ComponentListBuilder.NarrowSelectionToScope(
            new[] { "Y1", "C34", "U7" },
            new HashSet<string> { "U7", "C34", "Y1" });

        Assert.Equal(new[] { "Y1", "C34", "U7" }, kept);
    }

    // An area dragged clear of everything empties the selection - that IS the correct answer, and
    // is distinct from the "scope unknown" case, which the caller handles by not calling this.
    [Fact]
    public void An_area_touching_nothing_clears_the_selection()
    {
        var kept = ComponentListBuilder.NarrowSelectionToScope(
            new[] { "U7", "C34" },
            new HashSet<string>());

        Assert.Empty(kept);
    }

    [Fact]
    public void An_entry_with_no_selection_stays_empty()
    {
        Assert.Empty(ComponentListBuilder.NarrowSelectionToScope(
            Array.Empty<string>(), new HashSet<string> { "U7" }));

        Assert.Empty(ComponentListBuilder.NarrowSelectionToScope(
            null, new HashSet<string> { "U7" }));
    }

    // Blank entries identify no component, so they cannot be in scope and are not carried forward
    // as unmatchable labels.
    [Fact]
    public void Blank_labels_are_dropped_from_the_selection()
    {
        var kept = ComponentListBuilder.NarrowSelectionToScope(
            new[] { "U7", "", "   " },
            new HashSet<string> { "U7", "" });

        Assert.Equal(new[] { "U7" }, kept);
    }

    // -------------------------------------------------------- NarrowEntryToScope

    // Resizing an area narrows the in-scope list AND the completed list together. The completed
    // list is measured against the scope that REMAINS, not against the raw area - which is the
    // whole reason this exists as one function rather than two calls at the call site.
    //
    // C2 is the case that separates the two: the area still touches it, but the user had unticked
    // it from the scope. Narrowing completed against the area would keep it, claiming a component
    // was done for an entry that does not include it.
    [Fact]
    public void Completed_is_narrowed_to_the_remaining_scope_not_to_the_raw_area()
    {
        var (inScope, completed) = ComponentListBuilder.NarrowEntryToScope(
            selectedBoardLabels: new[] { "C1", "C3" },
            completedBoardLabels: new[] { "C1", "C2" },
            boardLabelsInScope: new[] { "C1", "C2", "C3" });

        Assert.Equal(new[] { "C1", "C3" }, inScope);

        // C2 is gone: it is in the area but not in the entry's scope.
        Assert.Equal(new[] { "C1" }, completed);
    }

    // Shrinking an area off a component drops it from both lists at once - it is no longer part of
    // the entry, so it cannot remain recorded as completed work for it.
    [Fact]
    public void A_component_the_area_no_longer_covers_leaves_both_lists()
    {
        var (inScope, completed) = ComponentListBuilder.NarrowEntryToScope(
            selectedBoardLabels: new[] { "C1", "C2" },
            completedBoardLabels: new[] { "C1", "C2" },
            boardLabelsInScope: new[] { "C1" });

        Assert.Equal(new[] { "C1" }, inScope);
        Assert.Equal(new[] { "C1" }, completed);
    }

    // The invariant, asserted directly: completed is always a subset of in-scope, whatever went in.
    [Fact]
    public void Completed_is_always_a_subset_of_what_is_in_scope()
    {
        var (inScope, completed) = ComponentListBuilder.NarrowEntryToScope(
            selectedBoardLabels: new[] { "C1" },
            completedBoardLabels: new[] { "C1", "C2", "C9" },
            boardLabelsInScope: new[] { "C1", "C2" });

        Assert.All(completed, label => Assert.Contains(label, inScope));
    }

    // Same case-insensitive matching as everything else that compares board labels - a label
    // stored as "c1" must still be recognised against a scope that keys it as "C1".
    [Fact]
    public void Narrowing_an_entry_matches_labels_case_insensitively()
    {
        var (inScope, completed) = ComponentListBuilder.NarrowEntryToScope(
            selectedBoardLabels: new[] { "c1" },
            completedBoardLabels: new[] { "C1" },
            boardLabelsInScope: new[] { "C1" });

        Assert.Equal(new[] { "c1" }, inScope);
        Assert.Equal(new[] { "C1" }, completed);
    }

    [Fact]
    public void An_entry_with_nothing_completed_narrows_to_nothing_completed()
    {
        var (inScope, completed) = ComponentListBuilder.NarrowEntryToScope(
            selectedBoardLabels: new[] { "C1" },
            completedBoardLabels: null,
            boardLabelsInScope: new[] { "C1" });

        Assert.Equal(new[] { "C1" }, inScope);
        Assert.Empty(completed);
    }
}
