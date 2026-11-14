using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for OverviewHtmlBuilder - the bill-of-materials grouping and the
// printable HTML the Overview tab writes to a temp file and hands to the browser.
//
// Extracted from TabOverview. Two things here are worth guarding and cannot be seen by looking
// at the rendered page: that every contributed field goes through HtmlEncode, and that the BOM
// collapses rows on category + technical name + friendly name.
public class OverviewHtmlBuilderTests
{
    private static OverviewRow Row(
        string component = "", string category = "",
        string technical = "", string friendly = "")
        => new()
        {
            Component = component,
            Category = category,
            TechnicalName = technical,
            FriendlyName = friendly
        };

    // -------------------------------------------------------------- BuildOverviewDisplayString

    [Fact]
    public void The_display_string_joins_component_friendly_name_and_technical_name_in_that_order()
    {
        var row = Row(component: "U1", friendly: "PLA", technical: "906114-01");

        Assert.Equal("U1 | PLA | 906114-01", OverviewHtmlBuilder.BuildOverviewDisplayString(row));
    }

    [Fact]
    public void Blank_parts_are_omitted_from_the_display_string()
    {
        Assert.Equal("U1 | 906114-01",
            OverviewHtmlBuilder.BuildOverviewDisplayString(Row(component: "U1", technical: "906114-01")));
    }

    [Fact]
    public void A_row_with_nothing_to_show_yields_an_empty_display_string()
    {
        Assert.Equal(string.Empty, OverviewHtmlBuilder.BuildOverviewDisplayString(Row()));
    }

    [Fact]
    public void Display_string_parts_are_trimmed()
    {
        Assert.Equal("U1 | PLA",
            OverviewHtmlBuilder.BuildOverviewDisplayString(Row(component: "  U1  ", friendly: "  PLA  ")));
    }

    // -------------------------------------------------------------- BuildQuantityGroups

    // The grouping key is category + technical + friendly. The component designator is NOT part
    // of it - that is the whole point, since U1 and U2 of the same part must collapse into one
    // BOM line of quantity 2.
    [Fact]
    public void Identical_parts_collapse_into_one_line_with_a_quantity()
    {
        var groups = OverviewHtmlBuilder.BuildQuantityGroups(new[]
        {
            Row(component: "U1", category: "IC", technical: "7400", friendly: "NAND"),
            Row(component: "U2", category: "IC", technical: "7400", friendly: "NAND")
        });

        var group = Assert.Single(groups);
        Assert.Equal(2, group.Quantity);
        Assert.Equal("U1, U2", group.Components);
    }

    [Fact]
    public void Parts_differing_in_technical_name_stay_separate()
    {
        var groups = OverviewHtmlBuilder.BuildQuantityGroups(new[]
        {
            Row(component: "U1", category: "IC", technical: "7400"),
            Row(component: "U2", category: "IC", technical: "7402")
        });

        Assert.Equal(2, groups.Count);
    }

    // The grouping key uses the DEFAULT comparer, so it is case-sensitive - "7400" and "7400 "
    // or a differently-cased name do not collapse. The ordering afterwards is case-INSENSITIVE.
    [Fact]
    public void Grouping_is_case_sensitive_even_though_ordering_is_not()
    {
        var groups = OverviewHtmlBuilder.BuildQuantityGroups(new[]
        {
            Row(component: "U1", category: "IC", technical: "7400"),
            Row(component: "U2", category: "ic", technical: "7400")
        });

        Assert.Equal(2, groups.Count);
    }

    [Fact]
    public void Groups_are_ordered_by_type_then_technical_name()
    {
        var groups = OverviewHtmlBuilder.BuildQuantityGroups(new[]
        {
            Row(component: "R1", category: "Resistor", technical: "1k"),
            Row(component: "U2", category: "IC", technical: "7402"),
            Row(component: "U1", category: "IC", technical: "7400")
        });

        Assert.Equal(new[] { "7400", "7402", "1k" }, groups.Select(g => g.TechnicalName));
    }

    [Fact]
    public void Components_within_a_group_are_listed_in_source_order()
    {
        var groups = OverviewHtmlBuilder.BuildQuantityGroups(new[]
        {
            Row(component: "U3", category: "IC", technical: "7400"),
            Row(component: "U1", category: "IC", technical: "7400")
        });

        Assert.Equal("U3, U1", groups.Single().Components);
    }

    [Fact]
    public void No_rows_yields_no_groups()
    {
        Assert.Empty(OverviewHtmlBuilder.BuildQuantityGroups(Array.Empty<OverviewRow>()));
    }

    // -------------------------------------------------------------- BuildPrintableHtml

    [Fact]
    public void The_printable_component_list_is_a_complete_html_document()
    {
        string html = OverviewHtmlBuilder.BuildPrintableHtml(new[] { Row(component: "U1") });

        Assert.StartsWith("<!DOCTYPE html>", html, StringComparison.Ordinal);
        Assert.Contains("<title>Component List</title>", html, StringComparison.Ordinal);
        Assert.Contains("</html>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void The_printable_component_list_emits_one_row_per_component()
    {
        string html = OverviewHtmlBuilder.BuildPrintableHtml(new[]
        {
            Row(component: "U1", technical: "7400", friendly: "NAND"),
            Row(component: "U2", technical: "7402", friendly: "NOR")
        });

        Assert.Contains("<td>U1</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>7402</td>", html, StringComparison.Ordinal);
    }

    // The document self-prints on load - that is how the tab "prints" without a print dialog of
    // its own, so losing this line would silently break the feature.
    [Fact]
    public void The_printable_document_triggers_printing_on_load()
    {
        string html = OverviewHtmlBuilder.BuildPrintableHtml(Array.Empty<OverviewRow>());

        Assert.Contains("window.print()", html, StringComparison.Ordinal);
    }

    // Component data is contributed, so it must never be able to inject markup into the page.
    [Fact]
    public void Contributed_text_is_html_encoded_in_the_component_list()
    {
        string html = OverviewHtmlBuilder.BuildPrintableHtml(new[]
        {
            Row(component: "<script>alert(1)</script>", technical: "a & b")
        });

        Assert.DoesNotContain("<script>alert(1)</script>", html, StringComparison.Ordinal);
        Assert.Contains("&lt;script&gt;", html, StringComparison.Ordinal);
        Assert.Contains("a &amp; b", html, StringComparison.Ordinal);
    }

    [Fact]
    public void A_printable_component_list_with_no_rows_is_still_a_valid_document()
    {
        string html = OverviewHtmlBuilder.BuildPrintableHtml(Array.Empty<OverviewRow>());

        Assert.Contains("<tbody>", html, StringComparison.Ordinal);
        Assert.Contains("</html>", html, StringComparison.Ordinal);
    }

    // -------------------------------------------------------------- BuildPrintableQuantitiesHtml

    [Fact]
    public void The_printable_bom_is_titled_as_a_bill_of_materials()
    {
        string html = OverviewHtmlBuilder.BuildPrintableQuantitiesHtml(Array.Empty<OverviewQuantityGroup>());

        Assert.Contains("<title>Bill of Materials (BOM)</title>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void The_printable_bom_emits_each_group_with_its_quantity()
    {
        string html = OverviewHtmlBuilder.BuildPrintableQuantitiesHtml(new[]
        {
            new OverviewQuantityGroup
            {
                Type = "IC", Components = "U1, U2", TechnicalName = "7400",
                FriendlyName = "NAND", Quantity = 2
            }
        });

        Assert.Contains("<td>U1, U2</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>2</td>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Contributed_text_is_html_encoded_in_the_bom()
    {
        string html = OverviewHtmlBuilder.BuildPrintableQuantitiesHtml(new[]
        {
            new OverviewQuantityGroup { Type = "<b>IC</b>", Quantity = 1 }
        });

        Assert.DoesNotContain("<b>IC</b>", html, StringComparison.Ordinal);
        Assert.Contains("&lt;b&gt;", html, StringComparison.Ordinal);
    }

    // -------------------------------------------------------------- OverviewRow

    // The print checkbox defaults to ticked, so "print" means "everything" until the user
    // deselects rows.
    [Fact]
    public void A_new_overview_row_is_selected_for_print_by_default()
    {
        Assert.True(Row().IsSelectedForPrint);
    }

    [Fact]
    public void Changing_the_print_selection_raises_a_property_changed_notification()
    {
        var row = Row();
        var raised = new List<string?>();
        row.PropertyChanged += (_, e) => raised.Add(e.PropertyName);

        row.IsSelectedForPrint = false;

        Assert.Equal(new[] { nameof(OverviewRow.IsSelectedForPrint) }, raised);
    }

    // Setting the same value short-circuits, so the grid does not churn on a no-op assignment.
    [Fact]
    public void Re_selecting_the_same_print_value_raises_nothing()
    {
        var row = Row();
        var raised = new List<string?>();
        row.PropertyChanged += (_, e) => raised.Add(e.PropertyName);

        row.IsSelectedForPrint = true;

        Assert.Empty(raised);
    }

    // -------------------------------------------------------------- OverviewLink

    [Theory]
    [InlineData(OverviewLinkType.LocalFile, true, false)]
    [InlineData(OverviewLinkType.WebLink, false, true)]
    public void A_link_reports_its_kind_through_the_two_binding_flags(
        OverviewLinkType type, bool expectLocal, bool expectWeb)
    {
        var link = new OverviewLink("name", "target", type);

        Assert.Equal(expectLocal, link.IsLocalFile);
        Assert.Equal(expectWeb, link.IsWebLink);
    }

    // -------------------------------------------------------------- BuildRows (CRT 4.0, slice 3c)

    // One row per component row, in the board's order, with its package, its files and links.
    [Fact]
    public void A_boards_rows_carry_each_components_fields_package_files_and_links()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Category = "IC", TechnicalNameOrValue = "6510", FriendlyName = "CPU", PartNumber = "906", Package = " DIP-40 ", Description = "The processor" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            ],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Component local files/6510.pdf" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "u8", Name = "Article", Url = "https://example.org" }],
        };

        List<OverviewRow> rows = OverviewHtmlBuilder.BuildRows(board);

        Assert.Equal(["U8", "C1"], rows.Select(row => row.Component));
        Assert.Equal("DIP-40", rows[0].Package);
        Assert.Equal("The processor", rows[0].ShortDescription);
        Assert.Equal(["Datasheet", "Article"], rows[0].Links.Select(link => link.Name));
        Assert.True(rows[0].Links[0].IsLocalFile);
        Assert.Empty(rows[1].Links);
    }

    // Notes are the component's first image's note.
    [Fact]
    public void Notes_are_the_first_images_note()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8" },
                new ComponentEntry { BoardLabel = "U9" },
                new ComponentEntry { BoardLabel = "U10" },
            ],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Note = "Image note of U8" },
                new ComponentImageEntry { BoardLabel = "U9", Note = "First of U9" },
                new ComponentImageEntry { BoardLabel = "U9", Note = "Second of U9" },
            ],
        };

        List<OverviewRow> rows = OverviewHtmlBuilder.BuildRows(board);

        Assert.Equal(["Image note of U8", "First of U9", string.Empty], rows.Select(row => row.Notes));
    }

    // ###########################################################################################
    // ONE ROW PER COMPONENT, AS THE REGION SHOWN HAS IT (owner report, 2026-11-13: "component "C116" has 4 regions applied,
    // so I see the component on 4 rows, which is wrong, as it should show only once"). One row per component ROW until then.
    // The row is the component's in the region shown, as its info window picks it - and its files and links are what that
    // window lists there: those of every region and of that region (a link of one region; owner request, 2026-11-02).
    // ###########################################################################################
    [Fact]
    public void A_component_of_several_regions_is_one_row_as_the_region_shown_has_it()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "C116", TechnicalNameOrValue = "10uF", Region = "Europe" },
                new ComponentEntry { BoardLabel = "U1", FriendlyName = "VIC", TechnicalNameOrValue = "6569", Region = "PAL" },
                new ComponentEntry { BoardLabel = "C116", TechnicalNameOrValue = "22uF", Region = "Japan" },
                new ComponentEntry { BoardLabel = "U1", FriendlyName = "VIC", TechnicalNameOrValue = "6567", Region = "NTSC" },
                new ComponentEntry { BoardLabel = "C116", TechnicalNameOrValue = "10uF", Region = "PAL" },
                new ComponentEntry { BoardLabel = "C116", TechnicalNameOrValue = "47uF", Region = "NTSC" },
                new ComponentEntry { BoardLabel = "U2" },
            ],
            ComponentLocalFiles =
            [
                new ComponentLocalFileEntry { BoardLabel = "U1", Name = "All", File = "a.pdf" },
                new ComponentLocalFileEntry { BoardLabel = "U1", Name = "NTSC", File = "b.pdf", Region = "NTSC" },
            ],
            ComponentLinks =
            [
                new ComponentLinkEntry { BoardLabel = "U1", Name = "PAL", Url = "https://example.org/pal", Region = "PAL" },
                new ComponentLinkEntry { BoardLabel = "U2", Name = "One region", Url = "https://example.org/u2", Region = "NTSC" },
            ],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U1", Note = "The PAL one", Region = "PAL" },
                new ComponentImageEntry { BoardLabel = "U1", Note = "The NTSC one", Region = "NTSC" },
            ],
        };

        List<OverviewRow> ntsc = OverviewHtmlBuilder.BuildRows(board, "NTSC");

        // In the board's order, each once.
        Assert.Equal(["C116", "U1", "U2"], ntsc.Select(row => row.Component));
        Assert.Equal(["47uF", "6567", string.Empty], ntsc.Select(row => row.TechnicalName));
        Assert.Equal(["All", "NTSC"], ntsc[1].Links.Select(link => link.Name));
        Assert.Equal(["One region"], ntsc[2].Links.Select(link => link.Name));
        Assert.Equal("The NTSC one", ntsc[1].Notes);

        List<OverviewRow> pal = OverviewHtmlBuilder.BuildRows(board, "PAL");

        Assert.Equal(["10uF", "6569", string.Empty], pal.Select(row => row.TechnicalName));
        Assert.Equal(["All", "PAL"], pal[1].Links.Select(link => link.Name));
        Assert.Empty(pal[2].Links);
        Assert.Equal("The PAL one", pal[1].Notes);

        // And so counted once in the bill of materials.
        Assert.Equal(1, OverviewHtmlBuilder.BuildQuantityGroups(pal).Single(group => group.Components == "C116").Quantity);
    }

    // ###########################################################################################
    // A row with no label is no component, and the Overview has no row for it (code review, 2026-11-14, checked): the tab never
    // listed one either way - it shows the labels of CRT's component list, which leaves a blank one out - so it is never in a
    // printed list or the bill of materials.
    // ###########################################################################################
    [Fact]
    public void A_row_with_no_label_gives_no_overview_row()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Category = "IC" },
                new ComponentEntry { BoardLabel = "  ", Category = "IC", FriendlyName = "Label cleared in the table" },
                new ComponentEntry { BoardLabel = string.Empty, Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            ],
        };

        List<OverviewRow> rows = OverviewHtmlBuilder.BuildRows(board);

        Assert.Equal(["U8", "C1"], rows.Select(row => row.Component));
        Assert.Equal(2, OverviewHtmlBuilder.BuildQuantityGroups(rows).Sum(group => group.Quantity));
    }

    // With no region picked every row of the component counts - its files and links of any region - and it is still one row.
    [Fact]
    public void With_no_region_a_component_is_one_row_with_all_its_links()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Region = "PAL" },
                new ComponentEntry { BoardLabel = "u1", Region = "NTSC" },
            ],
            ComponentLinks =
            [
                new ComponentLinkEntry { BoardLabel = "U1", Name = "PAL", Url = "https://example.org/pal", Region = "PAL" },
                new ComponentLinkEntry { BoardLabel = "U1", Name = "NTSC", Url = "https://example.org/ntsc", Region = "NTSC" },
            ],
        };

        OverviewRow row = Assert.Single(OverviewHtmlBuilder.BuildRows(board));

        Assert.Equal("U1", row.Component);
        Assert.Equal(["PAL", "NTSC"], row.Links.Select(link => link.Name));
    }
}
