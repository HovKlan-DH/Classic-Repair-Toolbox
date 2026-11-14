using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Aggregation and printable-HTML generation for the Overview tab.
    //
    // Extracted from TabOverview: building a string touches no control, and the two things worth
    // guarding here are exactly the things you cannot see by looking at the rendered page - that
    // every user-supplied field goes through HtmlEncode, and that the bill-of-materials grouping
    // collapses on category + technical name + friendly name (not on the component designator).
    // ###########################################################################################
    public static class OverviewHtmlBuilder
    {
        // ###########################################################################################
        // The Overview tab's rows of a board - ONE PER COMPONENT, in the board's order, AS THE REGION
        // SHOWN HAS IT (owner report, 2026-11-13: "for a component I now see it duplicated for as many
        // times as it has a region. E.g. component "C116" has 4 regions applied, so I see the component on
        // 4 rows, which is wrong, as it should show only once"). One row per component ROW until then, so a
        // component told apart per region was listed once per region - and counted so in the bill of
        // materials. Now each component is what its component info window shows in `region`: its row there
        // (ComponentImageQueries.PickComponentEntry), its files and links there (ComponentPopupContent -
        // a row of every region, and one of that region), and NOTES its first image's note there.
        //
        // `region` null: no region picked - every row's files, links and image, as a board with no regions has.
        // ###########################################################################################
        public static List<OverviewRow> BuildRows(BoardData boardData, string? region = null)
        {
            ArgumentNullException.ThrowIfNull(boardData);

            var rows = new List<OverviewRow>();

            // A row with no label is no component, and none was ever listed or printed: the tab shows only the labels CRT's
            // component list shows, which leaves a blank one out (TabOverview.GetFilteredRows - a code review of 2026-11-14
            // took the grouping's Where for a change, checked; the table marks such a row as an error).
            IEnumerable<IGrouping<string, ComponentEntry>> components = boardData.Components
                .Where(row => !string.IsNullOrWhiteSpace(row.BoardLabel))
                .GroupBy(row => row.BoardLabel!.Trim(), StringComparer.OrdinalIgnoreCase);

            // Each component's images, files and links, found once by label - in any capitals, as ComponentPaneEdits'
            // IsOfComponent finds them, in the board's order - not by a walk of every row per component (code review,
            // 2026-11-13: every region switch builds the rows again, on the UI thread).
            ILookup<string, ComponentImageEntry> imagesOf = boardData.ComponentImages.ToLookup(row => row.BoardLabel?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
            ILookup<string, ComponentLocalFileEntry> filesOf = boardData.ComponentLocalFiles.ToLookup(row => row.BoardLabel?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
            ILookup<string, ComponentLinkEntry> linksOf = boardData.ComponentLinks.ToLookup(row => row.BoardLabel?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase);

            bool InRegion(string? rowRegion) => region is null || ComponentPaneEdits.InRegion(rowRegion, region);

            foreach (IGrouping<string, ComponentEntry> ofLabel in components)
            {
                string label = ofLabel.Key;
                ComponentEntry comp = ComponentImageQueries.PickComponentEntry(ofLabel.ToList(), region?.Trim() ?? string.Empty)!;

                // A note's plain text - its formatting marks are no part of a table cell (RichNote, 2026-11-07).
                string note = RichNote.PlainOf(imagesOf[label]
                    .FirstOrDefault(ci => InRegion(ci.Region))?.Note);

                var links = new List<OverviewLink>();

                // In a region, what the component info window lists - its own rule, on this label's rows (code review,
                // 2026-11-13: a copy of it stood here); with none given, every row of the label.
                links.AddRange((region is null ? filesOf[label] : ComponentPopupContent.FilesOf(filesOf[label], label, region))
                    .Select(lf => new OverviewLink(lf.Name, lf.File, OverviewLinkType.LocalFile)));

                links.AddRange((region is null ? linksOf[label] : ComponentPopupContent.LinksOf(linksOf[label], label, region))
                    .Select(l => new OverviewLink(l.Name, l.Url, OverviewLinkType.WebLink)));

                rows.Add(new OverviewRow
                {
                    IsSelectedForPrint = true,
                    Component = comp.BoardLabel?.Trim() ?? label,
                    Category = comp.Category ?? string.Empty,
                    TechnicalName = comp.TechnicalNameOrValue ?? string.Empty,
                    FriendlyName = comp.FriendlyName ?? string.Empty,
                    PartNumber = comp.PartNumber ?? string.Empty,
                    Package = comp.Package?.Trim() ?? string.Empty,
                    ShortDescription = comp.Description ?? string.Empty,
                    Notes = note,
                    Links = links
                });
            }

            return rows;
        }

        // ###########################################################################################
        // Composes the "component | friendly | technical" display string for one overview row,
        // skipping blank parts so the separator never appears with nothing around it.
        // ###########################################################################################
        public static string BuildOverviewDisplayString(OverviewRow row)
        {
            var parts = new List<string>(3);

            if (!string.IsNullOrWhiteSpace(row.Component))
                parts.Add(row.Component.Trim());
            if (!string.IsNullOrWhiteSpace(row.FriendlyName))
                parts.Add(row.FriendlyName.Trim());
            if (!string.IsNullOrWhiteSpace(row.TechnicalName))
                parts.Add(row.TechnicalName.Trim());

            return string.Join(" | ", parts);
        }

        // ###########################################################################################
        // Collapses identical components into bill-of-materials lines, ordered by type then
        // technical name. Components sharing a group are listed comma-separated in source order.
        // ###########################################################################################
        public static List<OverviewQuantityGroup> BuildQuantityGroups(IEnumerable<OverviewRow> rows)
        {
            return rows
                .GroupBy(row => new
                {
                    Category = row.Category ?? string.Empty,
                    TechnicalName = row.TechnicalName ?? string.Empty,
                    FriendlyName = row.FriendlyName ?? string.Empty
                })
                .Select(g => new OverviewQuantityGroup
                {
                    Type = g.Key.Category,
                    Components = string.Join(", ", g.Select(row => row.Component)),
                    TechnicalName = g.Key.TechnicalName,
                    FriendlyName = g.Key.FriendlyName,
                    Quantity = g.Count()
                })
                .OrderBy(g => g.Type, StringComparer.OrdinalIgnoreCase)
                .ThenBy(g => g.TechnicalName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // Builds a temporary printable HTML document containing the component list table.
        // ###########################################################################################
        public static string BuildPrintableHtml(IEnumerable<OverviewRow> rows)
        {
            static string Encode(string? text) => WebUtility.HtmlEncode(text ?? string.Empty);

            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset=\"utf-8\" />");
            sb.AppendLine("<title>Component List</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("@page { size: A4 portrait; margin: 12mm; }");
            sb.AppendLine("html, body { margin: 0; padding: 0; }");
            sb.AppendLine("body { font-family: Segoe UI, Arial, sans-serif; font-size: 12px; color: #000; box-sizing: border-box; padding: 5px; }");
            sb.AppendLine("table { width: calc(100% - 1px); max-width: calc(100% - 4px); margin-right: 4px; border-collapse: collapse; table-layout: fixed; box-sizing: border-box; }");
            sb.AppendLine("th, td { border: 1px solid #666; padding: 6px 8px; text-align: left; vertical-align: top; word-wrap: break-word; box-sizing: border-box; }");
            sb.AppendLine("th { background: #eaeaea; font-weight: 700; }");
            sb.AppendLine("</style>");
            sb.AppendLine("<script>");
            sb.AppendLine("window.addEventListener('load', function () { window.print(); });");
            sb.AppendLine("</script>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<table>");
            sb.AppendLine("<thead>");
            sb.AppendLine("<tr>");
            sb.AppendLine("<th>Component</th>");
            sb.AppendLine("<th>Technical name</th>");
            sb.AppendLine("<th>Friendly name</th>");
            sb.AppendLine("<th colspan='2'>&nbsp;</th>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</thead>");
            sb.AppendLine("<tbody>");

            foreach (var row in rows)
            {
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td>{Encode(row.Component)}</td>");
                sb.AppendLine($"<td>{Encode(row.TechnicalName)}</td>");
                sb.AppendLine($"<td>{Encode(row.FriendlyName)}</td>");
                sb.AppendLine("<td colspan='2'>&nbsp;</td>");
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }

        // ###########################################################################################
        // Builds a temporary printable HTML document containing the bill of materials table.
        // ###########################################################################################
        public static string BuildPrintableQuantitiesHtml(IEnumerable<OverviewQuantityGroup> groups)
        {
            static string Encode(string? text) => WebUtility.HtmlEncode(text ?? string.Empty);

            var sb = new StringBuilder();

            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset=\"utf-8\" />");
            sb.AppendLine("<title>Bill of Materials (BOM)</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("@page { size: A4 portrait; margin: 12mm; }");
            sb.AppendLine("html, body { margin: 0; padding: 0; }");
            sb.AppendLine("body { font-family: Segoe UI, Arial, sans-serif; font-size: 12px; color: #000; box-sizing: border-box; padding: 5px; }");
            sb.AppendLine("table { width: calc(100% - 1px); max-width: calc(100% - 4px); margin-right: 4px; border-collapse: collapse; table-layout: fixed; box-sizing: border-box; }");
            sb.AppendLine("th, td { border: 1px solid #666; padding: 6px 8px; text-align: left; vertical-align: top; word-wrap: break-word; box-sizing: border-box; }");
            sb.AppendLine("th { background: #eaeaea; font-weight: 700; }");
            sb.AppendLine("</style>");
            sb.AppendLine("<script>");
            sb.AppendLine("window.addEventListener('load', function () { window.print(); });");
            sb.AppendLine("</script>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            sb.AppendLine("<table>");
            sb.AppendLine("<thead>");
            sb.AppendLine("<tr>");
            sb.AppendLine("<th>Type</th>");
            sb.AppendLine("<th>Components</th>");
            sb.AppendLine("<th>Technical name</th>");
            sb.AppendLine("<th>Friendly name</th>");
            sb.AppendLine("<th>Quantity</th>");
            sb.AppendLine("<th>&nbsp;</th>");
            sb.AppendLine("</tr>");
            sb.AppendLine("</thead>");
            sb.AppendLine("<tbody>");

            foreach (var group in groups)
            {
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td>{Encode(group.Type)}</td>");
                sb.AppendLine($"<td>{Encode(group.Components)}</td>");
                sb.AppendLine($"<td>{Encode(group.TechnicalName)}</td>");
                sb.AppendLine($"<td>{Encode(group.FriendlyName)}</td>");
                sb.AppendLine($"<td>{group.Quantity}</td>");
                sb.AppendLine("<td>&nbsp;</td>");
                sb.AppendLine("</tr>");
            }

            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            return sb.ToString();
        }
    }
}
