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
