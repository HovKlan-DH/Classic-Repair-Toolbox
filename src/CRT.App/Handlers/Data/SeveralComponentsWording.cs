using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE COMPONENT PANE SAYS WITH SEVERAL COMPONENTS PICKED (owner request, 2026-11-08: "It must be very clear for
    // the user, that whatever the user changes will affect all selected components ... Ideally it should be possible to
    // see if the data is already identical for all components, and which data differs"; ComponentPane.Several.cs). The
    // pages are named as their entries are, in quotes.
    // ###########################################################################################
    internal static class SeveralComponentsWording
    {
        // How many labels a line names before it counts the rest.
        public const int NamedAtMost = 6;

        // "R21, R22, R23" - past NamedAtMost, the first of them and how many more.
        public static string Labels(IReadOnlyList<string> labels)
        {
            ArgumentNullException.ThrowIfNull(labels);

            string named = string.Join(", ", labels.Take(NamedAtMost));

            return labels.Count > NamedAtMost
                ? string.Create(CultureInfo.InvariantCulture, $"{named} and {labels.Count - NamedAtMost} more")
                : named;
        }

        public static string Title(IReadOnlyList<string> labels) => "Components: " + SeveralComponentsWording.Labels(labels);

        // ###########################################################################################
        // How many and who, the number and each name in bold, nothing else (owner requests, 2026-11-09: "show the amount of
        // components and the component names with bold", then "Editing <b>2</b> components: <b>RP7</b>,<b>RP8</b>") -
        // BoldRuns.ShowMarked shows it.
        // ###########################################################################################
        public static string Editing(IReadOnlyList<string> labels)
        {
            ArgumentNullException.ThrowIfNull(labels);

            // A label is typed data: its own "<b>" or "</b>" taken out, or it would bold the wrong pieces (code review,
            // 2026-11-09 - as CloneFacts does).
            string named = string.Join(", ", labels.Take(NamedAtMost).Select(label => $"<b>{Handlers.Theming.BoldRuns.Plain(label)}</b>"));
            string more = labels.Count > NamedAtMost
                ? string.Create(CultureInfo.InvariantCulture, $" and {labels.Count - NamedAtMost} more")
                : string.Empty;

            return string.Create(CultureInfo.InvariantCulture, $"Editing <b>{labels.Count}</b> components: {named}{more} - every change you make here is made to all of them.");
        }

        // In a box whose value differs between them - empty, so nothing is set until something is typed or picked.
        public static string Differs(int values) =>
            string.Create(CultureInfo.InvariantCulture, $"Differs - {values} values");

        // Over the values listed under such a box while it has the focus.
        public const string DiffersHeading = "What they have now - pick one to set it on all of them, or type to narrow the list:";

        // A list page: what only some have - which stays as it is. Nothing to say when they all have the same.
        public static string RowsLine(ComponentPage page, IReadOnlyList<(string Label, string Row)> some)
        {
            ArgumentNullException.ThrowIfNull(some);

            if (some.Count == 0)
            {
                return string.Empty;
            }

            string listed = string.Join("; ", some.Take(NamedAtMost).Select(each => $"{each.Label}: {each.Row}"));
            string more = some.Count > NamedAtMost
                ? string.Create(CultureInfo.InvariantCulture, $" and {some.Count - NamedAtMost} more")
                : string.Empty;

            return $"Listed: what all of them have - a change reaches every one. Only some have {listed}{more} - each left as it is.";
        }

        // A list page with no row they all have.
        public const string NoneInCommon = "None that all of them have.";

        public const string BaselinesHeld =
            "An oscilloscope baseline is measured on one component - pick one of them on its own to see or change its baselines.";

        public static string PackagesDiffer(IEnumerable<string> packages)
        {
            string named = string.Join(", ", packages.Select(package => package.Length > 0 ? package : "none").Distinct(StringComparer.Ordinal));

            return $"Their packages differ ({named}) - give them one package in \"Package\" below to edit their pins together, or pick one of them on its own to change its pins.";
        }

        // ###########################################################################################
        // THE BUTTON SAYS IT CHANGES THE PAGE SHOWN ALONE, BY A FOOTNOTE (owner request, 2026-11-09: "maybe change the button
        // to "Make all have same data as [RP7] *1" ... "*1 Pressing button will only change data on this page." (word this
        // better, if you have any ideas)"). The pane draws the mark small, bold and raised.
        // ###########################################################################################
        public const string FootnoteMark = "*1";

        public static string MakeAllLike(string source) => $"Make all have the same data as [{source}]";

        public static string MakeAllLikeNote(ComponentPage page) =>
            $"The button changes only the data on this page, \"{ComponentPaneWording.PageName(page)}\" - the components' other pages stay as they are.";

        public static string MakeAllLikeTip(ComponentPage page, string source) =>
            $"Give every one of them what {source} has on \"{ComponentPaneWording.PageName(page)}\", in place of their own - one step, which \"Undo\" takes back";

        // A row of each list, as the line of what only some have names it.
        public static string Row(ComponentReplacementEntry row) => SeveralComponentsWording.Joined(row.PartNumber, row.Kind);

        public static string Row(ComponentPinEntry row) => SeveralComponentsWording.Joined("pin " + row.Pin?.Trim(), row.Name);

        public static string Row(ComponentImageEntry row) => SeveralComponentsWording.Joined(row.Name, row.File);

        public static string Row(ComponentLocalFileEntry row) => SeveralComponentsWording.Joined(row.Name, row.File);

        public static string Row(ComponentLinkEntry row) => SeveralComponentsWording.Joined(row.Name, row.Url);

        private static string Joined(string? first, string? second)
        {
            string a = first?.Trim() ?? string.Empty;
            string b = second?.Trim() ?? string.Empty;

            return a.Length == 0 ? b : b.Length == 0 ? a : $"{a} ({b})";
        }
    }
}
