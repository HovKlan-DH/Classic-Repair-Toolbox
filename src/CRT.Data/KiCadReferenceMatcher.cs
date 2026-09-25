using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What an imported KiCad project actually lines up with on the board - which board labels found
    // a matching footprint reference and, far more importantly, which did NOT
    // (NewContributeStrategy.md Phase 2, session 2c, task 9).
    //
    // This report exists because the failure it describes is completely silent today. A component
    // labelled "U17" finds its copper by looking for a footprint named "U17"; call it "PLA/U17" in
    // the board data and nothing highlights, with no error anywhere. The Wiki
    // (Add-new-board-with-KiCad-data.md) calls this "the number one cause of 'I did everything and
    // no traces appear'".
    // ###########################################################################################
    public sealed class KiCadReferenceMatchReport
    {
        // Board labels that DO have a footprint - these will light up copper.
        public IReadOnlyList<string> MatchedLabels { get; }

        // Board labels with no footprint of that name. The headline: every one of these is a
        // component that will highlight on the image and light up nothing.
        public IReadOnlyList<string> UnmatchedLabels { get; }

        // Footprints with no board label yet - not an error at all, but the to-do list for a board
        // being built up, and free from the same comparison.
        public IReadOnlyList<string> UnusedReferences { get; }

        // How many PCB files the project holds. More than one matters: only the FIRST is ever used
        // for trace matching, which is invisible in the app today.
        public int PcbFileCount { get; }

        // The view names the project generates, which are what the "CAD name" column has to carry.
        // The Wiki currently tells a contributor to read these out of the logfile by hand.
        public IReadOnlyList<string> ViewDisplayNames { get; }

        public KiCadReferenceMatchReport(
            IReadOnlyList<string> matchedLabels,
            IReadOnlyList<string> unmatchedLabels,
            IReadOnlyList<string> unusedReferences,
            int pcbFileCount,
            IReadOnlyList<string> viewDisplayNames)
        {
            this.MatchedLabels = matchedLabels;
            this.UnmatchedLabels = unmatchedLabels;
            this.UnusedReferences = unusedReferences;
            this.PcbFileCount = pcbFileCount;
            this.ViewDisplayNames = viewDisplayNames;
        }

        // ###########################################################################################
        // One plain sentence for the top of the report. Leads with the problem when there is one,
        // because that is the whole reason this report exists - a summary that opened with the good
        // news would bury it.
        // ###########################################################################################
        public string Summary
        {
            get
            {
                if (this.PcbFileCount == 0)
                {
                    return "No PCB file was found in this KiCad data, so no component can light up any copper.";
                }

                int total = this.MatchedLabels.Count + this.UnmatchedLabels.Count;

                if (total == 0)
                {
                    return this.UnusedReferences.Count == 1
                        ? "No components are labelled yet. The KiCad data has 1 component to match against."
                        : $"No components are labelled yet. The KiCad data has {this.UnusedReferences.Count} components to match against.";
                }

                if (this.UnmatchedLabels.Count == 0)
                {
                    return total == 1
                        ? "All 1 labelled component matches a component in the KiCad data."
                        : $"All {total} labelled components match a component in the KiCad data.";
                }

                string labelWord = this.UnmatchedLabels.Count == 1 ? "component" : "components";
                return $"{this.UnmatchedLabels.Count} of {total} labelled {labelWord} found no match in the KiCad data - those will not light up any copper.";
            }
        }

        // True when there is something the contributor should act on.
        public bool HasProblems => this.PcbFileCount != 1 || this.UnmatchedLabels.Count > 0;
    }

    // ###########################################################################################
    // Compares a board's component labels against a KiCad project's footprint references.
    //
    // IT MUST REPRODUCE THE RENDER-TIME RULE EXACTLY, or the report lies about what will actually
    // light up on screen. That rule lives in TabSchematics.KiCad.cs's
    // BuildKiCadNormalizedNetNamesForReferences, which is the ONLY place a board label becomes
    // copper, and it:
    //   - reads root.Pcb[0] and nothing else (other PCB files in the same project are ignored);
    //   - compares footprint.Reference, trimmed, case-insensitively.
    // So this compares against the FIRST PCB file's footprint references, trimmed, case-
    // insensitively. Matching against schematic symbols instead, or unioning every PCB file, would
    // report matches that highlight nothing - exactly the silent failure the report exists to end.
    //
    // Takes plain strings rather than a KiCadProjectRoot deliberately: the KiCad parsing types live
    // in CRT.App and are internal to it, and CRT.Data must never reference CRT.App (that would
    // invert the dependency and follow a desktop UI framework into the future server). The caller
    // extracts the strings; this owns the set logic. The same "resolve the reads at the call site
    // and hand over plain values" split LabelEditorSnapContext already follows.
    // ###########################################################################################
    public static class KiCadReferenceMatcher
    {
        public static KiCadReferenceMatchReport Build(
            IEnumerable<string> boardLabels,
            IEnumerable<string> footprintReferencesOfFirstPcb,
            int pcbFileCount,
            IEnumerable<string>? viewDisplayNames = null)
        {
            var references = Normalize(footprintReferencesOfFirstPcb);
            var labels = Normalize(boardLabels);

            var matched = new List<string>();
            var unmatched = new List<string>();

            // NATURAL order in every list (C1, C2, C10 - not C1, C10, C100, C2): each is shown in
            // full, one badge per component, and a board runs to hundreds of them, so the order has
            // to be the one a person counts in (maintainer request, 2026-09-24).
            foreach (string label in labels.OrderBy(label => label, NaturalLabelComparer.Instance))
            {
                if (references.Contains(label))
                {
                    matched.Add(label);
                }
                else
                {
                    unmatched.Add(label);
                }
            }

            var unused = references
                .Where(reference => !labels.Contains(reference))
                .OrderBy(reference => reference, NaturalLabelComparer.Instance)
                .ToList();

            return new KiCadReferenceMatchReport(
                matched,
                unmatched,
                unused,
                pcbFileCount,
                (viewDisplayNames ?? Enumerable.Empty<string>())
                    .Select(name => name?.Trim() ?? string.Empty)
                    .Where(name => name.Length > 0)
                    .ToList());
        }

        // Trimmed and de-duplicated case-insensitively, mirroring the render-time comparison. A
        // board label appearing twice (two rows for the same component, one per region) is one
        // component as far as "does this find copper" is concerned.
        private static HashSet<string> Normalize(IEnumerable<string> values) =>
            values
                .Select(value => value?.Trim() ?? string.Empty)
                .Where(value => value.Length > 0)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }
}
