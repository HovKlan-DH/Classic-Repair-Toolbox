using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A KiCad project's net names as the board's "Important signals" sheet names them, and which of
    // that sheet's rows find no net (owner request, 2026-10-10: show in the KiCad data window what
    // the logfile listed as "normalized KiCad net name [...]", each name copyable into the workbook).
    // ###########################################################################################
    public sealed class KiCadNetNameReport
    {
        // Every normalised net name in the project, once each, in natural order (A1, A2, A10).
        public IReadOnlyList<string> NetNames { get; }

        // The normalised names an "Important signals" row already finds.
        public IReadOnlySet<string> UsedNetNames { get; }

        // The "Important signals" rows whose "KiCad net name" is no net in the project - each one is
        // a signal the Schematics tab's panel drops without a word. In sheet order.
        public IReadOnlyList<KiCadImportantSignalEntry> UnresolvedSignals { get; }

        public KiCadNetNameReport(
            IReadOnlyList<string> netNames,
            IReadOnlySet<string> usedNetNames,
            IReadOnlyList<KiCadImportantSignalEntry> unresolvedSignals)
        {
            this.NetNames = netNames;
            this.UsedNetNames = usedNetNames;
            this.UnresolvedSignals = unresolvedSignals;
        }

        public bool IsUsed(string netName) => this.UsedNetNames.Contains(netName);
    }

    // ###########################################################################################
    // How an "Important signals" row finds its net.
    //
    // *** ONE RULE WITH THE SCHEMATICS TAB. *** Its "Important signals" panel
    // (TabSchematics.RebuildImportantSignalsPanel) resolves a row's "KiCad net name" through
    // Resolver below - the net's own name or its normalised name, trimmed, case-insensitively,
    // across EVERY PCB file - so the KiCad data window cannot call a row good that the panel drops,
    // or the other way round.
    //
    // Plain strings in, like KiCadReferenceMatcher: the KiCad types live in CRT.App.
    // ###########################################################################################
    public static class KiCadNetNames
    {
        // Every name a net is known by -> its normalised name. A net with no normalised name is no
        // net at all here, as it is to the panel.
        public static Dictionary<string, string> Resolver(IEnumerable<(string? Name, string? NormalizedName)> nets)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach ((string? name, string? normalizedName) in nets)
            {
                string normalized = normalizedName?.Trim() ?? string.Empty;

                if (normalized.Length == 0)
                {
                    continue;
                }

                result[normalized] = normalized;

                string raw = name?.Trim() ?? string.Empty;

                if (raw.Length > 0)
                {
                    result[raw] = normalized;
                }
            }

            return result;
        }

        // ###########################################################################################
        // Every normalised net name, once each (case-insensitively), in natural order - the list a
        // contributor picks from, so in the order a person counts in.
        //
        // *** BUT NOT KiCad's "NO NET" (owner decision, 2026-10-10: "yes, hide it"). *** Net 0, what
        // an unconnected pad is on - nobody lights up "no net". The LOADER says which net that is
        // (KiCadNetRef.IsNoNet: an older file's declaration (net 0 "")), because nothing else can:
        // a net with no name is named by its number, so it arrives as id "0", name "0" - and so does
        // a real KiCad 9 net NAMED "0" (SPICE's ground), since KiCad 9 refers to every net by its
        // name alone, (net "0"). It was told apart by id and name at first, which hid that one too
        // (code review, 2026-10-10). The resolver still knows net 0, as the panel always has.
        // ###########################################################################################
        public static List<string> NormalizedNames(IEnumerable<(string? Name, string? NormalizedName, bool IsNoNet)> nets) =>
            nets
                .Where(net => !net.IsNoNet)
                .Select(net => net.NormalizedName?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, NaturalLabelComparer.Instance)
                .ToList();

        // ###########################################################################################
        // The report for one project and one board's "Important signals" rows.
        //
        // A row missing its display name or its net name is skipped, as the panel skips it: half a
        // row lights up nothing either way, and the table's own checks are what point at it.
        // ###########################################################################################
        public static KiCadNetNameReport Build(
            IEnumerable<(string? Name, string? NormalizedName, bool IsNoNet)> nets,
            IEnumerable<KiCadImportantSignalEntry> importantSignals)
        {
            List<(string? Name, string? NormalizedName, bool IsNoNet)> netList = nets.ToList();
            Dictionary<string, string> resolver = KiCadNetNames.Resolver(netList.Select(net => (net.Name, net.NormalizedName)));

            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var unresolved = new List<KiCadImportantSignalEntry>();

            foreach (KiCadImportantSignalEntry entry in importantSignals)
            {
                string displayName = entry.DisplayName?.Trim() ?? string.Empty;
                string netName = entry.KiCadNetName?.Trim() ?? string.Empty;

                if (displayName.Length == 0 || netName.Length == 0)
                {
                    continue;
                }

                if (resolver.TryGetValue(netName, out string? normalized))
                {
                    used.Add(normalized);
                }
                else
                {
                    unresolved.Add(new KiCadImportantSignalEntry { DisplayName = displayName, KiCadNetName = netName });
                }
            }

            return new KiCadNetNameReport(KiCadNetNames.NormalizedNames(netList), used, unresolved);
        }
    }
}
