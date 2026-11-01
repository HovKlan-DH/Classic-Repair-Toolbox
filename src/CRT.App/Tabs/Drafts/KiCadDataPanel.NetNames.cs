using Handlers.DataHandling;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE BOARD'S KiCad NETS, IN THE PANEL RATHER THAN THE LOGFILE (owner request, 2026-10-10).
    //
    // The names are what the logfile listed as "normalized KiCad net name [...]". Each is a badge, and
    // one an important signal already uses is marked. A row naming a net the KiCad data does not have
    // is listed above the names: the Schematics tab's panel leaves such a row out without a word, the
    // same silent failure as a component that will not light up.
    //
    // *** NOTHING IS COPIED FROM HERE ANY MORE (owner request, 2026-11-01: "drop that 'copy' option from
    // the 'KiCad data' ... as this will not be useful"). *** A net name was copied to paste it into the
    // table; every net is now listed, with its "Add", in the Important KiCad signals part.
    //
    // The rule for which row finds a net is KiCadNetNames' - the panel's own.
    //
    // Part of the KiCadDataPanel partial class - see KiCadDataPanel.axaml.cs for the file map.
    // ###########################################################################################
    public partial class KiCadDataPanel
    {
        internal void ShowNetNameReport(KiCadNetNameReport report)
        {
            this.NetNamesPanel.IsVisible = report.NetNames.Count > 0;
            this.NetNamesHeaderText.Text = report.NetNames.Count == 1
                ? "1 KiCad net"
                : $"{report.NetNames.Count} KiCad nets";

            int usedCount = report.NetNames.Count(report.IsUsed);

            this.NetNamesIntroText.Text =
                "The nets of the board's KiCad data. Add the ones worth checking first under \"KiCad: Important signals\"." +
                KiCadImportedFiles.UsedNetNamesSentence(usedCount);

            this.NetNamesList.ItemsSource = report.NetNames
                .Select(name => new KiCadNetNameBadge { Name = name, IsUsed = report.IsUsed(name) })
                .ToList();

            this.UnresolvedSignalsPanel.IsVisible = report.UnresolvedSignals.Count > 0;
            this.UnresolvedSignalsHeaderText.Text = report.UnresolvedSignals.Count == 1
                ? "1 important KiCad signal will NOT light up"
                : $"{report.UnresolvedSignals.Count} important KiCad signals will NOT light up";
            this.UnresolvedSignalsList.ItemsSource = report.UnresolvedSignals
                .Select(signal => $"{signal.DisplayName}: {signal.KiCadNetName}")
                .ToList();
        }
    }
}
