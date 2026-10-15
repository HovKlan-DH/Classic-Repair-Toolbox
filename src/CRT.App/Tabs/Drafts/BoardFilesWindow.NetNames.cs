using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Handlers.DataHandling;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE KiCad NET NAMES, IN THE WINDOW RATHER THAN THE LOGFILE (owner request, 2026-10-10: "show
    // (in UI) the same data as is visible in the logfile ... It should be able to copy each of
    // these names, so they can be pasted into the Excel").
    //
    // The names are what the logfile listed as "normalized KiCad net name [...]" - what a row of
    // the board's "Important signals" sheet puts in its "KiCad net name" column. Each is a badge;
    // a click copies it. A name a row already uses is marked, and a row naming a net the KiCad data
    // does not have is listed above the names: the Schematics tab's panel leaves such a row out
    // without a word, the same silent failure as a component that will not light up.
    //
    // The rule for which row finds a net is KiCadNetNames' - the panel's own.
    //
    // Part of the BoardFilesWindow partial class - see BoardFilesWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardFilesWindow
    {
        internal void ShowNetNameReport(KiCadNetNameReport report)
        {
            this.NetNamesPanel.IsVisible = report.NetNames.Count > 0;
            this.NetNamesHeaderText.Text = report.NetNames.Count == 1
                ? "1 KiCad net name you can use"
                : $"{report.NetNames.Count} KiCad net names you can use";

            int usedCount = report.NetNames.Count(report.IsUsed);

            this.NetNamesIntroText.Text =
                "For the \"KiCad net name\" column of the board's \"Important signals\" sheet. Click a name to copy it, " +
                "then paste it into the sheet - in Excel, or in the Drafts tab's table." +
                (usedCount == 0
                    ? string.Empty
                    : usedCount == 1
                        ? " The one in green is already used by an important signal."
                        : $" The {usedCount} in green are already used by an important signal.");

            this.NetNamesList.ItemsSource = report.NetNames
                .Select(name => new KiCadNetNameBadge { Name = name, IsUsed = report.IsUsed(name) })
                .ToList();

            this.UnresolvedSignalsPanel.IsVisible = report.UnresolvedSignals.Count > 0;
            this.UnresolvedSignalsHeaderText.Text = report.UnresolvedSignals.Count == 1
                ? "1 important signal will NOT light up"
                : $"{report.UnresolvedSignals.Count} important signals will NOT light up";
            this.UnresolvedSignalsList.ItemsSource = report.UnresolvedSignals
                .Select(signal => $"{signal.DisplayName}: {signal.KiCadNetName}")
                .ToList();
        }

        private async void OnNetNameTapped(object? sender, TappedEventArgs e)
        {
            try
            {
                if (sender is Control { DataContext: KiCadNetNameBadge badge })
                {
                    e.Handled = true;
                    await this.CopyNetNameAsync(badge.Name);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnNetNameTapped failed - [{ex}]");
            }
        }

        // The keyboard's way to the same copy (code review, 2026-10-13): a badge reached with Tab,
        // then Enter or Space. CopyNetNameAsync reports its own failures, so nothing escapes.
        private async void OnNetNameKeyDown(object? sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key is (Key.Enter or Key.Space) && sender is Control { DataContext: KiCadNetNameBadge badge })
                {
                    e.Handled = true;
                    await this.CopyNetNameAsync(badge.Name);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnNetNameKeyDown failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // Puts one net name on the clipboard and says so on the status line. A clipboard another
        // program holds open throws on Windows - that is a line on the status bar, never a crash.
        // ###########################################################################################
        internal async Task CopyNetNameAsync(string netName)
        {
            IClipboard? clipboard = TopLevel.GetTopLevel(this)?.Clipboard;

            if (clipboard is null)
            {
                this.ShowStatus("The clipboard could not be reached.", isError: true);
                return;
            }

            try
            {
                await clipboard.SetTextAsync(netName);
                this.ShowStatus($"Copied [{netName}] - paste it into the \"KiCad net name\" column of the \"Important signals\" sheet.");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not copy the KiCad net name [{netName}] - [{ex.Message}]");
                this.ShowStatus("Could not copy - another program may be holding the clipboard. Try again.", isError: true);
            }
        }
    }
}
