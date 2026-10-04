using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // WHICH OF A SYSTEM'S SIX VIEWS IS SHOWN (owner request, 2026-10-03; History 2026-10-04) -
    // SystemSections says what each is. This part owns the switch:
    //
    //   - SWITCHING HIDES, IT NEVER CLOSES: the views are siblings whose visibility changes, so the
    //     table - unsaved changes and all - is where it was on coming back to Board data.
    //   - THE VIEW CHOSEN STAYS for the next system chosen (SystemSections' header says why).
    //   - BOARD DATA AND FILES ARE READ FROM THE SERVER when first shown for a system - not with
    //     every system chosen, since a board's workbook is read for both - and, once read, stay
    //     until another system is chosen OR BETA MOVES UNDER THEM (owner report, 2026-10-04: a fix
    //     published, and promoted, went on showing the warning it fixed until CRT was restarted).
    //     The system's own row says when (CatchUpWithBetaAsync). The other four come with the
    //     system's detail.
    // ###########################################################################################
    public partial class SystemView
    {
        private SystemSection thisSection = SystemSections.Opening;

        // The view on show, for tests.
        internal SystemSection ShownSection => this.thisSection;

        private static readonly (SystemSection Section, string Button, string Panel)[] SectionParts =
        [
            (SystemSection.BoardData, "BoardDataSectionButton", "BoardDataSection"),
            (SystemSection.Files, "FilesSectionButton", "FilesSection"),
            (SystemSection.Contributor, "ContributorSectionButton", "ContributorSection"),
            (SystemSection.Maintainer, "MaintainerSectionButton", "MaintainerSection"),
            (SystemSection.History, "HistorySectionButton", "HistorySection"),
            (SystemSection.Statistics, "StatisticsSectionButton", "StatisticsSection")
        ];

        private async void OnSectionClick(object? sender, RoutedEventArgs e)
        {
            if (sender is not Button button)
                return;

            foreach ((SystemSection section, string name, _) in SystemView.SectionParts)
            {
                if (string.Equals(button.Name, name, StringComparison.Ordinal))
                {
                    await this.ShowSectionAsync(section);
                    return;
                }
            }
        }

        // ###########################################################################################
        // Shows one view, reading it first when it needs the server and has not been read for this
        // system - under the window's "please wait", since the maintainer pressed for it.
        // ###########################################################################################
        internal async Task ShowSectionAsync(SystemSection section)
        {
            this.thisSection = section;
            this.ApplySection();

            if (this.ShownSystem is not { } system || !this.NeedsReading(section))
                return;

            string wait = section == SystemSection.Files
                ? MaintainerWaitWording.ReadingSystemFiles(system.SystemId)
                : MaintainerWaitWording.ReadingSystemTable(system.SystemId);

            if (!await ServerWait.RunAsync(this, wait, this.LoadShownSectionAsync))
                this.ShowMessage(WaitWording.NoAnswer, isError: true);
        }

        private void ApplySection()
        {
            foreach ((SystemSection section, string button, string panel) in SystemView.SectionParts)
            {
                bool shown = section == this.thisSection;

                if (this.FindControl<Button>(button) is Button sectionButton)
                    sectionButton.Classes.Set("Selected", shown);

                this.SetShown(panel, shown);
            }

            // The BETA / Stable switch shows with Board data and Files (SystemView.Stable.cs).
            this.ApplyTree();
        }

        // ###########################################################################################
        // Whether the view on show has still to be read from the server for the system on screen -
        // not read yet, or read before BETA moved. A table holding a change not published yet is
        // never read again under it; undone, the change is gone and it is.
        // ###########################################################################################
        private bool NeedsReading(SystemSection section)
        {
            // The stable source's half has its own table and files (SystemView.Stable.cs).
            if (this.ShowsStable)
                return this.StableNeedsReading(section);

            string? systemId = this.ShownSystem?.SystemId;

            return section switch
            {
                SystemSection.BoardData => !this.HoldsTableFor(systemId) || (!this.HasUnsavedTableEdits && this.TableIsBehind()),
                SystemSection.Files => !this.HoldsFilesFor(systemId) || this.FilesAreBehind(),
                _ => false
            };
        }

        private bool TableIsBehind() =>
            this.thisTableReadAt is { } readAt && this.ShownSystem is { } now &&
            QueueRefreshRules.SystemTableChanged(readAt, now, this.thisTableReadAtMaintainers, this.thisShownMaintainers);

        private bool FilesAreBehind() =>
            this.thisFilesReadAt is { } readAt && this.ShownSystem is { } now &&
            QueueRefreshRules.BetaBoardChanged(readAt, now);

        // ###########################################################################################
        // *** BETA MOVED UNDER WHAT IS SHOWN (owner report, 2026-10-04). *** Told the system as it
        // is now each time its row is read again - by the queue's own check, or after a publish, a
        // push-back or a promotion here. A table read before the move is read again, quietly, on
        // the sheet it was on; a files list too when on screen, and otherwise dropped, to be read
        // when next shown. A table holding a change not published yet stays - replacing it would
        // throw the change away - and says the change can no longer go (the server would refuse it).
        //
        // A table read just after a publish of its own takes this reading as the state it was read
        // at (thisTableReadAt null): reading it again would only replace the line saying what the
        // publish did.
        // ###########################################################################################
        private async Task CatchUpWithBetaAsync(SystemOverviewEntry now)
        {
            if (this.HoldsTableFor(now.SystemId))
            {
                if (this.thisTableReadAt is null)
                {
                    this.thisTableReadAt = now;
                    this.thisTableReadAtMaintainers = this.thisShownMaintainers;
                }
                else if (QueueRefreshRules.SystemTableChanged(
                    this.thisTableReadAt, now, this.thisTableReadAtMaintainers, this.thisShownMaintainers))
                {
                    if (this.HasUnsavedTableEdits)
                        this.ShowTableNote(SystemSections.BetaMovedUnderChange);
                    else
                        await this.LoadTableAsync();
                }
            }

            if (this.HoldsFilesFor(now.SystemId) &&
                this.thisFilesReadAt is { } filesReadAt &&
                QueueRefreshRules.BetaBoardChanged(filesReadAt, now))
            {
                if (this.thisSection == SystemSection.Files && !this.ShowsStable)
                    await this.LoadFilesAsync();
                else
                    this.ClearFiles();
            }

            // And the stable source's, which only a promotion moves (SystemView.Stable.cs).
            await this.CatchUpWithStableAsync(now);
        }

        // The same, told by a test without a server.
        internal Task CatchUpWithBetaForTests(SystemOverviewEntry now) => this.CatchUpWithBetaAsync(now);

        // Reads the view on show when it needs the server - the caller holds any wait.
        private Task LoadShownSectionAsync() => (this.thisSection, this.ShowsStable) switch
        {
            (SystemSection.BoardData, false) => this.LoadTableAsync(),
            (SystemSection.Files, false) => this.LoadFilesAsync(),
            (SystemSection.BoardData, true) => this.LoadStableTableAsync(),
            (SystemSection.Files, true) => this.LoadStableFilesAsync(),
            _ => Task.CompletedTask
        };

        // ###########################################################################################
        // Another system - or none: nothing of the previous one's board or files stays, so neither
        // can be read as this one's. The caller has asked about unsaved table changes.
        // ###########################################################################################
        private void ForgetSystemContent()
        {
            this.CloseTableNow();
            this.ClearFiles();
            this.ForgetStableContent();
            this.thisShownMaintainers = null;
            this.thisStageSubmissions = null;
            this.thisStageSubmissionsFor = null;
        }
    }
}
