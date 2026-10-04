using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE BETA / STABLE SWITCH ON BOARD DATA AND FILES (owner request, 2026-10-04: "Shouldn't there be
    // somewhere a possibility to see what we actually do have in BETA or stable for this?").
    //
    // BETA's table and file tree are SystemView.Table.cs' and .Files.cs' - the table the one a change
    // is made in. The stable source has a table and a file tree OF ITS OWN beside them, here: the
    // table read-only (the server says why - only a publish from BETA changes it), the tree opening
    // each file from the stable source's public address. So switching hides one and shows the other,
    // and a change in BETA's table is never touched - the tab's own "hide, never close" rule.
    //
    // WHICH IS SHOWN is SystemSections.ShowsStable: the one wanted, when the system is there; else
    // the stable source when only it holds the system; else BETA, which then says nothing is there.
    // The one wanted stays for the next system, as the chosen view does.
    //
    // READ when first shown for a system, like BETA's; read again when the stable source moved under
    // it - only a promotion moves it (QueueRefreshRules.StableBoardChanged) - quietly when on screen,
    // and otherwise dropped, to be read when next shown.
    // ###########################################################################################
    public partial class SystemView
    {
        // The maintainer's pick: the stable source, or BETA (the default).
        private bool thisStableWanted;

        // The stable source's board, and the system as it was when it was read.
        private SystemTableAnswer? thisStableTable;
        private SystemOverviewEntry? thisStableTableReadAt;

        // The system the stable files are held for, and as it was when they were read.
        private string? thisStableFilesFor;
        private SystemOverviewEntry? thisStableFilesReadAt;

        // The sheet last looked at in each system's stable table, for as long as CRT runs.
        private readonly Dictionary<string, string> thisStableSheetBySystem = new(StringComparer.Ordinal);

        private BoardTableEditor StableBoardTable => this.FindControl<BoardTableEditor>("StableTable")!;

        private FileTreeView StableFiles => this.FindControl<FileTreeView>("StableFileTree")!;

        // Whether the stable source's half is the one on screen for the system shown.
        private bool ShowsStable => SystemSections.ShowsStable(this.thisStableWanted, this.ShownSystem);

        internal bool ShowsStableForTests => this.ShowsStable;

        internal BoardTableEditor StableTableForTests => this.StableBoardTable;

        internal FileTreeView StableFileTreeForTests => this.StableFiles;

        // The stable table, for the picked pills the Maintainer tab's tables share
        // (TabMaintainer.UseRememberedChoices).
        internal BoardTableEditor StableTableEditorForSharedChoices => this.StableBoardTable;

        private async void OnTreeClick(object? sender, RoutedEventArgs e)
        {
            if (sender is not Button button)
                return;

            await this.ShowTreeAsync(string.Equals(button.Name, "StableTreeButton", StringComparison.Ordinal));
        }

        // ###########################################################################################
        // The maintainer's pick, then the view on show read when it needs the server - under the
        // window's "please wait", since the maintainer pressed for it.
        // ###########################################################################################
        internal async Task ShowTreeAsync(bool stable)
        {
            this.thisStableWanted = stable;
            this.ApplyTree();

            if (this.ShownSystem is not { } system || !this.NeedsReading(this.thisSection))
                return;

            string wait = this.thisSection == SystemSection.Files
                ? MaintainerWaitWording.ReadingSystemFiles(system.SystemId)
                : MaintainerWaitWording.ReadingSystemTable(system.SystemId);

            if (!await ServerWait.RunAsync(this, wait, this.LoadShownSectionAsync))
                this.ShowMessage(WaitWording.NoAnswer, isError: true);
        }

        // ###########################################################################################
        // The switch and the two halves, from the pick and the system on screen: the switch shows
        // with Board data and Files only, a half the system is not in cannot be chosen, and the half
        // shown is the one ShowsStable says.
        // ###########################################################################################
        private void ApplyTree()
        {
            SystemOverviewEntry? system = this.ShownSystem;
            bool stable = this.ShowsStable;

            this.SetShown(
                "TreeSwitchBar",
                system is not null && this.thisSection is SystemSection.BoardData or SystemSection.Files);

            if (this.FindControl<Button>("BetaTreeButton") is Button beta)
            {
                beta.Classes.Set("Selected", !stable);
                beta.IsEnabled = SystemSections.CanChooseBeta(system);
            }

            if (this.FindControl<Button>("StableTreeButton") is Button stableButton)
            {
                stableButton.Classes.Set("Selected", stable);
                stableButton.IsEnabled = SystemSections.CanChooseStable(system);
            }

            this.SetShown("BetaBoardPart", !stable);
            this.SetShown("StableBoardPart", stable);
            this.SetShown("BetaFilesPart", !stable);
            this.SetShown("StableFilesPart", stable);
        }

        private bool HoldsStableTableFor(string? systemId) =>
            this.thisStableTable is not null && string.Equals(this.thisStableTable.SystemId, systemId, StringComparison.Ordinal);

        private bool HoldsStableFilesFor(string? systemId) =>
            this.thisStableFilesFor is not null && string.Equals(this.thisStableFilesFor, systemId, StringComparison.Ordinal);

        private bool StableTableIsBehind() =>
            this.thisStableTableReadAt is { } readAt && this.ShownSystem is { } now && QueueRefreshRules.StableBoardChanged(readAt, now);

        private bool StableFilesAreBehind() =>
            this.thisStableFilesReadAt is { } readAt && this.ShownSystem is { } now && QueueRefreshRules.StableBoardChanged(readAt, now);

        // Whether the stable half of the view on show has still to be read for the system on screen.
        private bool StableNeedsReading(SystemSection section)
        {
            string? systemId = this.ShownSystem?.SystemId;

            return section switch
            {
                SystemSection.BoardData => !this.HoldsStableTableFor(systemId) || this.StableTableIsBehind(),
                SystemSection.Files => !this.HoldsStableFilesFor(systemId) || this.StableFilesAreBehind(),
                _ => false
            };
        }

        // ###########################################################################################
        // Reads the stable source's board for the system on screen and opens its read-only table. An
        // answer for a system no longer chosen is dropped. The caller holds the wait.
        // ###########################################################################################
        private async Task LoadStableTableAsync()
        {
            if (this.ShownSystem is not { } system || !this.CanReadStable)
                return;

            WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), null, isError: false);

            ReviewApiResult<SystemTableAnswer> result = await this.ReadStableTableAsync(system.SystemId);

            if (!string.Equals(this.ShownSystem?.SystemId, system.SystemId, StringComparison.Ordinal))
                return;

            if (!result.IsOk)
            {
                if (result.Failure == ReviewApiFailure.NotFound)
                    this.CloseStableTable();

                // "Not in the stable source" is an answer about the system, not a failure.
                WindowMessage.Show(
                    this.FindControl<TextBlock>("StableTableMessageText"),
                    result.Message,
                    isError: result.Failure != ReviewApiFailure.NotFound);

                return;
            }

            // An older server answered with BETA's board: never drawn as the stable source's.
            if (!SystemSections.IsStableAnswer(result.Value!))
            {
                this.CloseStableTable();
                WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), SystemSections.StableNeedsNewerServer, isError: true);
                return;
            }

            this.OpenStableTable(result.Value!);
            this.thisStableTableReadAt = system;
        }

        private bool CanReadStable =>
            this.ReadStableTableOverrideForTests is not null || (this.thisClient is not null && this.thisSession is not null);

        private Task<ReviewApiResult<SystemTableAnswer>> ReadStableTableAsync(string systemId) =>
            this.ReadStableTableOverrideForTests is { } read
                ? read(systemId)
                : this.thisClient!.GetSystemTableAsync(this.thisSession!, systemId, DataTreeNames.Production);

        // Reads a system's stable table without a server - for tests.
        internal Func<string, Task<ReviewApiResult<SystemTableAnswer>>>? ReadStableTableOverrideForTests { get; set; }

        private void OpenStableTable(SystemTableAnswer table)
        {
            BoardTableEditor editor = this.StableBoardTable;

            if (this.thisStableTable is not null && editor.CurrentSheet?.Name is string current)
                this.thisStableSheetBySystem[this.thisStableTable.SystemId] = current;

            this.thisStableTable = table;

            if (this.thisClient is ReviewApiClient client)
                editor.FileSource = new SystemTableFileSource(client, table.ProductionDataUrl, this.LaunchFileAsync, SystemSections.StableTreeName);

            BoardData stable = SubmissionRowsBoard.ToBoard(table.Rows);

            // Never editable: the server says so, and the table is told so whatever it says.
            editor.Clear();
            editor.IsReadOnly = true;
            editor.Open(
                BoardTableDocument.Create(stable, SubmissionRowsBoard.ToBoard(table.Rows), SystemSections.StableBaselineLabel),
                null,
                preferredSheet: this.thisStableSheetBySystem.TryGetValue(table.SystemId, out string? sheet) ? sheet : null);

            // Why it cannot be changed - under the switch, where BETA's table says what it is.
            WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), table.MayNotEditReason, isError: false);
        }

        // The line above the stable table as shown, or empty - for tests.
        internal string StableTableNoteForTests =>
            this.FindControl<TextBlock>("StableTableMessageText") is { IsVisible: true } line ? TabMaintainer.TextOf(line) : string.Empty;

        // The stable table opened on `table` without asking the server - for tests.
        internal void OpenStableTableForTests(SystemTableAnswer table, SystemOverviewEntry? readAt = null)
        {
            this.OpenStableTable(table);
            this.thisStableTableReadAt = readAt;
        }

        private void CloseStableTable()
        {
            if (this.thisStableTable is not null && this.StableBoardTable.CurrentSheet?.Name is string sheet)
                this.thisStableSheetBySystem[this.thisStableTable.SystemId] = sheet;

            this.thisStableTable = null;
            this.thisStableTableReadAt = null;

            BoardTableEditor editor = this.StableBoardTable;
            editor.Clear();
            editor.FileSource = null;

            WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), null, isError: false);
        }

        // ###########################################################################################
        // Reads the stable source's files for the system on screen and shows them, opened from its
        // public address. The caller holds the wait.
        // ###########################################################################################
        private async Task LoadStableFilesAsync()
        {
            if (this.thisClient is not ReviewApiClient client ||
                this.thisSession is not ReviewSession session ||
                this.ShownSystem is not { } system)
            {
                return;
            }

            ReviewApiResult<SystemFilesAnswer> answer = await client.GetSystemFilesAsync(session, system.SystemId, DataTreeNames.Production);

            if (!string.Equals(this.ShownSystem?.SystemId, system.SystemId, StringComparison.Ordinal))
                return;

            if (!answer.IsOk)
            {
                this.ShowStableFiles(null, null, system.SystemId);
                this.StableFiles.ShowMessage(answer.Message, isError: answer.Failure != ReviewApiFailure.NotFound);
                return;
            }

            // An older server answered with BETA's files: never drawn as the stable source's.
            if (!SystemSections.IsStableAnswer(answer.Value!))
            {
                this.ShowStableFiles(null, null, system.SystemId);
                this.StableFiles.ShowMessage(SystemSections.StableNeedsNewerServer, isError: true);
                return;
            }

            this.ShowStableFiles(answer.Value!, client, system.SystemId);
            this.thisStableFilesFor = system.SystemId;
            this.thisStableFilesReadAt = system;
        }

        private void ShowStableFiles(SystemFilesAnswer? files, ReviewApiClient? client, string systemId)
        {
            FileTreeView tree = this.StableFiles;

            tree.Files = files is null || client is null
                ? null
                : new FileTreeFiles(this, client, this.thisSession, submissionId: null, betaDataUrl: null, files.ProductionDataUrl);

            tree.ShowListing(files?.Files, openFolder: systemId);

            if (this.FindControl<TextBlock>("StableFilesHeadingText") is TextBlock heading)
                heading.Text = files is null ? string.Empty : SystemSections.StableFilesHeading;
        }

        // Shows a stable list without a server - for tests.
        internal void ShowStableFilesForTests(SystemFilesAnswer files, SystemOverviewEntry? readAt = null)
        {
            this.ShowStableFiles(files, client: null, files.SystemId);
            this.thisStableFilesFor = files.SystemId;
            this.thisStableFilesReadAt = readAt;
        }

        private void ClearStableFiles()
        {
            this.thisStableFilesFor = null;
            this.thisStableFilesReadAt = null;

            FileTreeView tree = this.StableFiles;
            tree.Files = null;
            tree.ShowListing(null, openFolder: null);

            if (this.FindControl<TextBlock>("StableFilesHeadingText") is TextBlock heading)
                heading.Text = string.Empty;
        }

        // ###########################################################################################
        // The stable source moved under what is held (a promotion): read again when on screen,
        // otherwise dropped, to be read when next shown.
        // ###########################################################################################
        private async Task CatchUpWithStableAsync(SystemOverviewEntry now)
        {
            bool onScreen = this.ShowsStable;

            if (this.HoldsStableTableFor(now.SystemId) &&
                this.thisStableTableReadAt is { } tableReadAt &&
                QueueRefreshRules.StableBoardChanged(tableReadAt, now))
            {
                if (onScreen && this.thisSection == SystemSection.BoardData)
                    await this.LoadStableTableAsync();
                else
                    this.CloseStableTable();
            }

            if (this.HoldsStableFilesFor(now.SystemId) &&
                this.thisStableFilesReadAt is { } filesReadAt &&
                QueueRefreshRules.StableBoardChanged(filesReadAt, now))
            {
                if (onScreen && this.thisSection == SystemSection.Files)
                    await this.LoadStableFilesAsync();
                else
                    this.ClearStableFiles();
            }
        }

        // Nothing of the stable source's board or files - another system chosen, or signed out.
        private void ForgetStableContent()
        {
            this.CloseStableTable();
            this.ClearStableFiles();
        }
    }
}
