using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace CRT
{
    // ###########################################################################################
    // OSCILLOSCOPE COLLECTIONS, WIRED (2026-11-26; CRT.Data's CollectionRules says what one is). Main hands each component info
    // window what it needs to collect captures (CollectionContextForPopup - the board shown, the region shown, the oscilloscope's
    // settings read back through the "Oscilloscope" tab) and opens a collection's window (OpenCollectionWindow) with the
    // contribution server, the test instance's for a pre-release (AppConfig.ContributionServerRootUrl). No account: a shared
    // collection's page is owned by the key kept with it on this computer (owner decision, 2026-11-28; LocalCollection.OwnerKey).
    //
    // A collection's window open stays one: opened again, it comes forward. A change in it refreshes every popup's drop-down,
    // and the "Oscilloscope" tab's "Collections" list (TabOscilloscope.Collections.cs - every collection, opened and deleted
    // with no oscilloscope connected; owner request, 2026-11-27).
    // ###########################################################################################
    public partial class Main
    {
        private CollectionStore? thisCollections;

        // The collections on this computer - a test's folder when it set one.
        internal CollectionStore Collections => this.thisCollections ??= new CollectionStore(AppConfig.CollectionsFolder);

        // A 4.1.0 pre-release takes the folder it kept its collections in until it had one of its own
        // (AppConfig.CollectionsFolderToTakeOver). Run by StartAsync alone - a test never moves the user's folder.
        private static void TakeOverEarlierCollections()
        {
            if (AppConfig.CollectionsFolderToTakeOver is { } earlier)
                CollectionStore.TakeOverEarlierFolder(earlier, AppConfig.CollectionsFolder);
        }

        // For a test: the server a collection's window shares with.
        internal CollectionClient? CollectionClientOverrideForTests { get; set; }

        private readonly Dictionary<string, CollectionWindow> thisCollectionWindows = new(StringComparer.Ordinal);

        // ###########################################################################################
        // What a popup on the board shown collects into - none with no board shown. `regions`: the board's, so a board without
        // them names none.
        // ###########################################################################################
        private CollectionContext? CollectionContextForPopup(IReadOnlyList<string> regions)
        {
            if (this.GetCurrentBoardEntry() is not { } board)
                return null;

            return new CollectionContext(
                this.Collections,
                this.GetCurrentBoardKey(),
                board.HardwareName,
                board.BoardName,
                regions.Count > 0 ? this._localRegion : null,
                this.ReadScopeSettingsForCollectionAsync,
                this.OpenCollectionWindow,
                this.CollectionChangedInPopup,
                localId => this.thisCollectionWindows.TryGetValue(localId, out CollectionWindow? open) && open.IsVisible && open.IsBusy,
                () => this.TabOscilloscopeControl.ConnectedOscilloscopeName);
        }

        // A capture added or taken out in a popup, or another collection picked there: every popup's drop-down, and the
        // collection's window if one is open.
        private void CollectionChangedInPopup(string? localId)
        {
            this.ShowCollectionsAsTheyAre();

            if (localId is not null && this.thisCollectionWindows.TryGetValue(localId, out CollectionWindow? open) && open.IsVisible)
                open.ShowAgain();
        }

        // ONE client for the run (code review, 2026-11-27: one made at every ask left an HttpClient and its connections behind).
        private CollectionClient? thisCollectionClient;

        private CollectionClient CollectionClientInUse() =>
            this.CollectionClientOverrideForTests ?? (this.thisCollectionClient ??= new CollectionClient(AppConfig.ContributionServerRootUrl));

        private async Task<ScopeSettingsReading?> ReadScopeSettingsForCollectionAsync() =>
            await ((IBaselineScope)this.TabOscilloscopeControl).ReadSettingsAsync(CancellationToken.None);

        // ###########################################################################################
        // A collection's window, over CRT's - the one open for it brought forward.
        // ###########################################################################################
        internal void OpenCollectionWindow(string localId) => this.ShowCollectionWindow(localId);

        // "Delete collection" on the "Collections" list: its window, ASKING - the window's delete takes the page down too, so the
        // list has no second copy of it.
        internal void OpenCollectionWindowToDelete(string localId) => this.ShowCollectionWindow(localId)?.AskToDelete();

        // The collection's window shown - the one open brought forward - or null when there is no such collection any more.
        private CollectionWindow? ShowCollectionWindow(string localId)
        {
            if (this.thisCollectionWindows.TryGetValue(localId, out CollectionWindow? open) && open.IsVisible)
            {
                open.Activate();
                return open;
            }

            var window = new CollectionWindow();

            // Gone meanwhile - deleted from its window, say: no window, and every popup and the list show the collections as
            // they are.
            if (!window.Initialize(this.CollectionWindowHost(), localId))
            {
                this.ShowCollectionsAsTheyAre();
                return null;
            }

            window.Closed += (_, _) =>
            {
                if (this.thisCollectionWindows.TryGetValue(localId, out CollectionWindow? kept) && ReferenceEquals(kept, window))
                    this.thisCollectionWindows.Remove(localId);
            };

            this.thisCollectionWindows[localId] = window;
            window.Show(this);
            return window;
        }

        // ###########################################################################################
        // The "Oscilloscope" tab's "Collections" list (owner request, 2026-11-27): the store asked for each time - a test sets its
        // own after Main is made - and every action the collection's window's.
        // ###########################################################################################
        private void HandCollectionsToOscilloscopeTab() =>
            this.TabOscilloscopeControl.UseCollections(new OscilloscopeCollectionsHost(
                () => this.Collections,
                this.OpenCollectionWindow,
                this.OpenCollectionWindowToDelete,
                link => ExternalTargetLauncher.OpenOrSay(link)));

        internal CollectionWindowHost CollectionWindowHost() =>
            new(
                this.Collections,
                this.CollectionClientInUse(),
                OpenLink: link => ExternalTargetLauncher.OpenOrSay(link),
                CopyText: async text =>
                {
                    if (TopLevel.GetTopLevel(this)?.Clipboard is { } clipboard)
                        await clipboard.SetTextAsync(text);
                },
                Changed: this.ShowCollectionsAsTheyAre,
                Place: () => UserSettings.CollectionWindowPlace,
                RememberPlace: UserSettings.SaveCollectionWindowPlace,
                LinkFormat: () => UserSettings.CollectionLinkFormat,
                RememberLinkFormat: UserSettings.SetCollectionLinkFormat,
                WorklogIsOn: () => UserSettings.EnableWorklog,
                AttachToWorklog: this.AttachCollectionToWorklogAsync);

        // ###########################################################################################
        // "ATTACH TO WORKLOG" ON A COLLECTION'S WINDOW (owner request, 2026-11-27; asked, "A copy of its captures" -
        // CollectionWorklogAttach): into a worklog of the workbook worked in on the COLLECTION'S board, picked in the oscilloscope
        // capture's own dialog - or, while that board is the one shown, a new worklog filed against its schematic image shown,
        // as the capture's "Create new worklog" files one (ComponentInfoWindow.CreateWorklogEntryForCaptureAsync). What came of
        // it in words for the window, or null when nothing was done.
        // ###########################################################################################
        private async Task<string?> AttachCollectionToWorklogAsync(LocalCollection collection, Window owner)
        {
            string boardKey = $"{collection.HardwareName}|{collection.BoardName}";

            if (Main.ResolveActiveWorkbookForBoard(boardKey) is not { } workbook)
                return CollectionWorklogAttach.NoWorkbook(collection);

            string CapturePath(string sha256) => this.Collections.CapturePath(collection.LocalId, sha256);

            IReadOnlyList<(LocalCapture Capture, string Path)> captures = CollectionWorklogAttach.CapturesOf(collection, CapturePath);

            if (captures.Count == 0)
                return CollectionWorklogAttach.NothingToAttach;

            bool BoardShown() => string.Equals(this.GetCurrentBoardKey(), boardKey, StringComparison.Ordinal);

            List<WorklogEntryRecord> entries = WorklogManager.GetEntries(workbook.Id);

            if (entries.Count == 0 && !BoardShown())
                return CollectionWorklogAttach.NoWorklogYet(collection);

            // The component the captures are of, when they are all of one - its worklogs are listed first.
            List<string> components = captures.Select(pair => pair.Capture.ComponentLabel).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            var dialog = new WorklogAttachCaptureWindow();
            dialog.InitializeForCollection(
                captures[0].Path,
                CollectionWorklogAttach.Summary(collection, captures.Count),
                workbook,
                entries,
                components.Count == 1 ? components[0] : null,
                offerNewEntry: BoardShown());

            if (await dialog.ShowDialog<WorklogAttachCaptureWindow.AttachResult?>(owner) is not { } result)
                return null;

            if (result.EntryId is { } entryId)
            {
                CollectionWorklogAttach.Outcome outcome = CollectionWorklogAttach.ToEntry(workbook.Id, entryId, collection, CapturePath);
                this.RefreshWorklogBar();

                return CollectionWorklogAttach.Said(outcome, $"worklog #{entryId.ToString(System.Globalization.CultureInfo.InvariantCulture)}");
            }

            // A new worklog is filed against the schematic image shown - of the collection's board, which may have been changed
            // while the dialog was open.
            if (!BoardShown())
                return CollectionWorklogAttach.BoardNoLongerShown(collection);

            var editor = new WorklogEntryEditorWindow();
            editor.InitializeForNewEntry(
                workbook.Id,
                this.TabSchematicsControl?.GetCurrentSchematicName() ?? string.Empty,
                default,
                this.TabSchematicsControl?.currentFullResBitmap);
            editor.SetShowMarkedAreaForNewEntry(false);

            CollectionWorklogAttach.Outcome added = CollectionWorklogAttach.ToEditor(collection, CapturePath, editor.AttachCapturedPhoto, editor.AttachLink);

            // Nothing could be copied in: no editor that looks as if it all went well - it is not opened at all (the capture's own
            // rule, ComponentInfoWindow.CreateWorklogEntryForCaptureAsync).
            if (added.Added == 0)
                return CollectionWorklogAttach.Said(added, "a new worklog");

            await editor.ShowDialog(owner);
            this.RefreshWorklogBar();

            return editor.WasSaved ? CollectionWorklogAttach.Said(added, "a new worklog") : null;
        }

        // Every component info window open shows the board's collections as they now are, and the "Collections" list every
        // collection.
        private void ShowCollectionsAsTheyAre()
        {
            this._singleComponentInfoWindow?.ShowCollectionChoices();

            foreach (ComponentInfoWindow popup in this._componentInfoWindowsByKey.Values.ToList())
                popup.ShowCollectionChoices();

            this.TabOscilloscopeControl.ShowCollections();
        }
    }
}
