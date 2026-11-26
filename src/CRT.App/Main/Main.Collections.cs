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
    // settings read back through the "Oscilloscope" tab) and opens a collection's window (OpenCollectionWindow) with the account
    // CRT is signed in to - the "Maintainer" tab's session, as everything signed in in CRT - and the contribution server, the
    // test instance's for a pre-release (AppConfig.ContributionServerRootUrl).
    //
    // A collection's window open stays one: opened again, it comes forward. A change in it refreshes every popup's drop-down.
    // ###########################################################################################
    public partial class Main
    {
        private CollectionStore? thisCollections;

        // The collections on this computer - a test's folder when it set one.
        internal CollectionStore Collections => this.thisCollections ??= new CollectionStore(AppConfig.CollectionsFolder);

        internal void UseCollectionsForTests(CollectionStore store) => this.thisCollections = store;

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

            this.AskForNewCommentsSoon();

            return new CollectionContext(
                this.Collections,
                this.GetCurrentBoardKey(),
                board.HardwareName,
                board.BoardName,
                regions.Count > 0 ? this._localRegion : null,
                this.ReadScopeSettingsForCollectionAsync,
                this.OpenCollectionWindow,
                serverId => this.thisLatestCommentIds.GetValueOrDefault(serverId));
        }

        // ###########################################################################################
        // NEW COMMENTS (step 3): the server's list of the account's collections says the newest comment each page shows; a
        // collection whose newest is newer than its owner read is marked in the popups' drop-downs ("new comments"). Asked
        // as a popup is handed its collections - at most every CommentsAskedEvery, only signed in and with one shared - never
        // on a timer of its own. The owner is mailed for every comment anyway (CollectionCommentFlows).
        // ###########################################################################################
        internal static readonly TimeSpan CommentsAskedEvery = TimeSpan.FromMinutes(5);

        private readonly Dictionary<string, long> thisLatestCommentIds = new(StringComparer.Ordinal);

        private DateTimeOffset thisCommentsAskedUtc = DateTimeOffset.MinValue;

        private void AskForNewCommentsSoon()
        {
            if (this.TabMaintainer.SignedIn is not { } session ||
                DateTimeOffset.UtcNow - this.thisCommentsAskedUtc < Main.CommentsAskedEvery ||
                !this.Collections.All().Any(collection => collection.IsShared))
            {
                return;
            }

            this.thisCommentsAskedUtc = DateTimeOffset.UtcNow;

            _ = HandlerGuard.RunAsync(
                async () =>
                {
                    IReadOnlyList<CollectionSummaryAnswer>? mine = await this.CollectionClientInUse().ListMineAsync(session.BearerToken);

                    if (mine is null)
                        return;

                    foreach (CollectionSummaryAnswer shared in mine)
                        this.thisLatestCommentIds[shared.Id] = shared.LatestCommentId;

                    this.ShowCollectionChoicesInPopups();
                },
                _ => { });
        }

        private CollectionClient CollectionClientInUse() =>
            this.CollectionClientOverrideForTests ?? new CollectionClient(AppConfig.ContributionServerRootUrl);

        private async Task<ScopeSettingsReading?> ReadScopeSettingsForCollectionAsync() =>
            await ((IBaselineScope)this.TabOscilloscopeControl).ReadSettingsAsync(CancellationToken.None);

        // ###########################################################################################
        // A collection's window, over CRT's - the one open for it brought forward.
        // ###########################################################################################
        internal void OpenCollectionWindow(string localId)
        {
            if (this.thisCollectionWindows.TryGetValue(localId, out CollectionWindow? open) && open.IsVisible)
            {
                open.Activate();
                return;
            }

            var window = new CollectionWindow();
            window.Initialize(this.CollectionWindowHost(), localId);
            window.Closed += (_, _) =>
            {
                if (this.thisCollectionWindows.TryGetValue(localId, out CollectionWindow? kept) && ReferenceEquals(kept, window))
                    this.thisCollectionWindows.Remove(localId);
            };

            this.thisCollectionWindows[localId] = window;
            window.Show(this);
        }

        internal CollectionWindowHost CollectionWindowHost() =>
            new(
                this.Collections,
                this.CollectionClientInUse(),
                SignedIn: () => this.TabMaintainer.SignedIn,
                SignIn: this.OpenAccountWindowAsync,
                OpenLink: link => ExternalTargetLauncher.OpenOrSay(link),
                CopyText: async text =>
                {
                    if (TopLevel.GetTopLevel(this)?.Clipboard is { } clipboard)
                        await clipboard.SetTextAsync(text);
                },
                Changed: this.ShowCollectionChoicesInPopups);

        // Every component info window open shows the board's collections as they now are.
        private void ShowCollectionChoicesInPopups()
        {
            this._singleComponentInfoWindow?.ShowCollectionChoices();

            foreach (ComponentInfoWindow popup in this._componentInfoWindowsByKey.Values.ToList())
                popup.ShowCollectionChoices();
        }
    }
}
