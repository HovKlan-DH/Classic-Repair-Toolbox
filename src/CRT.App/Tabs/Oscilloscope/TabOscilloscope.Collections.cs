using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.Collections;

namespace CRT
{
    // ###########################################################################################
    // WHAT THE "Collections" TAB NEEDS FROM CRT (Main.Collections.cs, HandCollectionsToOscilloscopeTab): the collections on
    // this computer - asked for each time, as a test sets its own store after Main is made - and the collection's window:
    // opened, or opened asking to delete it. A test hands its own.
    // ###########################################################################################
    internal sealed record OscilloscopeCollectionsHost(
        Func<CollectionStore> Store,
        Action<string> OpenCollection,
        Action<string> AskToDeleteCollection,
        Action<string> OpenLink);

    // ###########################################################################################
    // One collection as the list shows it - its words (CollectionListing), and its first capture's thumbnail, set IN PLACE as
    // it is decoded, so the row is not made again under a press.
    // ###########################################################################################
    public sealed class CollectionListRow : INotifyPropertyChanged
    {
        private Bitmap? thisThumbnail;

        public CollectionListRow(CollectionListItem item) => this.Item = item;

        public event PropertyChangedEventHandler? PropertyChanged;

        public CollectionListItem Item { get; }

        // "Open shared page" - while it is shared (owner request, 2026-11-28).
        public bool HasSharedPage => this.Item.IsShared;

        public bool HasNoCapture => this.Item.FirstCaptureSha256 is null;

        public Bitmap? Thumbnail
        {
            get => this.thisThumbnail;
            set
            {
                if (ReferenceEquals(this.thisThumbnail, value))
                    return;

                this.thisThumbnail = value;
                this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.Thumbnail)));
            }
        }
    }

    // ###########################################################################################
    // THE OUTPUT COLUMN'S TWO TABS, AND "Collections" (owner request, 2026-11-27: "In "Oscilloscope" tab - where we now have the
    // "debug window", or what to call it, can we then have two tabs - one for "Debug" (or name this better if you can) and one
    // for "Collections". In the new "Collections" tab we will show all collections and have a management interface for them").
    //
    //   - "Communication log" is the black box it always was - what CRT and the oscilloscope said to each other - and the tab
    //     shown first.
    //   - "Collections" lists every collection on this computer, of every board, the newest made first (CollectionListing): a
    //     card each with its first capture, its board, how many captures and where it stands. "Open collection" ("Show
    //     collection" until 2026-11-28 - owner request) or its picture clicked opens its window, where it is named and shared;
    //     "Open shared page" while it is shared (offline, or never shared, none - owner request, 2026-11-28); "Delete
    //     collection" opens its window ASKING - the window's delete is the one that takes the page down too
    //     (CollectionWindow.DeleteAsync), so there is no second copy of it here. No oscilloscope needs to be connected for any of
    //     it. (A badge counted collections with new comments until 2026-11-28, when the comments went - owner decision.)
    //   - The list is read again whenever Main says the collections changed (Main.ShowCollectionsAsTheyAre), as the tab is
    //     shown, and as "Collections" is picked.
    // ###########################################################################################
    public partial class TabOscilloscope
    {
        private OscilloscopeCollectionsHost? thisCollectionsHost;

        // The list's rows - ONE collection, changed in place (InPlaceList), a row kept while its words are the same.
        private readonly ObservableCollection<CollectionListRow> thisCollectionRows = [];

        // The thumbnails decoded, by the capture's checksum (the same bytes whichever collection holds them) - decoded small off
        // the UI thread once each, and never disposed by hand (a picture still drawn takes the render thread down; CLAUDE.md).
        private readonly Dictionary<string, Bitmap> thisCollectionThumbnails = new(StringComparer.Ordinal);

        private readonly HashSet<string> thisCollectionThumbnailsDecoding = new(StringComparer.Ordinal);

        internal const int CollectionThumbnailDecodeWidth = 320;

        private void InitializeCollectionsTab()
        {
            this.CollectionsList.ItemsSource = this.thisCollectionRows;

            // Shown again (the "Oscilloscope" tab picked): the list as it is now.
            this.AttachedToVisualTree += (_, _) => this.ShowCollections();
        }

        internal void UseCollections(OscilloscopeCollectionsHost host)
        {
            this.thisCollectionsHost = host;

            if (TopLevel.GetTopLevel(this) is not null)
                this.ShowCollections();
        }

        internal bool IsCollectionsTabShown => this.CollectionsPanel.IsVisible;

        // ###########################################################################################
        // The oscilloscope a capture is taken with, for its collection (owner request, 2026-11-28: "the oscilloscope brand and
        // model should be shown there") - the one picked here while a session is open on it (picking another ends the session),
        // null without one.
        // ###########################################################################################
        internal string? ConnectedOscilloscopeName =>
            this.HasActiveEstablishedOscilloscopeSessionForTitleState() && this.GetSelectedOscilloscope() is { } oscilloscope
                ? CollectionCaptures.OscilloscopeNameOf(oscilloscope.Brand, oscilloscope.SeriesOrModel)
                : null;

        // ###########################################################################################
        // "Communication log" or "Collections".
        // ###########################################################################################
        private void OnOutputTabClick(object? sender, RoutedEventArgs e) =>
            this.ShowOutputTab(collections: ReferenceEquals(sender, this.CollectionsTabButton));

        internal void ShowOutputTab(bool collections)
        {
            this.CollectionsPanel.IsVisible = collections;
            this.OutputTextBox.IsVisible = !collections;
            this.CollectionsTabButton.Classes.Set("Selected", collections);
            this.CommunicationLogTabButton.Classes.Set("Selected", !collections);

            if (collections)
                this.ShowCollections();
        }

        // ###########################################################################################
        // The collections as they are now - the rows and the words over them.
        // ###########################################################################################
        internal void ShowCollections()
        {
            if (this.thisCollectionsHost is not { } host)
                return;

            IReadOnlyList<CollectionListItem> items = CollectionListing.ItemsOf(host.Store().All());

            InPlaceList.Sync(
                this.thisCollectionRows,
                items,
                (row, item) => row.Item == item,
                item => new CollectionListRow(item)
                {
                    Thumbnail = item.FirstCaptureSha256 is { } sha256 ? this.thisCollectionThumbnails.GetValueOrDefault(sha256) : null
                },
                itemKey: row => row.Item.LocalId,
                wantedKey: item => item.LocalId);

            this.CollectionsHeadingText.Text = items.Count == 0 ? CollectionListing.NoneYet : CollectionListing.HeadingOf(items.ToList());
            this.CollectionsIntroText.Text = CollectionListing.Intro;
            this.CollectionsIntroText.IsVisible = items.Count > 0;

            // Pictures only for a list somebody looks at.
            if (this.IsCollectionsTabShown)
                this.LoadCollectionThumbnails(host.Store());
        }

        private void LoadCollectionThumbnails(CollectionStore store)
        {
            foreach (CollectionListRow row in this.thisCollectionRows.Where(row => row.Thumbnail is null && row.Item.FirstCaptureSha256 is not null))
            {
                string sha256 = row.Item.FirstCaptureSha256!;

                if (this.thisCollectionThumbnails.TryGetValue(sha256, out Bitmap? decoded))
                {
                    row.Thumbnail = decoded;
                    continue;
                }

                if (!this.thisCollectionThumbnailsDecoding.Add(sha256))
                    continue;

                string path = store.CapturePath(row.Item.LocalId, sha256);

                // Never called off: a thumbnail decoded for a list no longer shown is kept for the next time it is.
                _ = Task.Run(
                    () =>
                    {
                        try
                        {
                            // Opened so it can be deleted while read - a collection deleted as its thumbnail decodes (the
                            // collection window's own rule).
                            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                            return Bitmap.DecodeToWidth(stream, TabOscilloscope.CollectionThumbnailDecodeWidth);
                        }
                        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException)
                        {
                            return null;
                        }
                    },
                    System.Threading.CancellationToken.None).ContinueWith(
                    task =>
                    {
                        this.thisCollectionThumbnailsDecoding.Remove(sha256);

                        if (!task.IsCompletedSuccessfully || task.Result is not { } bitmap)
                            return;

                        this.thisCollectionThumbnails[sha256] = bitmap;

                        // The rows shown now with that capture first - the list may have changed meanwhile.
                        foreach (CollectionListRow shown in this.thisCollectionRows.Where(shown => shown.Item.FirstCaptureSha256 == sha256))
                            shown.Thumbnail = bitmap;
                    },
                    TaskScheduler.FromCurrentSynchronizationContext());
            }
        }

        // ###########################################################################################
        // A card's buttons, and its picture.
        // ###########################################################################################
        private void OnShowCollectionClick(object? sender, RoutedEventArgs e)
        {
            if (TabOscilloscope.RowOf(sender) is { } row)
                this.thisCollectionsHost?.OpenCollection(row.Item.LocalId);
        }

        private void OnCollectionThumbnailTapped(object? sender, TappedEventArgs e)
        {
            if (TabOscilloscope.RowOf(sender) is { } row)
                this.thisCollectionsHost?.OpenCollection(row.Item.LocalId);
        }

        private void OnOpenCollectionPageClick(object? sender, RoutedEventArgs e)
        {
            if (TabOscilloscope.RowOf(sender)?.Item is { IsShared: true, PageUrl: { } url })
                this.thisCollectionsHost?.OpenLink(url);
        }

        private void OnDeleteCollectionClick(object? sender, RoutedEventArgs e)
        {
            if (TabOscilloscope.RowOf(sender) is { } row)
                this.thisCollectionsHost?.AskToDeleteCollection(row.Item.LocalId);
        }

        private static CollectionListRow? RowOf(object? sender) => (sender as Control)?.DataContext as CollectionListRow;

        internal IReadOnlyList<CollectionListRow> CollectionRowsForTests => this.thisCollectionRows;
    }
}
