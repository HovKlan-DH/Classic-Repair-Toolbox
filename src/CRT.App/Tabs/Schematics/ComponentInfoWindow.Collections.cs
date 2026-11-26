using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Avalonia.Interactivity;
using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace CRT
{
    // ###########################################################################################
    // WHAT A POPUP NEEDS TO COLLECT CAPTURES (2026-11-26) - handed over by Main with the component (Main.ComponentPopup.cs): the
    // collections on this computer, the board shown, the region shown when the board has regions, how to read the oscilloscope's
    // settings back, and how to open a collection's window. Null for none: the row is not shown.
    // ###########################################################################################
    // `LatestCommentIdOf`: the newest comment a shared collection's page shows, by its server id (step 3) - 0 when not known.
    internal sealed record CollectionContext(
        CollectionStore Store,
        string BoardKey,
        string HardwareName,
        string BoardName,
        string? Region,
        Func<Task<ScopeSettingsReading?>> ReadScopeSettings,
        Action<string> OpenCollection,
        Func<string, long>? LatestCommentIdOf = null);

    // ###########################################################################################
    // "COLLECT CAPTURES INTO" (owner request, 2026-11-25: "the UI for this should ... be visible from the component info window,
    // as the user maybe are diagnosic a specific component and want some help with that. And maybe he then later want to add
    // another component to the same collection"). A row above the footer, shown once an oscilloscope session has been seen:
    //
    //   - a drop-down of the board's collections on this computer, "None" first and "New collection" last - the pick kept
    //     PER BOARD (UserSettings.CollectionOfBoard), so another component's popup, or the same one another day, fills the same;
    //   - every capture taken while one is picked goes into it too - the PNG, its pin, the oscilloscope's settings read back,
    //     and the baseline shown (CollectionCaptures) - the banner saying so, with "Take out of the collection";
    //   - "Show collection" opens it (CollectionWindow), where it is named, described and shared.
    //
    // Nothing is sent anywhere from here: a collection stays on this computer until "Share" in its window.
    // ###########################################################################################
    public partial class ComponentInfoWindow
    {
        private CollectionContext? thisCollections;

        // True while the drop-down is filled from code - no pick of anybody's.
        private bool thisShowingCollectionChoices;

        // The capture the banner says went into a collection - what "Take out of the collection" takes out.
        private (string LocalId, string Sha256)? thisCollected;

        internal sealed record CollectionChoice(string? LocalId, bool IsNew, string Text)
        {
            public override string ToString() => this.Text;
        }

        public const string NoCollectionChoice = "None";

        public const string NewCollectionChoice = "New collection";

        internal void UseCollections(CollectionContext? context)
        {
            this.thisCollections = context;
            this.ShowCollectionChoices();
        }

        private bool OffersCollections =>
            this.thisCollections is not null &&
            UserSettings.EnableNetworkConnectedOscilloscopeTab &&
            this._hasSeenOscilloscopeSessionTitleState;

        // The collection picked for the board, while it is still on this computer.
        private LocalCollection? PickedCollection =>
            this.thisCollections is { } context && UserSettings.CollectionOfBoard(context.BoardKey) is { } localId
                ? context.Store.Find(localId)
                : null;

        // ###########################################################################################
        // The row, and its drop-down as the board's collections are now.
        // ###########################################################################################
        internal void ShowCollectionChoices()
        {
            this.CollectionRow.IsVisible = this.OffersCollections;

            if (this.thisCollections is not { } context || !this.OffersCollections)
                return;

            IReadOnlyList<LocalCollection> collections = context.Store.OfBoard(context.HardwareName, context.BoardName);
            string? picked = UserSettings.CollectionOfBoard(context.BoardKey);

            if (picked is not null && !collections.Any(collection => collection.LocalId == picked))
                picked = null;

            var choices = new List<CollectionChoice> { new(null, false, ComponentInfoWindow.NoCollectionChoice) };

            choices.AddRange(collections.Select(collection => new CollectionChoice(
                collection.LocalId,
                false,
                CollectionSharing.LineOf(collection, collection.ServerId is { } serverId ? context.LatestCommentIdOf?.Invoke(serverId) ?? 0 : 0))));
            choices.Add(new CollectionChoice(null, true, ComponentInfoWindow.NewCollectionChoice));

            this.thisShowingCollectionChoices = true;

            try
            {
                this.CollectionComboBox.ItemsSource = choices;
                this.CollectionComboBox.SelectedItem = choices.FirstOrDefault(choice => !choice.IsNew && choice.LocalId == picked) ?? choices[0];
            }
            finally
            {
                this.thisShowingCollectionChoices = false;
            }

            this.ShowCollectionButton.IsEnabled = picked is not null;
        }

        private void OnCollectionPicked(object? sender, Avalonia.Controls.SelectionChangedEventArgs e)
        {
            if (this.thisShowingCollectionChoices || this.thisCollections is not { } context ||
                this.CollectionComboBox.SelectedItem is not CollectionChoice choice)
            {
                return;
            }

            string? localId = choice.LocalId;

            if (choice.IsNew)
            {
                localId = context.Store.Create(
                    context.HardwareName,
                    context.BoardName,
                    context.Region,
                    CollectionSharing.DefaultTitle(this._boardLabel, DateTimeOffset.Now),
                    DateTimeOffset.UtcNow)?.LocalId;
            }

            UserSettings.SetCollectionOfBoard(context.BoardKey, localId);
            this.ShowCollectionChoices();
        }

        private void OnShowCollectionClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisCollections is { } context && this.PickedCollection is { } collection)
                context.OpenCollection(collection.LocalId);
        }

        // ###########################################################################################
        // A capture just saved, into the collection picked too: its bytes copied in, its pin and the settings read back from the
        // oscilloscope with it. Nothing when none is picked. `baselinePath`: the file of the image shown, whose checksum names
        // the baseline in the published data.
        // ###########################################################################################
        internal async Task CollectCaptureAsync(ComponentImageEntry shown, string? baselinePath, string savedFilePath)
        {
            if (!this.OffersCollections || this.thisCollections is not { } context || this.PickedCollection is not { } collection)
                return;

            try
            {
                ScopeSettingsReading? reading = null;

                try
                {
                    reading = await context.ReadScopeSettings();
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Could not read the oscilloscope's settings for a capture into a collection: {ex.Message}");
                }

                (byte[] png, string? baselineSha) = await Task.Run(() => (
                    File.ReadAllBytes(savedFilePath),
                    baselinePath is { } path && File.Exists(path) ? DataTreeLayout.Sha256Of(path) : null));

                LocalCapture capture = CollectionCaptures.CaptureOf(shown, reading, baselineSha, DateTimeOffset.UtcNow);
                LocalCollection? changed = await Task.Run(() => context.Store.AddCapture(collection.LocalId, png, capture));

                this.CollectedText.IsVisible = true;

                if (changed is null)
                {
                    this.CollectedText.Text = $"Could not add it to [{CollectionSharing.NameOf(collection)}]";
                    return;
                }

                this.thisCollected = (collection.LocalId, Convert.ToHexStringLower(SHA256.HashData(png)));
                this.CollectedText.Text = $"Added to [{CollectionSharing.NameOf(changed)}]";
                this.TakeOutOfCollectionButton.IsVisible = true;
                this.ShowCollectionChoices();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not add a capture to the oscilloscope collection [{CollectionSharing.NameOf(collection)}]: {ex.Message}");
            }
        }

        private void OnTakeOutOfCollectionClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisCollections is not { } context || this.thisCollected is not { } collected)
                return;

            LocalCollection? changed = context.Store.RemoveCapture(collected.LocalId, collected.Sha256);

            this.thisCollected = null;
            this.TakeOutOfCollectionButton.IsVisible = false;
            this.CollectedText.Text = changed is null ? "Could not take it out of the collection" : $"Taken out of [{CollectionSharing.NameOf(changed)}]";
            this.ShowCollectionChoices();
        }

        // The banner gone with its capture: what it said of the collection goes too.
        private void ForgetCollectedCapture()
        {
            this.thisCollected = null;
            this.CollectedText.IsVisible = false;
            this.TakeOutOfCollectionButton.IsVisible = false;
        }

        internal string? CollectedTextForTests => this.CollectedText.IsVisible ? this.CollectedText.Text : null;
    }
}
