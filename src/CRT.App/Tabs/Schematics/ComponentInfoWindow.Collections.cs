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
        Func<string, long>? LatestCommentIdOf = null,
        // A collection's captures changed in this popup (its local id) - or, null, only the board's pick: every popup shows it.
        Action<string?>? CollectionChanged = null,
        Func<string, bool>? IsBeingShared = null);

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

        // The choices asked for again while the list was open - shown once it closes.
        private bool thisCollectionChoicesWaiting;

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

            if (!this.thisCollectionPickerWired)
            {
                this.thisCollectionPickerWired = true;

                // *** A PICK IS MADE FROM THE OPEN LIST (code review, 2026-11-27), as the worklog bar's device picker's. *** Closed
                // and focused, the arrow keys and the wheel stepped the selection - and every pass over "New collection" made an
                // empty one, picked for the board. Alt+Down, Enter and F4 still open the list.
                this.CollectionComboBox.AddHandler(KeyDownEvent, this.OnCollectionPickerKeyDown, RoutingStrategies.Tunnel);
                this.CollectionComboBox.AddHandler(PointerWheelChangedEvent, this.OnCollectionPickerWheel, RoutingStrategies.Tunnel);
            }

            this.ShowCollectionChoices();
        }

        private bool thisCollectionPickerWired;

        private void OnCollectionPickerKeyDown(object? sender, Avalonia.Input.KeyEventArgs e)
        {
            if (!this.CollectionComboBox.IsDropDownOpen &&
                !e.KeyModifiers.HasFlag(Avalonia.Input.KeyModifiers.Alt) &&
                e.Key is Avalonia.Input.Key.Up or Avalonia.Input.Key.Down or Avalonia.Input.Key.Home or Avalonia.Input.Key.End
                    or Avalonia.Input.Key.PageUp or Avalonia.Input.Key.PageDown)
            {
                e.Handled = true;
            }
        }

        private void OnCollectionPickerWheel(object? sender, Avalonia.Input.PointerWheelEventArgs e)
        {
            if (!this.CollectionComboBox.IsDropDownOpen)
                e.Handled = true;
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

            CollectionChoice pick = choices.FirstOrDefault(choice => !choice.IsNew && choice.LocalId == picked) ?? choices[0];

            this.ShowCollectionButton.IsEnabled = picked is not null;

            // *** NEVER UNDER AN OPEN LIST, AND NOTHING DONE FOR CHOICES AS SHOWN (code review, 2026-11-27). *** This runs at every
            // push of the oscilloscope's state and as the comment check answers: the list open, it was replaced under the pointer
            // and the item clicked was another. It waits for the list to close (OnCollectionDropDownClosed).
            if (this.CollectionComboBox.IsDropDownOpen)
            {
                this.thisCollectionChoicesWaiting = true;
                return;
            }

            this.thisCollectionChoicesWaiting = false;

            if (this.CollectionComboBox.ItemsSource is IReadOnlyList<CollectionChoice> shown && shown.SequenceEqual(choices) &&
                Equals(this.CollectionComboBox.SelectedItem, pick))
            {
                return;
            }

            this.thisShowingCollectionChoices = true;

            try
            {
                this.CollectionComboBox.ItemsSource = choices;
                this.CollectionComboBox.SelectedItem = pick;
            }
            finally
            {
                this.thisShowingCollectionChoices = false;
            }
        }

        private void OnCollectionDropDownClosed(object? sender, EventArgs e)
        {
            if (this.thisCollectionChoicesWaiting)
                this.ShowCollectionChoices();
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
                // The region THIS popup shows - its own buttons switch it, the context's is what Main showed as it opened (code
                // review, 2026-11-27: an NTSC popup's new collection read "(PAL)" on its page). None on a board without regions.
                localId = context.Store.Create(
                    context.HardwareName,
                    context.BoardName,
                    context.Region is null ? null : this._localRegion,
                    CollectionSharing.DefaultTitle(this._boardLabel, DateTimeOffset.Now),
                    DateTimeOffset.UtcNow)?.LocalId;

                // Not kept: the pick stays as it was, and it is said (code review, 2026-11-27: the pick went to "None" unsaid).
                if (localId is null)
                {
                    this.ShowCollectionChoices();
                    this.CollectedText.IsVisible = true;
                    this.CollectedText.Text = ComponentInfoWindow.NewCollectionNotKept;
                    return;
                }
            }

            UserSettings.SetCollectionOfBoard(context.BoardKey, localId);
            this.ShowCollectionChoices();

            // The pick is the board's, so every popup of it shows it - another one's captures go there now too (code review,
            // 2026-11-27: the others kept showing the pick before). No collection's content changed, so null: an open collection
            // window is left as it is, a note being typed in it too.
            context.CollectionChanged?.Invoke(null);
        }

        internal const string NewCollectionNotKept =
            "Could not make the new collection on this computer - see \"Errors and warnings\" on the \"Configuration\" tab.";

        private void OnShowCollectionClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisCollections is { } context && this.PickedCollection is { } collection)
                context.OpenCollection(collection.LocalId);
        }

        // ###########################################################################################
        // A capture just saved, into the collection picked too: its bytes copied in, its pin and the settings read back from the
        // oscilloscope with it - and what this window shows of its component AS IT IS NOW (owner request, 2026-11-27: the page
        // shows "something like the component info window" beside a capture seen large): its facts and pins, and a picture of
        // the drawn pinout with the shown image's pin marked. Taken before anything is awaited, so moving on meanwhile changes
        // nothing of it. Nothing when no collection is picked.
        // ###########################################################################################
        internal async Task CollectCaptureAsync(ComponentImageEntry shown, string savedFilePath)
        {
            if (!this.OffersCollections || this.thisCollections is not { } context || this.PickedCollection is not { } collection)
                return;

            CollectionComponent? component = this.ComponentShownForCollection();
            byte[]? pinout = this.PinoutPictureForCollection();

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

                byte[] png = await Task.Run(() => File.ReadAllBytes(savedFilePath));

                LocalCapture capture = CollectionCaptures.CaptureOf(shown, reading, component, DateTimeOffset.UtcNow);
                CaptureAdded added = await Task.Run(() => context.Store.AddCapture(collection.LocalId, png, capture, pinout));

                // *** THE BANNER IS THE CAPTURE ON SCREEN'S (code review, 2026-11-27). *** Moved to another image while the settings
                // were read, it showed "Added to" and "Take out" under that image - and took this capture out. The capture is
                // in the collection all the same; only its window says so then.
                if (!string.Equals(this.thisCapturedScopeImagePath, savedFilePath, StringComparison.Ordinal))
                {
                    if (added.Kind == CaptureAddedKind.Added)
                    {
                        this.ShowCollectionChoices();
                        context.CollectionChanged?.Invoke(collection.LocalId);
                    }

                    return;
                }

                this.CollectedText.IsVisible = true;

                // "Take out" only for a capture THIS click added (code review, 2026-11-27: the same screen twice - two ground
                // pins - took the earlier capture out).
                this.thisCollected = null;
                this.TakeOutOfCollectionButton.IsVisible = false;

                switch (added.Kind)
                {
                    case CaptureAddedKind.Added:
                        this.thisCollected = (collection.LocalId, Convert.ToHexStringLower(SHA256.HashData(png)));
                        this.CollectedText.Text = $"Added to [{CollectionSharing.NameOf(added.Collection!)}]";
                        this.TakeOutOfCollectionButton.IsVisible = true;
                        this.ShowCollectionChoices();

                        // Its window, if open, lists it - and a shared one offers "Share" (code review, 2026-11-27).
                        context.CollectionChanged?.Invoke(collection.LocalId);
                        break;

                    case CaptureAddedKind.AlreadyIn:
                        this.CollectedText.Text = $"The very same screen is in [{CollectionSharing.NameOf(collection)}] already";
                        break;

                    case CaptureAddedKind.Refused:
                        this.CollectedText.Text = $"Not added to [{CollectionSharing.NameOf(collection)}]: {added.Problem}";
                        break;

                    default:
                        this.CollectedText.Text = $"Could not add it to [{CollectionSharing.NameOf(collection)}]";
                        break;
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Said where it happened - a screen held by a virus scanner as it was read is missing from the collection - under its
                // own capture alone, as above; moved on from it, the "Something went wrong" card says it instead.
                bool shownHere = string.Equals(this.thisCapturedScopeImagePath, savedFilePath, StringComparison.Ordinal);

                if (shownHere)
                {
                    this.CollectedText.IsVisible = true;
                    this.CollectedText.Text = $"Could not add it to [{CollectionSharing.NameOf(collection)}]";
                    this.thisCollected = null;
                    this.TakeOutOfCollectionButton.IsVisible = false;
                }

                Logger.Critical($"Could not add a capture to the oscilloscope collection [{CollectionSharing.NameOf(collection)}]: {ex.Message}", alreadyShown: shownHere);
            }
        }

        private void OnTakeOutOfCollectionClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisCollections is not { } context || this.thisCollected is not { } collected)
                return;

            // Its window may be sharing it - the share reads the capture's file as it goes (code review, 2026-11-27).
            if (context.IsBeingShared?.Invoke(collected.LocalId) == true)
            {
                this.CollectedText.Text = CollectionWindow.WaitForTheShare;
                return;
            }

            LocalCollection? changed = context.Store.RemoveCapture(collected.LocalId, collected.Sha256);

            this.thisCollected = null;
            this.TakeOutOfCollectionButton.IsVisible = false;
            this.CollectedText.Text = changed is null ? "Could not take it out of the collection" : $"Taken out of [{CollectionSharing.NameOf(changed)}]";
            this.ShowCollectionChoices();

            if (changed is not null)
                context.CollectionChanged?.Invoke(collected.LocalId);
        }

        // The banner gone with its capture: what it said of the collection goes too.
        private void ForgetCollectedCapture()
        {
            this.thisCollected = null;
            this.CollectedText.IsVisible = false;
            this.TakeOutOfCollectionButton.IsVisible = false;
        }

        internal string? CollectedTextForTests => this.CollectedText.IsVisible ? this.CollectedText.Text : null;

        // The capture on screen, as showing it makes it (ShowTemporaryCapturedOscilloscopeImage) - without decoding a picture,
        // so a test can hand CollectCaptureAsync a file that is none.
        internal void CaptureOnScreenForTests(string? savedFilePath) => this.thisCapturedScopeImagePath = savedFilePath;

        // ###########################################################################################
        // THE COMPONENT AS THIS WINDOW SHOWS IT NOW (CollectionCaptures.ComponentOf): its title, category line, description and
        // purpose as the info pane shows them, and its pins as the pin list does - of the region shown.
        // ###########################################################################################
        internal CollectionComponent? ComponentShownForCollection()
        {
            ComponentEntry? entry = ComponentImageQueries.PickComponentEntry(this._allComponentEntries, this._localRegion);

            return CollectionCaptures.ComponentOf(
                this._boardLabel,
                ComponentPopupContent.TitleOf(this._boardLabel, entry),
                ComponentPopupContent.CategoryLine(entry),
                entry?.Description,
                entry?.Purpose,
                this._pinRows);
        }

        // ###########################################################################################
        // THE DRAWN PINOUT AS THIS WINDOW SHOWS IT NOW - the case, a part drawn as it looks, or a resistor's or capacitor's bands -
        // with the pin of the image shown marked (ShowViewedPin), as a PNG for the collection's page (owner request, 2026-11-27,
        // asked: "Also the drawn pinout"). CRT's own drawing, pictured as it stands on screen in the theme shown: drawn off
        // screen it would find none of the theme's colors (PinoutDrawing.Brush asks the tree it stands in). On the window's
        // own ground - a picture is opaque on the page - at twice its size for a sharp page, less when that would pass what a
        // collection takes (CollectionPng). Null when no drawing is shown, or it cannot be pictured.
        // ###########################################################################################
        internal byte[]? PinoutPictureForCollection()
        {
            if (this.PinoutPictureOverrideForTests is { } standIn)
                return standIn();

            Avalonia.Controls.Control? drawing = new Avalonia.Controls.Control[] { this.InfoPinoutDrawing, this.InfoPartDrawing, this.InfoResistorPicture }
                .FirstOrDefault(control => control.IsEffectivelyVisible && control.Bounds.Width >= 1 && control.Bounds.Height >= 1);

            if (drawing is null)
                return null;

            Avalonia.Size size = drawing.Bounds.Size;
            double scale = Math.Min(2.0, Math.Min(CollectionPng.MaximumWidth / size.Width, CollectionPng.MaximumHeight / size.Height));
            var pixels = new Avalonia.PixelSize(Math.Max(1, (int)Math.Floor(size.Width * scale)), Math.Max(1, (int)Math.Floor(size.Height * scale)));
            var dpi = new Avalonia.Vector(96 * scale, 96 * scale);

            try
            {
                using var drawn = new Avalonia.Media.Imaging.RenderTargetBitmap(pixels, dpi);
                drawn.Render(drawing);

                // Rendering clears its bitmap first, so the ground goes on a second one, under the drawing.
                using var framed = new Avalonia.Media.Imaging.RenderTargetBitmap(pixels, dpi);

                using (Avalonia.Media.DrawingContext context = framed.CreateDrawingContext())
                {
                    context.FillRectangle(this.Background ?? Avalonia.Media.Brushes.White, new Avalonia.Rect(size));
                    context.DrawImage(drawn, new Avalonia.Rect(size));
                }

                using var stream = new MemoryStream();
                framed.Save(stream, new Avalonia.Media.Imaging.PngBitmapEncoderOptions());

                return stream.Length == 0 ? null : stream.ToArray();
            }
            catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or NotSupportedException)
            {
                Logger.Warning($"Could not picture the pinout of [{this._boardLabel}] for a collection: {ex.Message}");
                return null;
            }
        }

        // Stands in for picturing the drawing in a test - the headless platform draws nothing. Null in the app.
        internal Func<byte[]?>? PinoutPictureOverrideForTests { get; set; }
    }
}
