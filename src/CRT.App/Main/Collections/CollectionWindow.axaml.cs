using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // WHAT THE COLLECTION WINDOW NEEDS FROM CRT (Main.Collections.cs): the collections on this computer, the server, the
    // account CRT is signed in to and the way to sign in, opening a link, and the clipboard - and, for "Attach collection to
    // worklog", whether the worklog is turned on and the attach itself (what came of it in words, or null for nothing done) - and
    // where the window was left and how to remember it, and the format "Copy link" wrote in last and how to remember it
    // (UserSettings; owner requests, 2026-11-28). A test hands its own.
    // ###########################################################################################
    internal sealed record CollectionWindowHost(
        CollectionStore Store,
        CollectionClient Client,
        Func<ReviewSession?> SignedIn,
        Func<Task> SignIn,
        Action<string> OpenLink,
        Func<string, Task> CopyText,
        Action Changed,
        Func<bool>? WorklogIsOn = null,
        Func<LocalCollection, Window, Task<string?>>? AttachToWorklog = null,
        Func<WindowPlace?>? Place = null,
        Action<WindowPlace>? RememberPlace = null,
        Func<CollectionLinkFormat>? LinkFormat = null,
        Action<CollectionLinkFormat>? RememberLinkFormat = null);

    // ###########################################################################################
    // ONE CAPTURE AS THE WINDOW LISTS IT - a view of the collection's capture; the note is what the window writes back
    // (CollectionWindow.EditedCollection). Its thumbnail and comments CHANGE IN PLACE as they arrive (code review,
    // 2026-11-27: each thumbnail decoded built the whole list again, and comments read again took the focus from a note being
    // typed).
    // ###########################################################################################
    public sealed class CollectionCaptureRow : INotifyPropertyChanged
    {
        private Bitmap? thisThumbnail;

        private bool thisShowsComments;

        private string thisCommentsHeading = string.Empty;

        private IReadOnlyList<CollectionCommentRow> thisComments = [];

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Sha256 { get; init; } = string.Empty;

        public Bitmap? Thumbnail
        {
            get => this.thisThumbnail;
            set => this.Set(ref this.thisThumbnail, value);
        }

        public string Heading { get; init; } = string.Empty;

        public string Settings { get; init; } = string.Empty;

        public string? Caption { get; set; }

        // Its comments on the page (step 3) - shown once they are read; the answer being typed.
        public bool ShowsComments
        {
            get => this.thisShowsComments;
            set => this.Set(ref this.thisShowsComments, value);
        }

        public string CommentsHeading
        {
            get => this.thisCommentsHeading;
            set => this.Set(ref this.thisCommentsHeading, value);
        }

        public IReadOnlyList<CollectionCommentRow> Comments
        {
            get => this.thisComments;
            set => this.Set(ref this.thisComments, value);
        }

        // The answer being typed - emptied once it is sent, so a second press sends nothing again (code review, 2026-11-27).
        public string? Answer
        {
            get => this.thisAnswer;
            set => this.Set(ref this.thisAnswer, value);
        }

        private string? thisAnswer;

        private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return;

            field = value;
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    // (One comment as the window lists it, CollectionCommentRow, is in Handlers/Collections since 2026-11-27 - pure.)

    // ###########################################################################################
    // AN OSCILLOSCOPE COLLECTION'S WINDOW (2026-11-26; "Show collection" in the component info window). Here it is named and
    // described, its captures noted - a thumbnail clicked opens the capture large in the worklogs' zooming viewer - or taken
    // out, and it is SHARED:
    //
    //   - "Share" sends it as it is now (CollectionClient) - the first time a page is made, its address kept; later, the page
    //     follows the change, and a page whose sharing was stopped comes back under its link. It needs an account (the account
    //     icon), and asks to sign in when there is none.
    //   - "Stop sharing" takes the page offline - its link then says it is not shared at the moment, its captures and comments
    //     kept on the server (owner decision, 2026-11-27: "Offline, same link back").
    //   - "Delete in" is how long the server keeps the page - a day, a week, a month or a year, counted from EVERY share (owner
    //     decisions, 2026-11-27: "every time the user shares it again, it resets its time"; 2026-11-28 the four), so "Share" is
    //     offered for a shared collection with nothing changed too: it keeps the page for another stretch. The date it goes is
    //     said bold (DeleteDateLine).
    //   - "Open shared page" while it is shared (owner request, 2026-11-28: "only if it is shared - if not shared, this button
    //     should be disabled"; "Open web page" until then), and "Copy link" - the link with a line about it, in the format picked
    //     last, its arrow listing BBCode, Markdown and a plain link with where each works (owner decision, 2026-11-28; "Copy for
    //     forum" and "Copy for Reddit" until then).
    //   - "Delete image" ("Take out" until 2026-11-28 - owner request) takes a capture out of it - and off a shared page AT ONCE,
    //     its deletion date left alone (owner request, 2026-11-28: "do not touch the expiry date").
    //   - "Attach collection to worklog" ("Attach to worklog" until 2026-11-28 - owner request) copies every capture into a
    //     worklog as photos (owner decision, 2026-11-27: "A copy of its captures"; CollectionWorklogAttach).
    //   - It opens where it was left, its size too (KeptWindowPlace; owner request, 2026-11-28), and Esc closes it - or, a delete
    //     being asked, keeps the collection.
    //   - "Delete collection" takes it off this computer and its page off the server, after asking. NOTHING ELSE EVER DELETES
    //     IT FROM THIS COMPUTER (owner request, 2026-11-27): a page whose time is up only forgets its page (ForgetPage).
    //
    // Every edit is kept on this computer as its box is left (SaveEdits) - nothing goes to the server until "Share". ("Looks
    // wrong", a mark per capture, went on 2026-11-27 - owner request: "No need for the 'Looks wrong' checkbox".)
    // ###########################################################################################
    public partial class CollectionWindow : Window
    {
        private CollectionWindowHost? thisHost;

        private string thisLocalId = string.Empty;

        private bool thisShowing;

        private bool thisBusy;

        // The list's rows - ONE collection, changed in place (Show), so a capture added or taken out elsewhere leaves every other
        // row, and the box being typed in, as it is.
        private readonly System.Collections.ObjectModel.ObservableCollection<CollectionCaptureRow> thisRows = [];

        // Where it was left, followed as it is moved and sized - remembered as it closes (the host's RememberPlace).
        private readonly KeptWindowPlace thisPlace;

        public CollectionWindow()
        {
            this.InitializeComponent();

            this.thisPlace = new KeptWindowPlace(this);

            this.BuildCopyLinkMenu();

            // No taller than the screen it opens on - its buttons stay in reach on a laptop's 768 (code review, 2026-11-27); the
            // captures scroll. Never re-centred as it is resized: it is a window people size themselves.
            this.Opened += (_, _) =>
            {
                if (this.Screens.ScreenFromWindow(this) is { } screen)
                    this.MaxHeight = Handlers.Geometry.DialogPlacement.MaxHeight(screen.WorkingArea.Height, screen.Scaling);
            };
        }

        // False when there is no such collection (any more) - the window is then not to be shown: Show(null) closes it, and
        // Avalonia refuses to show a window closed (code review, 2026-11-27: "Show collection" threw on one deleted elsewhere).
        internal bool Initialize(CollectionWindowHost host, string localId)
        {
            this.thisHost = host;
            this.thisLocalId = localId;

            if (host.Store.Find(localId) is not { } collection)
                return false;

            // Where a collection's window was left (owner request, 2026-11-28) - a place on no screen any more is given up, and it
            // opens over CRT's window as the first time.
            if (host.Place?.Invoke() is { } place)
                this.thisPlace.Use(place, ownerScreen: null);

            this.LifetimeComboBox.ItemsSource = CollectionRules.Lifetimes.Select(lifetime => lifetime.Words).ToList();
            this.thisLinkFormat = host.LinkFormat?.Invoke() ?? CollectionLinkFormat.BbCode;
            this.ShowLinkFormat();
            this.Show(collection);

            // Its comments, when it is shared and CRT is signed in - read as it opens, never asking to sign in for them.
            _ = HandlerGuard.RunAsync(this.LoadCommentsAsync, this.Say);
            return true;
        }

        internal string LocalId => this.thisLocalId;

        // ###########################################################################################
        // The collection changed somewhere else - a capture added or taken out in a component info window (code review,
        // 2026-11-27: the window kept its old rows, and a shared one's "Share" stayed greyed out). What the boxes hold is kept
        // first, then everything shown again from the collection as kept.
        // ###########################################################################################
        internal void ShowAgain()
        {
            this.SaveEdits();
            this.Show(this.Collection);
        }

        private LocalCollection? Collection => this.thisHost?.Store.Find(this.thisLocalId);

        // ###########################################################################################
        // Everything shown, from the collection as kept.
        // ###########################################################################################
        private void Show(LocalCollection? collection)
        {
            if (collection is null)
            {
                this.Close();
                return;
            }

            this.thisShowing = true;

            try
            {
                this.Title = $"Oscilloscope collection: {CollectionSharing.NameOf(collection)}";
                this.TitleBox.Text = collection.Title ?? string.Empty;
                this.DescriptionBox.Text = collection.Description ?? string.Empty;
                this.ForumLinksBox.Text = string.Join("\n", collection.ForumLinks);
                this.BoardText.Text = CollectionListing.BoardOf(collection);
                this.LifetimeComboBox.SelectedItem = CollectionRules.LifetimeOf(collection.Lifetime).Words;

                IReadOnlyList<LocalCapture> captures = CollectionSharing.InPageOrder(collection.Captures);

                this.CapturesHeading.Text = captures.Count == 0
                    ? "No captures yet - pick this collection under \"Collect captures into\" in a component's info window, and capture."
                    : $"{captures.Count} capture{(captures.Count == 1 ? string.Empty : "s")}";

                // *** THE ROWS ARE CHANGED IN PLACE (code review, 2026-11-27). *** A capture added or taken out in a component
                // info window adds or takes out its row alone - made again, every row lost its container, and a caption or answer
                // being typed lost the caret and the keyboard at each capture. A capture is one row by its checksum (one screen is
                // one capture), its words fixed but for what this window edits, which a kept row holds.
                InPlaceList.Sync(
                    this.thisRows,
                    captures,
                    (row, capture) => string.Equals(row.Sha256, capture.Sha256, StringComparison.Ordinal),
                    this.RowOf,
                    itemKey: row => row.Sha256,
                    wantedKey: capture => capture.Sha256);

                if (!ReferenceEquals(this.CapturesList.ItemsSource, this.thisRows))
                    this.CapturesList.ItemsSource = this.thisRows;

                this.LoadThumbnails(collection.LocalId);
            }
            finally
            {
                this.thisShowing = false;
            }

            this.ShowSharedState(collection);
        }

        private CollectionCaptureRow RowOf(LocalCapture capture)
        {
            var row = new CollectionCaptureRow
            {
                Sha256 = capture.Sha256,
                Heading = CollectionRules.WhereOf(CollectionSharing.EntryOf(capture)) + (capture.SignalName is { } signal ? $" - {signal}" : string.Empty),
                Settings = CollectionSharing.SettingsLine("This capture", capture.TimePerDivision, capture.VoltsPerDivision, capture.TriggerLevel) +
                           (capture.ExpectedReading is { } reading ? $" | Expected reading: {reading}" : string.Empty),
                Caption = capture.Caption,
                Thumbnail = this.thisThumbnails.GetValueOrDefault(capture.Sha256)
            };

            this.ShowCommentsOn(row);
            return row;
        }

        // A capture's comments, as read last - set on its row in place.
        private void ShowCommentsOn(CollectionCaptureRow row)
        {
            List<CollectionCommentRow> comments = (this.thisComments ?? [])
                .Where(comment => comment.ImageSha256 == row.Sha256)
                .Select(comment => CollectionCommentRow.Of(comment, this.thisSeenBefore))
                .ToList();

            int fresh = comments.Count(comment => comment.IsNew);

            row.ShowsComments = this.thisComments is not null;
            row.CommentsHeading = comments.Count == 0
                ? "No comments yet"
                : $"{comments.Count} comment{(comments.Count == 1 ? string.Empty : "s")}" + (fresh > 0 ? $" - {fresh} new" : string.Empty);
            row.Comments = comments;
        }

        // ###########################################################################################
        // THE COMMENTS (step 3): read from the server while the collection is shared and CRT is signed in - when the window opens,
        // and after anything is done to one. What is newer than the owner had read is marked NEW for as long as the window is
        // open; the newest read is then kept (SeenCommentId), so the popups' "new comments" goes.
        // ###########################################################################################
        private IReadOnlyList<CollectionCommentAnswer>? thisComments;

        private long thisSeenBefore = long.MaxValue;

        internal async Task LoadCommentsAsync()
        {
            if (this.thisHost is not { } host || this.Collection is not { ServerId: { } serverId } collection ||
                host.SignedIn() is not { } session)
            {
                return;
            }

            // Another account's: said, so no comment section missing reads as "no comments" (code review, 2026-11-27).
            if (CollectionSharing.OtherAccountLine(collection, session.AccountId) is { } otherAccount)
            {
                this.Say(otherAccount);
                return;
            }

            CollectionClient.CommentsResult result = await host.Client.ListCommentsAsync(serverId, session.BearerToken);

            if (result.Kind == CollectionClient.Outcome.GoneFromServer)
            {
                this.ForgetPage();
                return;
            }

            if (result.Kind != CollectionClient.Outcome.Done)
            {
                this.Say(result.Message);
                return;
            }

            if (this.thisComments is null)
                this.thisSeenBefore = collection.SeenCommentId;

            this.thisComments = result.Comments;

            // The collection AS KEPT NOW - an edit kept while the comments were on their way stays (code review, 2026-11-27) -
            // and the rows changed in place, so a box being typed in keeps its text and the keyboard.
            long newest = result.Comments.Count == 0 ? 0 : result.Comments.Max(comment => comment.Id);

            bool seen = false;

            host.Store.Update(
                this.thisLocalId,
                current =>
                {
                    seen = newest > current.SeenCommentId;
                    return seen ? current with { SeenCommentId = newest } : current;
                },
                out bool seenKept);

            if (seen && seenKept)
                host.Changed();

            foreach (CollectionCaptureRow row in this.thisRows)
                this.ShowCommentsOn(row);
        }

        private async void OnAnswerClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.AnswerAsync(sender), this.Say);

        // The answer typed under a capture, sent as a comment of the owner's - and its box emptied once it is.
        internal async Task AnswerAsync(object? sender)
        {
            if (sender is not Button { Tag: string sha256 } || this.thisRows.FirstOrDefault(row => row.Sha256 == sha256) is not { } row)
                return;

            string answer = row.Answer ?? string.Empty;
            bool done = await this.CommentActionAsync((client, serverId, token) => client.AddCommentAsync(serverId, sha256, answer, token));

            // Sent - or, its answer not come in time, among the comments read back: the box emptied ON THE ROW SHOWN NOW, which a
            // capture added meanwhile made again (code review, 2026-11-27: the old row was emptied, the text stayed to be sent
            // twice).
            bool posted = done || (this.thisComments ?? []).Any(comment =>
                comment.Mine && comment.ImageSha256 == sha256 && comment.Body == CollectionComments.Tidied(answer));

            if (posted && this.thisRows.FirstOrDefault(shown => shown.Sha256 == sha256) is { } shownRow &&
                string.Equals(shownRow.Answer, answer, StringComparison.Ordinal))
            {
                shownRow.Answer = null;
            }
        }

        private async void OnHideCommentClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.CommentActionAsync(sender, (client, serverId, id, token) => client.SetCommentHiddenAsync(serverId, id, true, token)), this.Say);

        private async void OnShowCommentClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.CommentActionAsync(sender, (client, serverId, id, token) => client.SetCommentHiddenAsync(serverId, id, false, token)), this.Say);

        private async void OnDeleteCommentClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.CommentActionAsync(sender, (client, serverId, id, token) => client.DeleteCommentAsync(serverId, id, token)), this.Say);

        private Task CommentActionAsync(object? sender, Func<CollectionClient, string, long, string, Task<CollectionClient.Result>> act) =>
            sender is Button { Tag: long id }
                ? this.CommentActionAsync((client, serverId, token) => act(client, serverId, id, token))
                : Task.CompletedTask;

        // ###########################################################################################
        // One thing done to the comments, then them read again - what went wrong said. ONE AT A TIME, UNDER THE WINDOW'S
        // BusyOverlay (code review, 2026-11-27: "Answer" clicked twice posted the answer twice), and by the account that shared
        // the collection alone (CollectionSharing.OtherAccountLine). A server that has not answered in the wait's time is waited
        // for still, the window busy, and the comments read again once it has - never "it failed".
        // ###########################################################################################
        // True when the server said it was done.
        internal async Task<bool> CommentActionAsync(Func<CollectionClient, string, string, Task<CollectionClient.Result>> act)
        {
            if (this.thisBusy || this.thisHost is not { } host || this.Collection?.ServerId is not { } serverId)
                return false;

            if (host.SignedIn() is null)
            {
                this.Say(CollectionClient.SignInFirst);
                await host.SignIn();
            }

            // Asked again once signed in: the account window leaves this one clickable, and a second press made meanwhile
            // would go on too (code review, 2026-11-27).
            if (this.thisBusy || host.SignedIn() is not { } session || this.Collection is not { } collection)
                return false;

            if (CollectionSharing.OtherAccountLine(collection, session.AccountId) is { } otherAccount)
            {
                this.Say(otherAccount);
                return false;
            }

            bool done = false;

            this.SetBusy(true);

            try
            {
                WaitResult<CollectionClient.Result> waited =
                    await BusyOverlay.RunAsync(this, "Telling the server", _ => act(host.Client, serverId, session.BearerToken));

                // *** NOT ANSWERED IN THE WAIT'S TIME, THE WINDOW STAYS BUSY UNTIL THE REQUEST REALLY ENDS (code review, 2026-11-27),
                // as "Share" does. *** Let go at the overlay's limit, "Answer" could be pressed again while the first was still on
                // its way - the answer posted twice, its writer mailed twice.
                if (waited.IsTimedOut)
                    this.Say(CollectionWindow.NoAnswerYet);

                CollectionClient.Result result = waited.IsTimedOut ? await waited.Work : waited.Value!;

                if (result.Kind == CollectionClient.Outcome.GoneFromServer)
                {
                    this.ForgetPage();
                    return false;
                }

                if (result.Kind != CollectionClient.Outcome.Done)
                {
                    this.Say(result.Message);
                    return false;
                }

                done = true;
                this.Say(string.Empty);
            }
            finally
            {
                this.SetBusy(false);
            }

            await this.LoadCommentsAsync();
            return done;
        }

        internal const string NoAnswerYet = "The server has not answered for a while - still waiting for it to finish.";

        // The thumbnails, decoded small off the UI thread ONCE each while the window is open (thisThumbnails), each set on its
        // row as it comes - never disposed by hand (a picture still drawn takes the render thread down; CLAUDE.md).
        private readonly Dictionary<string, Bitmap> thisThumbnails = new(StringComparer.Ordinal);

        private void LoadThumbnails(string localId)
        {
            if (this.thisHost is not { } host)
                return;

            // A capture of no checksum (a file edited by hand, or written by a later CRT) has no picture to show - and CapturePath
            // would throw on it (code review, 2026-11-27).
            foreach (CollectionCaptureRow row in this.thisRows.Where(row => row.Thumbnail is null && DataTreeLayout.IsSha256(row.Sha256)))
            {
                string path = host.Store.CapturePath(localId, row.Sha256);

                _ = Task.Run(() =>
                {
                    try
                    {
                        // Opened so it can be deleted while read - a capture taken out as its thumbnail decodes (code review,
                        // 2026-11-27; BaselineCaptureTarget.OpenCapture says why).
                        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                        return Bitmap.DecodeToWidth(stream, 400);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException)
                    {
                        return null;
                    }
                }).ContinueWith(
                    decoded =>
                    {
                        if (decoded.Result is not { } bitmap)
                            return;

                        this.thisThumbnails[row.Sha256] = bitmap;

                        // The row shown now for that capture - the list may have been made again meanwhile.
                        if (this.thisRows.FirstOrDefault(shown => shown.Sha256 == row.Sha256) is { } shownRow)
                            shownRow.Thumbnail = bitmap;
                    },
                    TaskScheduler.FromCurrentSynchronizationContext());
            }
        }

        // ###########################################################################################
        // Shared or not, and whether it changed since - and which buttons there is anything for.
        // ###########################################################################################
        private void ShowSharedState(LocalCollection collection)
        {
            bool changed = CollectionSharing.HasChangesNotShared(collection);

            string state = CollectionWindow.SharedStateLine(collection, changed);
            this.SharedStateText.Text = state;
            this.SharedStateText.IsVisible = state.Length > 0;

            BoldRuns.ShowMarked(this.DeleteDateText, CollectionWindow.DeleteDateLine(collection));

            // Shared, "Share" is offered changed or not: it keeps the page another stretch, counted from that day.
            this.ShareButton.IsEnabled = !this.thisBusy && collection.Captures.Count > 0;
            this.StopSharingButton.IsVisible = collection.IsOnline;
            this.StopSharingButton.IsEnabled = !this.thisBusy;

            // Only while it is shared - offline, or never shared, there is no page to show anybody (owner request, 2026-11-28).
            this.OpenPageButton.IsEnabled = !this.thisBusy && collection.IsOnline && collection.PageUrl is not null;

            // What is copied is the page's link - changed since, the page is not what was described (code review, 2026-11-27):
            // "Share" first. And its sharing stopped, the link says only that it is not shared at the moment.
            this.CopyLinkButton.IsEnabled = !this.thisBusy && collection.IsOnline && !changed;
            this.DeleteButton.IsEnabled = !this.thisBusy;

            this.AttachToWorklogButton.IsVisible = this.thisHost?.AttachToWorklog is not null && this.thisHost.WorklogIsOn?.Invoke() != false;
            this.AttachToWorklogButton.IsEnabled = !this.thisBusy && collection.Captures.Count > 0;
        }

        // ###########################################################################################
        // When the page goes, its date bold, and what "Share" does to it - counted from every share (owner decision, 2026-11-27;
        // the words the owner's, 2026-11-28: "Its page is deleted on <b>2027-November-28</b>. Clicking "Share" again will update the
        // deletion day."). Marked for BoldRuns.ShowMarked.
        // ###########################################################################################
        internal static string DeleteDateLine(LocalCollection collection)
        {
            ArgumentNullException.ThrowIfNull(collection);

            return collection.IsShared && collection.DeleteAtUtc is { } deleteAt
                ? $"Its page is deleted on <b>{CollectionRules.DateText(deleteAt)}</b>. Clicking \"Share\" again will update the deletion day."
                : $"Counted from the day it is shared - {CollectionRules.LifetimeOf(collection.Lifetime).Words}.";
        }

        // ###########################################################################################
        // Where the collection stands, in a sentence - kept here only, shared and changed since, or shared and stopped; shared as it
        // is, NOTHING (owner request, 2026-11-28, of "Shared on 2026-...": "no need for that") - the date it goes and "Open shared
        // page" say the rest.
        // ###########################################################################################
        internal static string SharedStateLine(LocalCollection collection, bool changedSinceShared)
        {
            ArgumentNullException.ThrowIfNull(collection);

            if (!collection.IsShared)
                return "Not shared yet - it is on this computer only. \"Share\" makes its page, which anybody with the link can see.";

            if (!collection.IsOnline)
            {
                return $"Not shared at the moment - its page is offline, and its link says so. \"Share\" brings it back under the same " +
                       $"link: {collection.PageUrl}";
            }

            return changedSinceShared ? "Changed since it was shared - \"Share\" sends the change." : string.Empty;
        }

        internal void Say(string text) => this.StatusText.Text = text;

        // ###########################################################################################
        // THE COLLECTION AS THE WINDOW HOLDS IT - its boxes, the "Delete in" picked and each capture's note and mark - over the
        // collection as kept.
        // ###########################################################################################
        internal LocalCollection? EditedCollection() => this.Collection is { } collection ? this.EditedFrom(collection) : null;

        // The window's edits over `collection` - the one kept now, inside the store's Update.
        private LocalCollection EditedFrom(LocalCollection collection)
        {
            Dictionary<string, CollectionCaptureRow> rows = this.thisRows.ToDictionary(row => row.Sha256, StringComparer.Ordinal);

            string lifetime = CollectionRules.Lifetimes
                .FirstOrDefault(choice => choice.Words == this.LifetimeComboBox.SelectedItem as string)?.Key ?? collection.Lifetime;

            return collection with
            {
                Title = CollectionWindow.Blank(this.TitleBox.Text),
                Description = CollectionWindow.Blank(this.DescriptionBox.Text),
                ForumLinks = (this.ForumLinksBox.Text ?? string.Empty)
                    .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .ToList(),
                Lifetime = lifetime,
                Captures = collection.Captures
                    .Select(capture => rows.TryGetValue(capture.Sha256, out CollectionCaptureRow? row)
                        ? capture with { Caption = CollectionWindow.Blank(row.Caption) }
                        : capture)
                    .ToList()
            };
        }

        internal const string NotKept = "Could not keep the change on this computer - see \"Errors and warnings\" on the \"Configuration\" tab.";

        // Kept on this computer - true when it was (or nothing changed). Made on the collection as kept now (CollectionStore
        // .Update), so a capture a component info window added meanwhile stays.
        internal bool SaveEdits()
        {
            if (this.thisShowing || this.thisHost is not { } host)
                return true;

            LocalCollection? before = null;

            LocalCollection? edited = host.Store.Update(
                this.thisLocalId,
                collection =>
                {
                    before = collection;
                    LocalCollection changed = this.EditedFrom(collection);

                    // The records hold lists, compared by reference - the same content is the same JSON.
                    return System.Text.Json.JsonSerializer.Serialize(changed) == System.Text.Json.JsonSerializer.Serialize(collection) ? collection : changed;
                },
                out bool kept);

            if (!kept)
            {
                this.Say(CollectionWindow.NotKept);
                return false;
            }

            if (edited is null || ReferenceEquals(edited, before))
                return true;

            this.Title = $"Oscilloscope collection: {CollectionSharing.NameOf(edited)}";
            this.ShowSharedState(edited);
            host.Changed();

            return true;
        }

        private void OnEditLeft(object? sender, RoutedEventArgs e) => this.SaveEdits();

        // ###########################################################################################
        // A THUMBNAIL CLICKED OPENS THE CAPTURE LARGE (owner request, 2026-11-27: "the same possibility as other places where we
        // have small images/thumbnails - it should open in another window and be zoomable") - the worklogs' viewer, which zooms
        // with the wheel, its title where the capture was taken and its comment the settings, the baseline's and the note, the
        // values bold. Shown beside this window and gone at a click outside it, as the capture window's pictures are
        // (BaselineCaptureWindow.ShowPicture).
        // ###########################################################################################
        private void OnThumbnailTapped(object? sender, TappedEventArgs e)
        {
            if (sender is Control { DataContext: CollectionCaptureRow row })
                this.OpenCapture(row.Sha256);
        }

        internal void OpenCapture(string sha256)
        {
            if (this.thisHost is not { } host || !DataTreeLayout.IsSha256(sha256) ||
                this.Collection?.Captures.FirstOrDefault(capture => capture.Sha256 == sha256) is not { } capture ||
                this.thisRows.FirstOrDefault(row => row.Sha256 == sha256) is not { } row)
            {
                return;
            }

            var viewer = new WorklogPhotoViewerWindow();
            viewer.Initialize(row.Heading, CollectionSharing.WorklogCommentOf(capture), host.Store.CapturePath(this.thisLocalId, sha256));
            viewer.ClosesWhenLeft = true;

            if (this.ShowPictureOverrideForTests is { } show)
            {
                show(viewer);
                return;
            }

            viewer.Show(this);
        }

        // Stands in for showing the viewer in a test. Null in the app.
        internal Action<WorklogPhotoViewerWindow>? ShowPictureOverrideForTests { get; set; }

        // ###########################################################################################
        // "Delete in": kept, and counted from the next share - which is when the server is told (CollectionSharing.DeleteAtFor).
        // ###########################################################################################
        private void OnLifetimePicked(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisShowing || this.thisHost is not { } host)
                return;

            LocalCollection? edited = host.Store.Update(this.thisLocalId, this.EditedFrom, out bool kept);

            if (kept && edited is not null)
            {
                this.ShowSharedState(edited);
                host.Changed();
            }
        }

        // While sharing, nothing is taken out: the share reads each capture the server asks for as it goes, and one taken out
        // meanwhile - its file deleted - stopped it halfway (code review, 2026-11-27). The popups ask IsBusy for the same.
        internal bool IsBusy => this.thisBusy;

        internal const string WaitForTheShare = "Wait for the server to answer - a capture is taken out once it has.";

        private async void OnTakeOutClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: string sha256 })
                await HandlerGuard.RunAsync(() => this.TakeOutAsync(sha256), this.Say);
        }

        // ###########################################################################################
        // "DELETE IMAGE". A shared collection's capture goes from its page AT ONCE (owner request, 2026-11-28: "should also remove it
        // right away from the shared page, and do not touch the expiry date. This date should only change when the USER does
        // something") - that one capture alone is taken off the server, signed in with the account that shared it, under the
        // window's BusyOverlay; then from this computer. Nothing else is sent: the date, the sharing and a change not shared yet
        // stay as they are. A server that cannot take it off keeps it here too, so it can be tried again; a page the server no
        // longer has is forgotten, and the capture deleted here.
        // ###########################################################################################
        internal async Task TakeOutAsync(string sha256)
        {
            if (this.thisHost is not { } host)
                return;

            if (this.thisBusy)
            {
                this.Say(CollectionWindow.WaitForTheShare);
                return;
            }

            this.SaveEdits();

            if (this.Collection is not { ServerId: { } serverId })
            {
                this.DeleteHere(sha256, CollectionWindow.ImageDeleted);
                return;
            }

            if (host.SignedIn() is null)
            {
                this.Say(CollectionWindow.SignInToDeleteImage);
                await host.SignIn();
            }

            if (this.thisBusy || host.SignedIn() is not { } session || this.Collection is not { } collection)
                return;

            if (CollectionSharing.OtherAccountLine(collection, session.AccountId) is { } otherAccount)
            {
                this.Say(otherAccount);
                return;
            }

            this.SetBusy(true);

            try
            {
                WaitResult<CollectionClient.Result> waited =
                    await BusyOverlay.RunAsync(this, "Deleting the image from its shared page", _ => host.Client.RemoveImageAsync(serverId, sha256, session.BearerToken));

                if (waited.IsTimedOut)
                    this.Say(CollectionWindow.NoAnswerYet);

                CollectionClient.Result result = waited.IsTimedOut ? await waited.Work : waited.Value!;

                switch (result.Kind)
                {
                    case CollectionClient.Outcome.Done:
                        this.DeleteHere(sha256, CollectionWindow.ImageDeletedShared);
                        break;

                    case CollectionClient.Outcome.GoneFromServer:
                        this.ForgetPage();
                        this.DeleteHere(sha256, $"{CollectionWindow.ImageDeleted} {CollectionWindow.PageGone}");
                        break;

                    default:
                        this.Say(result.Message);
                        break;
                }
            }
            finally
            {
                this.SetBusy(false);
            }
        }

        // ###########################################################################################
        // The capture deleted from this computer, and what is said. ON A PAGE THAT SHOWED EVERYTHING THE COLLECTION HOLDS, IT STILL
        // DOES: its page is the collection less that capture now, so what was shared is taken to be this collection as it now is -
        // "changed since it was shared" is not said for a delete the page has followed. A collection with a change not shared yet
        // keeps saying so.
        // ###########################################################################################
        private void DeleteHere(string sha256, string said)
        {
            if (this.thisHost is not { } host)
                return;

            bool wasInStep = this.Collection is { } before && before.IsShared && !CollectionSharing.HasChangesNotShared(before);

            if (host.Store.RemoveCapture(this.thisLocalId, sha256) is not { } changed)
                return;

            if (wasInStep && changed.IsShared && changed.DeleteAtUtc is { } deleteAt &&
                host.Store.Update(
                    this.thisLocalId,
                    current => current with { SharedContentHash = CollectionSharing.ContentHash(CollectionSharing.RequestOf(current, deleteAt)) },
                    out _) is { } inStep)
            {
                changed = inStep;
            }

            this.Show(changed);
            host.Changed();
            this.Say(said);
        }

        internal const string ImageDeleted = "The image is deleted from this collection.";

        internal const string ImageDeletedShared = "The image is deleted from this collection and from its shared page.";

        internal const string SignInToDeleteImage = "Sign in first - the account that shared it takes the image off its shared page.";

        // ###########################################################################################
        // "SHARE": signed in first, then the collection as it is now, and what came of it kept and said.
        // ###########################################################################################
        private async void OnShareClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.ShareAsync, this.Say);

        internal async Task ShareAsync()
        {
            if (this.thisHost is not { } host || !this.SaveEdits() || this.Collection is null)
                return;

            if (host.SignedIn() is null)
            {
                this.Say(CollectionClient.SignInFirst);
                await host.SignIn();
            }

            // The collection as kept once signed in - an edit kept while signing in goes with it.
            if (host.SignedIn() is not { } session || this.thisBusy || this.Collection is not { } collection)
                return;

            // ###########################################################################################
            // ANOTHER ACCOUNT'S: asked first - pressed again, it is shared AS A NEW PAGE with the account signed in, its captures
            // kept and the old page left to its date (code review, 2026-11-27: with that account gone - the test instance's data
            // reset deletes them - it could never be shared again, only deleted whole).
            // ###########################################################################################
            if (CollectionSharing.OtherAccountLine(collection, session.AccountId) is { } otherAccount)
            {
                if (!this.thisShareAsNewAsked)
                {
                    this.thisShareAsNewAsked = true;
                    this.Say(CollectionWindow.ShareAsNewLine(otherAccount));
                    return;
                }

                this.thisShareAsNewAsked = false;

                if (this.Unshare() is not { } unshared)
                    return;

                collection = unshared;
            }

            this.SetBusy(true);

            try
            {
                // ###########################################################################################
                // UNDER THE WINDOW'S BusyOverlay, each step said on it (code review, 2026-11-27: every wait the user watches is
                // under it). A share that goes quiet past the wait's limit may still be done: the overlay lifts, and the
                // share's own end - its HttpClient's time at most - is waited for and kept, never "it failed".
                // ###########################################################################################
                // A first share whose answer is lost is sent again under the collection's own key - the server then changes the
                // page it made, never makes another (CollectionClient, CollectionRoutes.ClientKeyHeader).
                DateTimeOffset sentAt = DateTimeOffset.UtcNow;

                WaitResult<CollectionClient.Result> waited = await BusyOverlay.RunAsync(this, "Sharing the collection", context =>
                    host.Client.ShareAsync(
                        collection,
                        session.BearerToken,
                        sha256 => host.Store.ReadCapture(collection.LocalId, sha256),
                        sentAt,
                        new Progress<string>(step =>
                        {
                            context.Report(step);
                            this.Say(step);
                        })));

                if (waited.IsTimedOut)
                    this.Say("The server has not answered for a while - still waiting for the share to finish.");

                CollectionClient.Result result = waited.IsTimedOut ? await waited.Work : waited.Value!;

                this.AfterShare(session, result);
            }
            finally
            {
                this.SetBusy(false);
            }
        }

        // ###########################################################################################
        // What the share came to, kept ON THE COLLECTION AS IT IS NOW (code review, 2026-11-27): a note or a title kept while it
        // was on its way stays - only the server's part is written, and the boxes are left as they are. What was sent is the
        // checksum kept (SharedContentHash), so such an edit reads as "changed since".
        // ###########################################################################################
        private void AfterShare(ReviewSession session, CollectionClient.Result result)
        {
            if (this.thisHost is not { } host)
                return;

            switch (result.Kind)
            {
                case CollectionClient.Outcome.Shared when result.Answer is { } answer:
                    LocalCollection? shared = host.Store.Update(
                        this.thisLocalId,
                        current => current with
                        {
                            ServerId = answer.Id,
                            PageUrl = answer.Url,
                            DeleteAtUtc = result.SentDeleteAtUtc ?? answer.DeleteAtUtc,
                            SharedUtc = DateTimeOffset.UtcNow,
                            SharedContentHash = result.SharedContentHash,
                            SharedByAccountId = session.AccountId == 0 ? current.SharedByAccountId : session.AccountId,
                            SharedByName = session.AccountId == 0 ? current.SharedByName : session.DisplayName,

                            // A share brings a page whose sharing was stopped back - the server clears it with the content.
                            SharingStoppedUtc = null
                        },
                        out _);

                    if (shared is not null)
                        this.ShowSharedState(shared);

                    this.Say($"Shared - anybody with the link sees it: {answer.Url}. \"Copy link\" puts the link on the clipboard, with a line about it.");
                    break;

                case CollectionClient.Outcome.GoneFromServer:
                    this.ForgetPage();
                    break;

                default:
                    this.Say(result.Message);
                    break;
            }

            host.Changed();
        }

        // ###########################################################################################
        // ITS PAGE IS NO LONGER ON THE SERVER - its time up, or taken down by the administrator on the website (code review,
        // 2026-11-27: the comments' "not found" was only said, and the window went on showing a page that was gone, "Share"
        // greyed out). Found by a share, the comments read or an action on one: the page forgotten on the collection as kept
        // now, its comments with it, and "Share" makes a new page.
        // ###########################################################################################
        internal const string PageGone =
            "Its page is no longer on the server - its time was up, or it was taken down. \"Share\" makes a new page.";

        private void ForgetPage()
        {
            if (this.Unshare() is not null)
                this.Say(CollectionWindow.PageGone);
        }

        // The page forgotten on the collection as kept now - its comments with it - the captures and words kept. The collection
        // as it now is, or null when there is none.
        private LocalCollection? Unshare()
        {
            if (this.thisHost is not { } host ||
                host.Store.Update(
                    this.thisLocalId,
                    current => current with
                    {
                        ServerId = null,
                        PageUrl = null,
                        DeleteAtUtc = null,
                        SharedUtc = null,
                        SharedContentHash = null,
                        SharedByAccountId = null,
                        SharedByName = null,
                        SharingStoppedUtc = null
                    },
                    out _) is not { } unshared)
            {
                return null;
            }

            this.ShowSharedState(unshared);

            this.thisComments = null;

            foreach (CollectionCaptureRow row in this.thisRows)
                this.ShowCommentsOn(row);

            host.Changed();
            return unshared;
        }

        // Asked once a share meets a collection of another account - the next "Share" shares it as a new page.
        private bool thisShareAsNewAsked;

        internal static string ShareAsNewLine(string otherAccount) =>
            $"{otherAccount} Or press \"Share\" again to share it as a new page with the account you are signed in to - " +
            "its captures are kept, and the old page stays up until its date.";

        // Busy, "Delete in" waits too: the date a share sends is written back over the collection once it is in (code review,
        // 2026-11-27: one picked meanwhile was written over, the page then deleted at the old date).
        private void SetBusy(bool busy)
        {
            this.thisBusy = busy;
            this.LifetimeComboBox.IsEnabled = !busy;

            if (this.Collection is { } collection)
                this.ShowSharedState(collection);
        }

        // ###########################################################################################
        // "STOP SHARING" (owner decision, 2026-11-27: "Offline, same link back"): its page offline - its link says it is not
        // shared at the moment - its captures and comments kept on the server, its date counting; "Share" brings it back under
        // the same link. Signed in with the account that shared it, under the window's BusyOverlay, as a comment's action is - a
        // server not answering in the wait's time waited for still, never "it failed".
        // ###########################################################################################
        private async void OnStopSharingClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.StopSharingAsync, this.Say);

        internal const string SharingStopped =
            "Not shared any more - its page is offline, and its link says it is not shared at the moment. \"Share\" brings it back " +
            "under the same link.";

        internal async Task StopSharingAsync()
        {
            if (this.thisBusy || this.thisHost is not { } host || this.Collection is not { ServerId: { } serverId })
                return;

            this.SaveEdits();

            if (host.SignedIn() is null)
            {
                this.Say(CollectionClient.SignInFirst);
                await host.SignIn();
            }

            if (this.thisBusy || host.SignedIn() is not { } session || this.Collection is not { IsOnline: true } collection)
                return;

            if (CollectionSharing.OtherAccountLine(collection, session.AccountId) is { } otherAccount)
            {
                this.Say(otherAccount);
                return;
            }

            this.SetBusy(true);

            try
            {
                WaitResult<CollectionClient.Result> waited =
                    await BusyOverlay.RunAsync(this, "Taking its page offline", _ => host.Client.SetSharingAsync(serverId, false, session.BearerToken));

                if (waited.IsTimedOut)
                    this.Say(CollectionWindow.NoAnswerYet);

                CollectionClient.Result result = waited.IsTimedOut ? await waited.Work : waited.Value!;

                switch (result.Kind)
                {
                    case CollectionClient.Outcome.Done:
                        if (host.Store.Update(this.thisLocalId, current => current with { SharingStoppedUtc = DateTimeOffset.UtcNow }, out _) is { } stopped)
                            this.ShowSharedState(stopped);

                        this.Say(CollectionWindow.SharingStopped);
                        host.Changed();
                        break;

                    case CollectionClient.Outcome.GoneFromServer:
                        this.ForgetPage();
                        break;

                    default:
                        this.Say(result.Message);
                        break;
                }
            }
            finally
            {
                this.SetBusy(false);
            }
        }

        // ###########################################################################################
        // "ATTACH TO WORKLOG" (owner decision, 2026-11-27: "A copy of its captures"): what the window holds kept first, then the
        // worklog picked and the captures copied in by CRT (Main.AttachCollectionToWorklogAsync), what came of it said. One at a
        // time, and never while a share is on its way.
        // ###########################################################################################
        private async void OnAttachToWorklogClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.AttachToWorklogAsync, this.Say);

        internal async Task AttachToWorklogAsync()
        {
            if (this.thisBusy || this.thisHost?.AttachToWorklog is not { } attach || !this.SaveEdits() || this.Collection is not { } collection)
                return;

            this.SetBusy(true);

            try
            {
                if (await attach(collection, this) is { } said)
                    this.Say(said);
            }
            finally
            {
                this.SetBusy(false);
            }
        }

        private void OnOpenPageClick(object? sender, RoutedEventArgs e)
        {
            if (this.Collection?.PageUrl is { } url)
                this.thisHost?.OpenLink(url);
        }

        // ###########################################################################################
        // "COPY LINK" (owner decision, 2026-11-28: "OK to your proposal for the forum button"): forums differ only in how they
        // write a link, so a FORMAT is picked, never a forum - CRT.Data's CollectionLinkFormats, as the page's "Link" panel
        // offers them. A click copies in the one picked last (the host's LinkFormat, UserSettings), which the button names; the
        // arrow lists all three, each with where it works, and the one picked there is copied and remembered.
        // ###########################################################################################
        private CollectionLinkFormat thisLinkFormat = CollectionLinkFormat.BbCode;

        private readonly List<MenuItem> thisLinkFormatItems = [];

        private void BuildCopyLinkMenu()
        {
            foreach (CollectionLinkFormat format in CollectionLinkFormats.All)
            {
                // Wrapped, never cut: a menu is no wider than Fluent's flyout (456), and an item measures its header with no
                // limit, so the words carry their own.
                var words = new TextBlock { TextWrapping = Avalonia.Media.TextWrapping.Wrap, MaxWidth = 360 };
                BoldRuns.ShowMarked(words, $"<b>{CollectionLinkFormats.NameOf(format)}</b> - {CollectionLinkFormats.WhereOf(format)}");

                var item = new MenuItem
                {
                    Header = words,
                    Tag = format,
                    ToggleType = MenuItemToggleType.Radio,
                    GroupName = "CollectionLinkFormat"
                };

                item.Click += async (_, _) => await HandlerGuard.RunAsync(() => this.CopyAsync(format, remember: true), this.Say);

                this.thisLinkFormatItems.Add(item);
                ((MenuFlyout)this.CopyLinkButton.Flyout!).Items.Add(item);
            }
        }

        // The format a click copies in, named on the button and checked in its list.
        private void ShowLinkFormat()
        {
            this.CopyLinkButton.Content = CollectionWindow.CopyLinkLabel(this.thisLinkFormat);

            foreach (MenuItem item in this.thisLinkFormatItems)
                item.IsChecked = item.Tag is CollectionLinkFormat format && format == this.thisLinkFormat;
        }

        // "Copy link (BBCode)", "Copy link (Markdown)", "Copy link (plain)".
        internal static string CopyLinkLabel(CollectionLinkFormat format) =>
            format == CollectionLinkFormat.Plain ? "Copy link (plain)" : $"Copy link ({CollectionLinkFormats.NameOf(format)})";

        internal CollectionLinkFormat LinkFormatShown => this.thisLinkFormat;

        internal IReadOnlyList<MenuItem> LinkFormatItemsForTests => this.thisLinkFormatItems;

        private async void OnCopyLinkClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.CopyAsync(this.thisLinkFormat), this.Say);

        internal async Task CopyAsync(CollectionLinkFormat format, bool remember = false)
        {
            if (this.thisHost is not { } host || this.Collection is not { } collection || collection.PageUrl is null)
                return;

            if (remember && format != this.thisLinkFormat)
            {
                this.thisLinkFormat = format;
                host.RememberLinkFormat?.Invoke(format);
            }

            this.ShowLinkFormat();

            await host.CopyText(CollectionSharing.ForumText(collection, format));
            this.Say(CollectionWindow.CopiedLine(format));
        }

        internal static string CopiedLine(CollectionLinkFormat format) => format switch
        {
            CollectionLinkFormat.Markdown => "Copied as Markdown - paste it into the post (Reddit, Discord, GitHub and Discourse show it as a link).",
            CollectionLinkFormat.Plain => "Copied - the link after a line saying what it is, to paste anywhere.",
            _ => "Copied as BBCode - paste it into the forum post (Lemon64, Forum64 and most forums show it as a link)."
        };

        // ###########################################################################################
        // "DELETE COLLECTION": asked first, in place of the buttons - then its page off the server (signed in), and it off this
        // computer. A page the server cannot take down now is said, and the collection kept, so it can be tried again.
        // ###########################################################################################
        private void OnDeleteClick(object? sender, RoutedEventArgs e) => this.AskToDelete();

        // Also from the "Oscilloscope" tab's "Collections" list, which opens the window to ask (Main.OpenCollectionWindowToDelete)
        // - not while a share or another wait is on: the window shows that, and asks nothing over it.
        internal void AskToDelete()
        {
            if (this.thisBusy || this.Collection is not { } collection)
                return;

            this.thisLocalDeleteAsked = false;
            this.DeleteQuestionText.Text = collection.IsShared
                ? $"Delete [{CollectionSharing.NameOf(collection)}] from this computer, and its page from the server?"
                : $"Delete [{CollectionSharing.NameOf(collection)}] from this computer?";
            this.ShowDeleteQuestion(true);
        }

        private void OnKeepClick(object? sender, RoutedEventArgs e)
        {
            this.thisLocalDeleteAsked = false;
            this.ShowDeleteQuestion(false);
        }

        private void ShowDeleteQuestion(bool asking)
        {
            this.DeleteQuestion.IsVisible = asking;
            this.ButtonsPanel.IsVisible = !asking;
            this.LeftButtonsPanel.IsVisible = !asking;
        }

        private async void OnConfirmDeleteClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.DeleteAsync, this.Say);

        internal async Task DeleteAsync()
        {
            this.ShowDeleteQuestion(false);

            if (this.thisBusy || this.thisHost is not { } host || this.Collection is not { } collection)
                return;

            if (collection.ServerId is { } serverId)
            {
                if (host.SignedIn() is null)
                {
                    this.Say("Sign in first - the account that shared it takes its page off the server.");
                    await host.SignIn();
                }

                // Asked again once signed in, as a comment's action is.
                if (this.thisBusy || host.SignedIn() is not { } session)
                    return;

                // ###########################################################################################
                // ANOTHER ACCOUNT'S: its "not found" says nothing of the page, so the server is not asked - and the collection may
                // be deleted FROM THIS COMPUTER ONLY, asked first, its page left up to its date (code review, 2026-11-27: with
                // the account gone - the test instance's data reset deletes them - it could never be deleted at all).
                // ###########################################################################################
                if (CollectionSharing.OtherAccountLine(collection, session.AccountId) is { } otherAccount)
                {
                    if (!this.thisLocalDeleteAsked)
                    {
                        this.thisLocalDeleteAsked = true;
                        this.DeleteQuestionText.Text = CollectionWindow.LocalDeleteQuestion(otherAccount, collection);
                        this.ShowDeleteQuestion(true);
                        return;
                    }
                }
                else
                {
                    if (!await this.DeletePageAsync(host, session, serverId))
                        return;
                }
            }

            if (!host.Store.Delete(collection.LocalId))
            {
                this.Say("Could not delete it from this computer - see \"Errors and warnings\" on the \"Configuration\" tab.");
                return;
            }

            host.Changed();
            this.Close();
        }

        // Asked once a delete meets a collection of another account - the next "Delete collection" deletes it here only.
        private bool thisLocalDeleteAsked;

        internal static string LocalDeleteQuestion(string otherAccount, LocalCollection collection) =>
            $"{otherAccount} Or delete it from this computer only - its page then stays up until " +
            $"{CollectionRules.DateText(collection.DeleteAtUtc ?? collection.CreatedUtc)}, and only that account can take it down.";

        // The page taken off the server - true when it is gone (or was), so the collection goes from this computer too.
        private async Task<bool> DeletePageAsync(CollectionWindowHost host, ReviewSession session, string serverId)
        {
            this.SetBusy(true);

            CollectionClient.Result? result;

            try
            {
                WaitResult<CollectionClient.Result> waited =
                    await BusyOverlay.RunAsync(this, "Taking its page off the server", _ => host.Client.DeleteAsync(serverId, session.BearerToken));

                // Not answered in time - it may be gone or not: ASKED AGAIN, as a timeout is never "it failed" (code review,
                // 2026-11-27). Deleting what is deleted already is done, so the second ask says where it stands.
                if (waited.IsTimedOut)
                {
                    this.Say("The server has not answered yet - asking it again whether the page is gone.");
                    waited = await BusyOverlay.RunAsync(this, "Asking the server whether the page is gone", _ => host.Client.DeleteAsync(serverId, session.BearerToken));
                }

                result = waited.IsTimedOut ? null : waited.Value;
            }
            finally
            {
                this.SetBusy(false);
            }

            if (result is null)
            {
                this.Say("The server still has not answered - the collection is kept here. Press \"Delete collection\" again in a moment.");
                return false;
            }

            if (result.Kind != CollectionClient.Outcome.Deleted)
            {
                this.Say(result.Message);
                return false;
            }

            return true;
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();

        // ###########################################################################################
        // *** NOT CLOSED BY HAND WHILE IT WAITS (code review, 2026-11-27). *** The overlay covers no title bar: closed during a
        // share, nothing guarded the collection any more - a window opened again shared it a second time, and a popup's "Take
        // out" deleted a capture the share had still to send. CRT itself quitting closes it whatever.
        // ###########################################################################################
        internal const string WaitBeforeClosing = "Wait for the server to answer, then close the window.";

        protected override void OnClosing(WindowClosingEventArgs e)
        {
            if (this.thisBusy && e.CloseReason == WindowCloseReason.WindowClosing && !e.IsProgrammatic)
            {
                e.Cancel = true;
                this.Say(CollectionWindow.WaitBeforeClosing);
                return;
            }

            this.SaveEdits();

            // Opened there again, its size too (owner request, 2026-11-28).
            this.thisHost?.RememberPlace?.Invoke(this.thisPlace.Current());

            base.OnClosing(e);
        }

        // ###########################################################################################
        // ESC CLOSES IT (owner request, 2026-11-28: "ESC should also be able to close the window") - or, a delete being asked, keeps
        // the collection, as "Keep it" does. An open drop-down list takes its Esc first and closes itself.
        // ###########################################################################################
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);

            if (e.Handled || e.Key != Key.Escape || e.KeyModifiers != KeyModifiers.None)
                return;

            e.Handled = true;

            if (this.DeleteQuestion.IsVisible)
            {
                this.thisLocalDeleteAsked = false;
                this.ShowDeleteQuestion(false);
                return;
            }

            this.Close();
        }

        private static string? Blank(string? text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();

        internal IReadOnlyList<CollectionCaptureRow> RowsForTests => this.thisRows;
    }
}
