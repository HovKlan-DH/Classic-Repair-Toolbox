using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // WHAT THE COLLECTION WINDOW NEEDS FROM CRT (Main.Collections.cs): the collections on this computer, the server, the
    // account CRT is signed in to and the way to sign in, opening a link, and the clipboard. A test hands its own.
    // ###########################################################################################
    internal sealed record CollectionWindowHost(
        CollectionStore Store,
        CollectionClient Client,
        Func<ReviewSession?> SignedIn,
        Func<Task> SignIn,
        Action<string> OpenLink,
        Func<string, Task> CopyText,
        Action Changed);

    // ###########################################################################################
    // ONE CAPTURE AS THE WINDOW LISTS IT - a view of the collection's capture; the note and the mark are what the window
    // writes back (CollectionWindow.EditedCollection).
    // ###########################################################################################
    public sealed class CollectionCaptureRow
    {
        public string Sha256 { get; init; } = string.Empty;

        public Bitmap? Thumbnail { get; set; }

        public string Heading { get; init; } = string.Empty;

        public string Settings { get; init; } = string.Empty;

        public string BaselineSettings { get; init; } = string.Empty;

        public bool HasBaseline { get; init; }

        public bool LooksWrong { get; set; }

        public string? Caption { get; set; }

        // Its comments on the page (step 3) - shown once they are read; the answer being typed.
        public bool ShowsComments { get; init; }

        public string CommentsHeading { get; init; } = string.Empty;

        public IReadOnlyList<CollectionCommentRow> Comments { get; init; } = [];

        public string? Answer { get; set; }
    }

    // ###########################################################################################
    // ONE COMMENT AS THE WINDOW LISTS IT (step 3): who and when - "You" for the owner's own - its text, and what the owner may do
    // with it: hide another's, show a hidden one again, delete their own. Editing is the page's.
    // ###########################################################################################
    public sealed class CollectionCommentRow
    {
        public long Id { get; init; }

        public string Heading { get; init; } = string.Empty;

        public string Body { get; init; } = string.Empty;

        public bool CanHide { get; init; }

        public bool CanShow { get; init; }

        public bool CanDelete { get; init; }

        // Newer than the owner had read as the window opened - and not their own.
        public bool IsNew { get; init; }

        internal static CollectionCommentRow Of(CollectionCommentAnswer comment, long seenBefore)
        {
            bool isNew = comment.Id > seenBefore && !comment.Mine;

            return new()
            {
                Id = comment.Id,
                Heading =
                    (comment.Mine ? "You" : comment.AuthorName) +
                    (comment.ByAdministrator && !comment.Mine ? $" ({CollectionComments.AdministratorBadge})" : string.Empty) +
                    $" - {CollectionComments.WhenText(comment.CreatedUtc)}" +
                    (comment.EditedUtc is null ? string.Empty : " - edited") +
                    (comment.Hidden ? " - hidden from the page" : string.Empty) +
                    (isNew ? " - NEW" : string.Empty),
                Body = comment.Body,
                CanHide = !comment.Mine && !comment.Hidden,
                CanShow = comment.Hidden,
                CanDelete = comment.Mine,
                IsNew = isNew
            };
        }
    }

    // ###########################################################################################
    // AN OSCILLOSCOPE COLLECTION'S WINDOW (2026-11-26; "Show collection" in the component info window). Here it is named and
    // described, its captures noted and marked "Looks wrong" or taken out, and it is SHARED:
    //
    //   - "Share" sends it as it is now (CollectionClient) - the first time a page is made, its address kept; later, the page
    //     follows the change. It needs an account (the account icon), and asks to sign in when there is none.
    //   - "Delete in" is how long the server keeps the page - counted from the day it is shared, or from the day a new choice
    //     is made for one shared (and then shared again). At most 5 years ahead.
    //   - "Open web page", and the link with a line about it for a forum ("Copy for forum", BBCode) or Reddit (Markdown),
    //     each capture marked "Looks wrong" linked on its own.
    //   - "Delete collection" takes it off this computer and its page off the server, after asking.
    //
    // Every edit is kept on this computer as its box is left or its mark changes (SaveEdits) - nothing goes to the server
    // until "Share".
    // ###########################################################################################
    public partial class CollectionWindow : Window
    {
        private CollectionWindowHost? thisHost;

        private string thisLocalId = string.Empty;

        private bool thisShowing;

        private bool thisBusy;

        private List<CollectionCaptureRow> thisRows = [];

        public CollectionWindow()
        {
            this.InitializeComponent();
        }

        internal void Initialize(CollectionWindowHost host, string localId)
        {
            this.thisHost = host;
            this.thisLocalId = localId;

            this.LifetimeComboBox.ItemsSource = CollectionRules.Lifetimes.Select(lifetime => lifetime.Words).ToList();
            this.Show(host.Store.Find(localId));

            // Its comments, when it is shared and CRT is signed in - read as it opens, never asking to sign in for them.
            _ = HandlerGuard.RunAsync(this.LoadCommentsAsync, this.Say);
        }

        internal string LocalId => this.thisLocalId;

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
                this.BoardText.Text = $"{collection.HardwareName} - {collection.BoardName}" + (collection.Region is { } region ? $" ({region})" : string.Empty);
                this.LifetimeComboBox.SelectedItem = CollectionRules.LifetimeOf(collection.Lifetime).Words;

                IReadOnlyList<LocalCapture> captures = CollectionSharing.InPageOrder(collection.Captures);

                this.CapturesHeading.Text = captures.Count == 0
                    ? "No captures yet - pick this collection under \"Collect captures into\" in a component's info window, and capture."
                    : $"{captures.Count} capture{(captures.Count == 1 ? string.Empty : "s")}";

                this.thisRows = captures.Select(this.RowOf).ToList();
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
            List<CollectionCommentRow> comments = (this.thisComments ?? [])
                .Where(comment => comment.ImageSha256 == capture.Sha256)
                .Select(comment => CollectionCommentRow.Of(comment, this.thisSeenBefore))
                .ToList();

            int fresh = comments.Count(comment => comment.IsNew);

            return new CollectionCaptureRow
            {
                Sha256 = capture.Sha256,
                Heading = CollectionRules.WhereOf(CollectionSharing.EntryOf(capture)) + (capture.SignalName is { } signal ? $" - {signal}" : string.Empty),
                Settings = CollectionWindow.SettingsLine("This capture", capture.TimePerDivision, capture.VoltsPerDivision, capture.TriggerLevel) +
                           (capture.ExpectedReading is { } reading ? $" | Expected reading: {reading}" : string.Empty),
                BaselineSettings = CollectionWindow.SettingsLine("Baseline", capture.BaselineTimePerDivision, capture.BaselineVoltsPerDivision, capture.BaselineTriggerLevel),
                HasBaseline = capture.BaselineTimePerDivision is not null || capture.BaselineVoltsPerDivision is not null || capture.BaselineTriggerLevel is not null,
                LooksWrong = capture.LooksWrong,
                Caption = capture.Caption,
                ShowsComments = this.thisComments is not null,
                CommentsHeading = comments.Count == 0
                    ? "No comments yet"
                    : $"{comments.Count} comment{(comments.Count == 1 ? string.Empty : "s")}" + (fresh > 0 ? $" - {fresh} new" : string.Empty),
                Comments = comments
            };
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

            CollectionClient.CommentsResult result = await host.Client.ListCommentsAsync(serverId, session.BearerToken);

            if (result.Kind != CollectionClient.Outcome.Done)
            {
                this.Say(result.Message);
                return;
            }

            if (this.thisComments is null)
                this.thisSeenBefore = collection.SeenCommentId;

            this.thisComments = result.Comments;

            long newest = result.Comments.Count == 0 ? 0 : result.Comments.Max(comment => comment.Id);

            if (newest > collection.SeenCommentId)
            {
                collection = collection with { SeenCommentId = newest };
                host.Store.Save(collection);
                host.Changed();
            }

            this.Show(collection);
        }

        private async void OnAnswerClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.AnswerAsync(sender), this.Say);

        // The answer typed under a capture, sent as a comment of the owner's.
        private Task AnswerAsync(object? sender) =>
            sender is Button { Tag: string sha256 } && this.thisRows.FirstOrDefault(row => row.Sha256 == sha256) is { } row
                ? this.CommentActionAsync((client, serverId, token) => client.AddCommentAsync(serverId, sha256, row.Answer ?? string.Empty, token))
                : Task.CompletedTask;

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

        // One thing done to the comments, then them read again - what went wrong said.
        internal async Task CommentActionAsync(Func<CollectionClient, string, string, Task<CollectionClient.Result>> act)
        {
            if (this.thisHost is not { } host || this.Collection?.ServerId is not { } serverId)
                return;

            if (host.SignedIn() is null)
            {
                this.Say(CollectionClient.SignInFirst);
                await host.SignIn();
            }

            if (host.SignedIn() is not { } session)
                return;

            CollectionClient.Result result = await act(host.Client, serverId, session.BearerToken);

            if (result.Kind != CollectionClient.Outcome.Done)
            {
                this.Say(result.Message);
                return;
            }

            this.Say(string.Empty);
            await this.LoadCommentsAsync();
        }

        internal static string SettingsLine(string whose, string? timePerDivision, string? voltsPerDivision, string? triggerLevel) =>
            $"{whose}: T/DIV {timePerDivision ?? "not read"} | V/DIV {voltsPerDivision ?? "not read"} | Trigger level {triggerLevel ?? "not read"}";

        // The thumbnails, decoded small off the UI thread - never disposed by hand (a picture still drawn takes the render thread
        // down; CLAUDE.md).
        private void LoadThumbnails(string localId)
        {
            if (this.thisHost is not { } host)
                return;

            foreach (CollectionCaptureRow row in this.thisRows)
            {
                string path = host.Store.CapturePath(localId, row.Sha256);

                _ = Task.Run(() =>
                {
                    try
                    {
                        using FileStream stream = File.OpenRead(path);
                        return Bitmap.DecodeToWidth(stream, 400);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException)
                    {
                        return null;
                    }
                }).ContinueWith(
                    decoded =>
                    {
                        if (decoded.Result is { } bitmap && this.thisRows.Contains(row))
                        {
                            row.Thumbnail = bitmap;

                            // Rows are plain objects - the list shown again takes the picture.
                            this.CapturesList.ItemsSource = null;
                            this.CapturesList.ItemsSource = this.thisRows;
                        }
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

            this.SharedStateText.Text = !collection.IsShared
                ? "Not shared yet - it is on this computer only. \"Share\" makes its page, which anybody with the link can see."
                : changed
                    ? $"Shared on {CollectionRules.DateText(collection.SharedUtc ?? collection.CreatedUtc)}, and changed since - \"Share\" sends the change."
                    : $"Shared on {CollectionRules.DateText(collection.SharedUtc ?? collection.CreatedUtc)}: {collection.PageUrl}";

            this.DeleteDateText.Text = collection.DeleteAtUtc is { } deleteAt
                ? $"Its page is deleted on {CollectionRules.DateText(deleteAt)}."
                : "Counted from the day it is shared.";

            this.ShareButton.IsEnabled = !this.thisBusy && collection.Captures.Count > 0 && (!collection.IsShared || changed);
            this.OpenPageButton.IsEnabled = !this.thisBusy && collection.PageUrl is not null;
            this.CopyForForumButton.IsEnabled = !this.thisBusy && collection.PageUrl is not null;
            this.CopyForRedditButton.IsEnabled = !this.thisBusy && collection.PageUrl is not null;
            this.DeleteButton.IsEnabled = !this.thisBusy;
        }

        internal void Say(string text) => this.StatusText.Text = text;

        // ###########################################################################################
        // THE COLLECTION AS THE WINDOW HOLDS IT - its boxes, the "Delete in" picked and each capture's note and mark - over the
        // collection as kept.
        // ###########################################################################################
        internal LocalCollection? EditedCollection()
        {
            if (this.Collection is not { } collection)
                return null;

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
                        ? capture with { LooksWrong = row.LooksWrong, Caption = CollectionWindow.Blank(row.Caption) }
                        : capture)
                    .ToList()
            };
        }

        // Kept on this computer - true when it was (or nothing changed).
        internal bool SaveEdits()
        {
            if (this.thisShowing || this.thisHost is not { } host || this.EditedCollection() is not { } edited)
                return true;

            // The records hold lists, compared by reference - the same content is the same JSON.
            if (System.Text.Json.JsonSerializer.Serialize(edited) == System.Text.Json.JsonSerializer.Serialize(this.Collection))
                return true;

            if (!host.Store.Save(edited))
            {
                this.Say("Could not keep the change on this computer - see \"Errors and warnings\" on the \"Configuration\" tab.");
                return false;
            }

            this.Title = $"Oscilloscope collection: {CollectionSharing.NameOf(edited)}";
            this.ShowSharedState(edited);
            host.Changed();

            return true;
        }

        private void OnEditLeft(object? sender, RoutedEventArgs e) => this.SaveEdits();

        private void OnCaptureMarkChanged(object? sender, RoutedEventArgs e) => this.SaveEdits();

        // ###########################################################################################
        // "Delete in": for a collection shared, a new choice counts from today - and "Share" sends it.
        // ###########################################################################################
        private void OnLifetimePicked(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisShowing || this.thisHost is not { } host || this.EditedCollection() is not { } edited)
                return;

            if (edited.IsShared)
                edited = edited with { DeleteAtUtc = CollectionRules.DeleteAt(DateTimeOffset.UtcNow, edited.Lifetime) };

            if (host.Store.Save(edited))
            {
                this.ShowSharedState(edited);
                host.Changed();
            }
        }

        private void OnTakeOutClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisHost is not { } host || sender is not Button { Tag: string sha256 })
                return;

            this.SaveEdits();

            if (host.Store.RemoveCapture(this.thisLocalId, sha256) is { } changed)
            {
                this.Show(changed);
                host.Changed();
            }
        }

        // ###########################################################################################
        // "SHARE": signed in first, then the collection as it is now, and what came of it kept and said.
        // ###########################################################################################
        private async void OnShareClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.ShareAsync, this.Say);

        internal async Task ShareAsync()
        {
            if (this.thisHost is not { } host || !this.SaveEdits() || this.Collection is not { } collection)
                return;

            if (host.SignedIn() is null)
            {
                this.Say(CollectionClient.SignInFirst);
                await host.SignIn();
            }

            if (host.SignedIn() is not { } session)
                return;

            this.SetBusy(true);

            try
            {
                var progress = new Progress<string>(this.Say);

                CollectionClient.Result result = await host.Client.ShareAsync(
                    collection,
                    session.BearerToken,
                    sha256 => host.Store.ReadCapture(collection.LocalId, sha256),
                    DateTimeOffset.UtcNow,
                    progress);

                this.AfterShare(collection, result);
            }
            finally
            {
                this.SetBusy(false);
            }
        }

        private void AfterShare(LocalCollection collection, CollectionClient.Result result)
        {
            if (this.thisHost is not { } host)
                return;

            switch (result.Kind)
            {
                case CollectionClient.Outcome.Shared when result.Answer is { } answer:
                    LocalCollection shared = collection with
                    {
                        ServerId = answer.Id,
                        PageUrl = answer.Url,
                        DeleteAtUtc = result.SentDeleteAtUtc ?? answer.DeleteAtUtc,
                        SharedUtc = DateTimeOffset.UtcNow,
                        SharedContentHash = result.SharedContentHash
                    };

                    host.Store.Save(shared);
                    this.Show(shared);
                    this.Say($"Shared - anybody with the link sees it: {answer.Url}. \"Copy for forum\" puts the link on the clipboard, with a line about it.");
                    break;

                case CollectionClient.Outcome.GoneFromServer:
                    LocalCollection unshared = collection with { ServerId = null, PageUrl = null, DeleteAtUtc = null, SharedUtc = null, SharedContentHash = null };

                    host.Store.Save(unshared);
                    this.Show(unshared);
                    this.Say("Its page is no longer on the server - its time was up, or it was taken down. \"Share\" makes a new page.");
                    break;

                default:
                    this.Say(result.Message);
                    break;
            }

            host.Changed();
        }

        private void SetBusy(bool busy)
        {
            this.thisBusy = busy;

            if (this.Collection is { } collection)
                this.ShowSharedState(collection);
        }

        private void OnOpenPageClick(object? sender, RoutedEventArgs e)
        {
            if (this.Collection?.PageUrl is { } url)
                this.thisHost?.OpenLink(url);
        }

        private async void OnCopyForForumClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.CopyAsync(CollectionSharing.ForumFormat.BbCode), this.Say);

        private async void OnCopyForRedditClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(() => this.CopyAsync(CollectionSharing.ForumFormat.Markdown), this.Say);

        internal async Task CopyAsync(CollectionSharing.ForumFormat format)
        {
            if (this.thisHost is not { } host || this.Collection is not { } collection || collection.PageUrl is null)
                return;

            await host.CopyText(CollectionSharing.ForumText(collection, format));
            this.Say(format == CollectionSharing.ForumFormat.BbCode
                ? "Copied - paste it into the forum post (Lemon64, Forum64 and most forums read it as a link)."
                : "Copied - paste it into the Reddit post.");
        }

        // ###########################################################################################
        // "DELETE COLLECTION": asked first, in place of the buttons - then its page off the server (signed in), and it off this
        // computer. A page the server cannot take down now is said, and the collection kept, so it can be tried again.
        // ###########################################################################################
        private void OnDeleteClick(object? sender, RoutedEventArgs e)
        {
            if (this.Collection is not { } collection)
                return;

            this.DeleteQuestionText.Text = collection.IsShared
                ? $"Delete [{CollectionSharing.NameOf(collection)}] from this computer, and its page from the server?"
                : $"Delete [{CollectionSharing.NameOf(collection)}] from this computer?";
            this.ShowDeleteQuestion(true);
        }

        private void OnKeepClick(object? sender, RoutedEventArgs e) => this.ShowDeleteQuestion(false);

        private void ShowDeleteQuestion(bool asking)
        {
            this.DeleteQuestion.IsVisible = asking;
            this.ButtonsPanel.IsVisible = !asking;
            this.DeleteButton.IsVisible = !asking;
        }

        private async void OnConfirmDeleteClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.DeleteAsync, this.Say);

        internal async Task DeleteAsync()
        {
            this.ShowDeleteQuestion(false);

            if (this.thisHost is not { } host || this.Collection is not { } collection)
                return;

            if (collection.ServerId is { } serverId)
            {
                if (host.SignedIn() is null)
                {
                    this.Say("Sign in first - the account that shared it takes its page off the server.");
                    await host.SignIn();
                }

                if (host.SignedIn() is not { } session)
                    return;

                this.SetBusy(true);

                CollectionClient.Result result;

                try
                {
                    result = await host.Client.DeleteAsync(serverId, session.BearerToken);
                }
                finally
                {
                    this.SetBusy(false);
                }

                if (result.Kind != CollectionClient.Outcome.Deleted)
                {
                    this.Say(result.Message);
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

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();

        protected override void OnClosing(WindowClosingEventArgs e)
        {
            this.SaveEdits();
            base.OnClosing(e);
        }

        private static string? Blank(string? text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();

        internal IReadOnlyList<CollectionCaptureRow> RowsForTests => this.thisRows;
    }
}
