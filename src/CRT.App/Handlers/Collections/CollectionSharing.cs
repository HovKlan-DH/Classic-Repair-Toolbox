using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.Collections
{
    // ###########################################################################################
    // A COLLECTION AS IT IS SHARED (2026-11-26) - pure: what CRT sends the server (CollectionContract's request), the order its
    // captures are shown in, whether it changed since it was shared, and what "Copy for forum" puts on the clipboard.
    // ###########################################################################################
    public static class CollectionSharing
    {
        public enum ForumFormat
        {
            // Lemon64, Forum64 and most forums.
            BbCode,

            // Reddit and the like.
            Markdown
        }

        // ###########################################################################################
        // THE PAGE'S ORDER: by component (as every list of labels CRT shows - U2 before U10), then by pin, then as taken.
        // ###########################################################################################
        public static IReadOnlyList<LocalCapture> InPageOrder(IEnumerable<LocalCapture> captures) =>
            captures
                .OrderBy(capture => capture.ComponentLabel, NaturalLabelComparer.Instance)
                .ThenBy(capture => capture.Pin ?? string.Empty, NaturalLabelComparer.Instance)
                .ThenBy(capture => capture.CapturedUtc)
                .ToList();

        // When the server is to delete it: the date kept, else the "Delete in" chosen counted from `now`.
        public static DateTimeOffset DeleteAtFor(LocalCollection collection, DateTimeOffset now) =>
            collection.DeleteAtUtc ?? CollectionRules.DeleteAt(now, collection.Lifetime);

        // ###########################################################################################
        // The request that shares the collection as it is now - every capture, in the page's order.
        // ###########################################################################################
        public static CollectionContentRequest RequestOf(LocalCollection collection, DateTimeOffset deleteAtUtc) =>
            new(
                collection.HardwareName,
                collection.BoardName,
                collection.Region,
                collection.Title,
                collection.Description,
                collection.ForumLinks,
                deleteAtUtc,
                CollectionSharing.InPageOrder(collection.Captures).Select(CollectionSharing.EntryOf).ToList());

        public static CollectionImageEntry EntryOf(LocalCapture capture) =>
            new(
                capture.Sha256,
                capture.ComponentLabel,
                capture.Pin,
                capture.SignalName,
                capture.ExpectedReading,
                capture.TimePerDivision,
                capture.VoltsPerDivision,
                capture.TriggerLevel,
                CollectionSharing.HasBaseline(capture)
                    ? new CollectionBaseline(
                        capture.BaselineSha256,
                        capture.BaselineFileName,
                        capture.BaselineTimePerDivision,
                        capture.BaselineVoltsPerDivision,
                        capture.BaselineTriggerLevel)
                    : null,
                capture.Caption,
                capture.LooksWrong,
                capture.CapturedUtc);

        private static bool HasBaseline(LocalCapture capture) =>
            capture.BaselineSha256 is not null || capture.BaselineTimePerDivision is not null ||
            capture.BaselineVoltsPerDivision is not null || capture.BaselineTriggerLevel is not null;

        // ###########################################################################################
        // What was shared, as a checksum of the request - the same collection is the same checksum, so a change since it was
        // shared is told by comparing (HasChangesNotShared).
        // ###########################################################################################
        public static string ContentHash(CollectionContentRequest request) =>
            Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(request, ReviewApiContract.WireSettings))));

        // Shared, and changed since - "Share" sends the change.
        public static bool HasChangesNotShared(LocalCollection collection) =>
            collection.IsShared &&
            collection.DeleteAtUtc is { } deleteAt &&
            !string.Equals(
                CollectionSharing.ContentHash(CollectionSharing.RequestOf(collection, deleteAt)),
                collection.SharedContentHash,
                StringComparison.Ordinal);

        // A collection as a line of the component info window's drop-down list - "new comments" while its page has comments
        // newer than its owner has read (`latestCommentId`, from the server's list of the account's collections).
        public static string LineOf(LocalCollection collection, long latestCommentId = 0) =>
            $"{CollectionSharing.NameOf(collection)} ({collection.Captures.Count} capture{(collection.Captures.Count == 1 ? string.Empty : "s")}" +
            (CollectionSharing.HasNewComments(collection, latestCommentId) ? ", new comments)" : ")");

        public static bool HasNewComments(LocalCollection collection, long latestCommentId) =>
            collection.IsShared && latestCommentId > collection.SeenCommentId;

        public static string NameOf(LocalCollection collection) =>
            string.IsNullOrWhiteSpace(collection.Title) ? "Untitled collection" : collection.Title.Trim();

        // A new collection's title, until the owner types one: the component it was started from and the day.
        public static string DefaultTitle(string componentLabel, DateTimeOffset now) =>
            $"{componentLabel} captures, {CollectionRules.DateText(now)}";

        // ###########################################################################################
        // "COPY FOR FORUM": the link with a line saying what it is - and each capture marked "looks wrong" linked on its own, so
        // the post can point at exactly those. BBCode for Lemon64 and Forum64, Markdown for Reddit.
        // ###########################################################################################
        public static string ForumText(LocalCollection collection, ForumFormat format)
        {
            if (collection.PageUrl is not { } url)
                return string.Empty;

            string Link(string address, string text) => format == ForumFormat.BbCode ? $"[url={address}]{text}[/url]" : $"[{text}]({address})";

            IReadOnlyList<LocalCapture> captures = CollectionSharing.InPageOrder(collection.Captures);
            IReadOnlyList<string> anchors = CollectionRules.AnchorsOf(captures.Select(CollectionSharing.EntryOf).ToList());

            var text = new StringBuilder()
                .Append(Link(url, CollectionSharing.NameOf(collection)))
                .Append($" - {captures.Count} oscilloscope capture{(captures.Count == 1 ? string.Empty : "s")} of ")
                .Append($"{collection.HardwareName} {collection.BoardName}")
                .Append(collection.Region is { } region ? $" ({region})" : string.Empty)
                .Append(", shared with Classic Repair Toolbox");

            List<int> wrong = Enumerable.Range(0, captures.Count).Where(at => captures[at].LooksWrong).ToList();

            if (wrong.Count > 0)
            {
                text.Append("\n\nThese look wrong to me:");

                if (format == ForumFormat.BbCode)
                    text.Append("\n[list]");

                foreach (int at in wrong)
                {
                    LocalCapture capture = captures[at];
                    string name = CollectionRules.WhereOf(CollectionSharing.EntryOf(capture)) + (capture.SignalName is { } signal ? $" ({signal})" : string.Empty);

                    text.Append(format == ForumFormat.BbCode ? "\n[*] " : "\n- ").Append(Link($"{url}#{anchors[at]}", name));
                }

                if (format == ForumFormat.BbCode)
                    text.Append("\n[/list]");
            }

            return text.ToString();
        }
    }
}
