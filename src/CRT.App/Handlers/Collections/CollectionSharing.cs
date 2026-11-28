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
    // captures are shown in, whether it changed since it was shared, and what "Copy link" puts on the clipboard.
    // ###########################################################################################
    public static class CollectionSharing
    {
        // ###########################################################################################
        // THE PAGE'S ORDER: by component (as every list of labels CRT shows - U2 before U10), then by pin, then as taken.
        // ###########################################################################################
        public static IReadOnlyList<LocalCapture> InPageOrder(IEnumerable<LocalCapture> captures) =>
            captures
                .OrderBy(capture => capture.ComponentLabel, NaturalLabelComparer.Instance)
                .ThenBy(capture => capture.Pin ?? string.Empty, NaturalLabelComparer.Instance)
                .ThenBy(capture => capture.CapturedUtc)
                .ToList();

        // A capture's settings as the collection window lists them - "This capture: T/DIV 50us | V/DIV 2V | Trigger level 1V".
        public static string SettingsLine(string whose, string? timePerDivision, string? voltsPerDivision, string? triggerLevel) =>
            $"{whose}: T/DIV {timePerDivision ?? "not read"} | V/DIV {voltsPerDivision ?? "not read"} | Trigger level {triggerLevel ?? "not read"}";

        // ###########################################################################################
        // When the server is to delete it, sent with EVERY share: the "Delete in" chosen, counted from `now` (owner decision,
        // 2026-11-27: "every time the user shares it again, it resets its time, so it will be for another year or another month
        // etc."). The date kept is only what the last share sent.
        // ###########################################################################################
        public static DateTimeOffset DeleteAtFor(LocalCollection collection, DateTimeOffset now) =>
            CollectionRules.DeleteAt(now, collection.Lifetime);

        // ###########################################################################################
        // The request that shares the collection as it is now - every capture, in the page's order, and what CRT's component
        // info window said of each component captured: the latest capture's word of it (a component captured again later, its
        // data changed meanwhile, is told as it was then).
        // ###########################################################################################
        public static CollectionContentRequest RequestOf(LocalCollection collection, DateTimeOffset deleteAtUtc)
        {
            ArgumentNullException.ThrowIfNull(collection);

            IReadOnlyList<LocalCapture> captures = CollectionSharing.InPageOrder(collection.Captures);

            List<CollectionComponent> components = captures
                .Where(capture => capture.Component is not null)
                .OrderBy(capture => capture.CapturedUtc)
                .GroupBy(capture => capture.ComponentLabel.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(group => group.Last().Component! with { Label = group.Key })
                .OrderBy(component => component.Label, NaturalLabelComparer.Instance)
                .ToList();

            // The oscilloscopes they were taken with, each once, as first taken (owner request, 2026-11-28) - a few at most.
            List<string> oscilloscopes = captures
                .OrderBy(capture => capture.CapturedUtc)
                .Select(capture => capture.Oscilloscope?.Trim())
                .OfType<string>()
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(CollectionRules.MaximumOscilloscopes)
                .ToList();

            return new(
                collection.HardwareName,
                collection.BoardName,
                collection.Region,
                collection.Title,
                collection.Description,
                collection.ForumLinks,
                deleteAtUtc,
                captures.Select(CollectionSharing.EntryOf).ToList(),
                components.Count == 0 ? null : components,
                oscilloscopes.Count == 0 ? null : oscilloscopes);
        }

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
                capture.Caption,
                capture.CapturedUtc,
                capture.PinoutSha256);

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

        // A collection as a line of the component info window's drop-down list.
        public static string LineOf(LocalCollection collection) =>
            $"{CollectionSharing.NameOf(collection)} ({collection.Captures.Count} capture{(collection.Captures.Count == 1 ? string.Empty : "s")})";

        public static string NameOf(LocalCollection collection) =>
            string.IsNullOrWhiteSpace(collection.Title) ? "Untitled collection" : collection.Title.Trim();

        // A new collection's title, until the owner types one: the component it was started from and the day.
        public static string DefaultTitle(string componentLabel, DateTimeOffset now) =>
            $"{componentLabel} captures, {CollectionRules.DateText(now)}";

        // ###########################################################################################
        // "COPY LINK": the link with a line saying what it is, in the format picked (owner decision, 2026-11-28 - by format, never
        // a button per forum; CRT.Data's CollectionLinkFormats, as the page's "Link" panel writes it): its title as the link's
        // words and the line after it, or - a plain link - the title and the line, then the address. ("Copy for forum" and "Copy
        // for Reddit" until then.)
        // ###########################################################################################
        public static string ForumText(LocalCollection collection, CollectionLinkFormat format)
        {
            ArgumentNullException.ThrowIfNull(collection);

            if (collection.PageUrl is not { } url)
                return string.Empty;

            string about = new StringBuilder()
                .Append($" - {collection.Captures.Count} oscilloscope capture{(collection.Captures.Count == 1 ? string.Empty : "s")} of ")
                .Append($"{collection.HardwareName} {collection.BoardName}")
                .Append(collection.Region is { } region ? $" ({region})" : string.Empty)
                .Append(", shared with Classic Repair Toolbox")
                .ToString();

            return format switch
            {
                CollectionLinkFormat.Plain => CollectionLinkFormats.Write(format, url, CollectionSharing.NameOf(collection) + about),

                // The line after the link is HTML too - a board's name with "&" in it stays a character.
                CollectionLinkFormat.Html => CollectionLinkFormats.Write(format, url, CollectionSharing.NameOf(collection)) + System.Net.WebUtility.HtmlEncode(about),
                _ => CollectionLinkFormats.Write(format, url, CollectionSharing.NameOf(collection)) + about
            };
        }

        // ###########################################################################################
        // A CAPTURE AS A WORKLOG'S PHOTO'S COMMENT (owner decision, 2026-11-27: a collection attached to a worklog is "A copy of its
        // captures"): where it was taken, the oscilloscope's settings, what it should read and the note typed on it - written as
        // a note with formatting (RichNote), the names plain and the values bold, as the capture window's large picture writes
        // settings (ComponentPaneWording.RunSettingsNamed). A character the format reads stays a character (RichNote.Serialize).
        // ###########################################################################################
        public static string WorklogCommentOf(LocalCapture capture)
        {
            ArgumentNullException.ThrowIfNull(capture);

            var text = new StringBuilder();
            var styles = new List<RichNoteStyle>();

            void Line(params (string Text, bool Bold)[] parts)
            {
                if (text.Length > 0)
                {
                    text.Append('\n');
                    styles.Add(RichNoteStyle.None);
                }

                foreach ((string words, bool bold) in parts)
                {
                    text.Append(words);
                    styles.AddRange(Enumerable.Repeat(bold ? RichNoteStyle.Bold : RichNoteStyle.None, words.Length));
                }
            }

            string where = CollectionRules.WhereOf(CollectionSharing.EntryOf(capture));

            Line((where, true), (capture.SignalName is { } signal ? $" - {signal}" : string.Empty, false));
            Line(CollectionSharing.SettingsRuns("This capture", capture.TimePerDivision, capture.VoltsPerDivision, capture.TriggerLevel));

            if (capture.ExpectedReading is { } reading)
                Line(("Expected reading: ", false), (reading, true));

            if (capture.Caption is { } caption)
            {
                foreach (string line in caption.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n'))
                    Line((line, false));
            }

            return RichNote.Serialize(new RichNoteText(text.ToString(), styles));
        }

        // "This capture: T/DIV **50us** | V/DIV **2V** | Trigger level **1V**" - a setting not read said so, plain.
        private static (string Text, bool Bold)[] SettingsRuns(string whose, string? timePerDivision, string? voltsPerDivision, string? triggerLevel) =>
        [
            ($"{whose}: T/DIV ", false), (timePerDivision ?? "not read", timePerDivision is not null),
            (" | V/DIV ", false), (voltsPerDivision ?? "not read", voltsPerDivision is not null),
            (" | Trigger level ", false), (triggerLevel ?? "not read", triggerLevel is not null)
        ];
    }
}
