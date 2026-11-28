using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // OSCILLOSCOPE COLLECTIONS (owner requests, 2026-11-25: "a possibility for a user to capture one or more oscilloscope images
    // and make that a "collection". Then that collection can be shared to the internet on various forums like Lemon64, Forum64
    // or Reddit ... the goal would be for the user to get input from the community"). EVERY RULE BOTH ENDS KEEP: what a
    // collection's id looks like, when it is deleted, and what it may hold - CRT checks a collection before it shares it, and
    // the server again as it takes it (CollectionFlows), in these words.
    //
    //   - A collection is ONE BOARD'S, made from captures taken in CRT's component info window, kept on the computer until it
    //     is shared (CRT's own store) - shared only from CRT, by anybody, NO ACCOUNT (owner decision, 2026-11-28: "Anybody, no
    //     account"): its page is owned by a key kept with it on that computer (IsOwnerKey).
    //   - Its page is https://classic-repair-toolbox.dk/collections/<id> - anybody with the link sees it, nobody gets a list of
    //     them, no search engine is asked to keep it, and it names nobody (owner decision, 2026-11-25).
    //   - "Delete in" 1 day, 1 week, 1 month or 1 year (the default), counted from every share - never past a year ahead.
    // ###########################################################################################
    public static class CollectionRules
    {
        // ---------------------------------------------------------------------------------------
        // THE ID: 12 characters of lower-case letters and digits, none that looks like another (no 0, o, 1 or l) - it cannot be
        // guessed (32^12, 60 bits), and it is typed right from a screenshot.
        // ---------------------------------------------------------------------------------------
        public const string IdCharacters = "abcdefghijkmnpqrstuvwxyz23456789";

        public const int IdLength = 12;

        public static string NewId()
        {
            Span<char> id = stackalloc char[CollectionRules.IdLength];

            for (int i = 0; i < id.Length; i++)
                id[i] = CollectionRules.IdCharacters[RandomNumberGenerator.GetInt32(CollectionRules.IdCharacters.Length)];

            return new string(id);
        }

        public static bool IsId(string? id) =>
            id is { Length: CollectionRules.IdLength } && id.All(c => CollectionRules.IdCharacters.Contains(c, StringComparison.Ordinal));

        // ###########################################################################################
        // *** THE KEY A SHARED COLLECTION IS OWNED BY (owner decision, 2026-11-28: "Entirely remove the new "Account" possibility";
        // asked, "Anybody, no account"). *** 64 lower-case hex characters, 256 random bits, made by CRT as the collection is first
        // shared and kept with it on that computer alone - so only that computer changes or deletes its page. Sent with every
        // request (CollectionRoutes.OwnerKeyHeader); the server keeps nothing but its SHA-256, as it keeps a session's.
        // ###########################################################################################
        public const int OwnerKeyLength = 64;

        public static bool IsOwnerKey(string? key) =>
            key is { Length: CollectionRules.OwnerKeyLength } && key.All(c => c is (>= '0' and <= '9') or (>= 'a' and <= 'f'));

        public static string NewOwnerKey() => RandomNumberGenerator.GetHexString(CollectionRules.OwnerKeyLength, lowercase: true);

        // ---------------------------------------------------------------------------------------
        // "DELETE IN" (owner decisions, 2026-11-25: "no need to show the "2 years""; 2026-11-27: "The max years a collection can
        // be shared is 1 year, and not 5 years - however, every time the user shares it again, it resets its time"; 2026-11-28:
        // "1 day, 1 week, 1 month, 1 year"). Counted from the day it is SHARED - every share, so sharing it again keeps it another
        // day, week, month or year - never more than MaximumYearsAhead ahead. A "5y" or "6m" an earlier CRT kept reads as the
        // default (LifetimeOf).
        // ---------------------------------------------------------------------------------------
        public sealed record Lifetime(string Key, string Words, int Days, int Months);

        public static readonly IReadOnlyList<Lifetime> Lifetimes =
        [
            new("1d", "1 day", 1, 0),
            new("1w", "1 week", 7, 0),
            new("1m", "1 month", 0, 1),
            new("1y", "1 year", 0, 12)
        ];

        public const string DefaultLifetimeKey = "1y";

        public const int MaximumYearsAhead = 1;

        public static Lifetime LifetimeOf(string? key) =>
            CollectionRules.Lifetimes.FirstOrDefault(lifetime => string.Equals(lifetime.Key, key, StringComparison.Ordinal)) ??
            CollectionRules.Lifetimes.First(lifetime => lifetime.Key == CollectionRules.DefaultLifetimeKey);

        public static DateTimeOffset DeleteAt(DateTimeOffset from, string? lifetimeKey)
        {
            Lifetime lifetime = CollectionRules.LifetimeOf(lifetimeKey);

            return from.ToUniversalTime().AddMonths(lifetime.Months).AddDays(lifetime.Days);
        }

        // The latest a deletion may be set to from `now` - a day of room, so a date worked out on a computer whose clock is a
        // little ahead is not refused.
        public static DateTimeOffset LatestDeleteAt(DateTimeOffset now) =>
            now.ToUniversalTime().AddYears(CollectionRules.MaximumYearsAhead).AddDays(1);

        // ---------------------------------------------------------------------------------------
        // WHAT A COLLECTION MAY HOLD. A 40-pin chip is 40 captures; a whole board's worth of chips a few hundred.
        // ---------------------------------------------------------------------------------------
        public const int MaximumImages = 300;

        public const int MaximumTitleLength = 120;

        public const int MaximumDescriptionLength = 4000;

        public const int MaximumCaptionLength = 1000;

        public const int MaximumNameLength = 100;

        public const int MaximumSettingLength = 40;

        public const int MaximumForumLinks = 10;

        public const int MaximumLinkLength = 500;

        // The oscilloscopes the captures were taken with (server 16.1.0) - a few at most, each its vendor and model.
        public const int MaximumOscilloscopes = 10;

        // How many collections one INTERNET CONNECTION may share, and captures it may upload, in a day - an IPv4 address, or an
        // IPv6 /64 (no account says who shares, since server 17.0.0) - the server's (CollectionFlows), said here so CRT says the
        // same numbers.
        public const int MaximumCollectionsPerDay = 20;

        public const int MaximumUploadsPerDay = 1000;

        // ###########################################################################################
        // A link to a forum thread: an http or https address with a host, as typed - the page shows it as a link that carries
        // nothing of the page (rel="nofollow ugc noopener noreferrer").
        // ###########################################################################################
        public static bool IsForumLink(string? link)
        {
            string text = link?.Trim() ?? string.Empty;

            return text.Length is > 0 and <= CollectionRules.MaximumLinkLength &&
                   !text.Any(char.IsWhiteSpace) &&
                   Uri.TryCreate(text, UriKind.Absolute, out Uri? uri) &&
                   (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp) &&
                   uri.Host.Length > 0;
        }

        // ###########################################################################################
        // Everything wrong with a collection as sent - CRT before it shares, the server as it takes it. Images not uploaded yet
        // are the server's to say (CollectionFlows), as only it knows.
        // ###########################################################################################
        public static IReadOnlyList<string> ProblemsOf(CollectionContentRequest? content, DateTimeOffset now)
        {
            var problems = new List<string>();

            if (content is null)
            {
                problems.Add("Nothing was sent.");
                return problems;
            }

            CollectionRules.Required(problems, content.HardwareName, "the hardware", "The hardware's name");
            CollectionRules.Required(problems, content.BoardName, "the board", "The board's name");
            CollectionRules.AtMost(problems, content.Region, CollectionRules.MaximumNameLength, "The region");
            CollectionRules.AtMost(problems, content.Title, CollectionRules.MaximumTitleLength, "The title");
            CollectionRules.AtMost(problems, content.Description, CollectionRules.MaximumDescriptionLength, "The description");

            IReadOnlyList<string> links = content.ForumLinks ?? [];

            if (links.Count > CollectionRules.MaximumForumLinks)
                problems.Add($"A collection links to at most {CollectionRules.MaximumForumLinks} forum threads.");

            foreach (string link in links)
            {
                if (!CollectionRules.IsForumLink(link))
                    problems.Add($"[{CollectionRules.Shortened(link)}] is no web address - a forum thread's link starts with https://.");
            }

            if (content.DeleteAtUtc <= now)
                problems.Add("The date it is deleted on has passed - choose how long to keep it again.");
            else if (content.DeleteAtUtc > CollectionRules.LatestDeleteAt(now))
                problems.Add("A collection is kept at most a year ahead - share it again to keep it longer.");

            IReadOnlyList<CollectionImageEntry> images = content.Images ?? [];

            if (images.Count == 0)
                problems.Add("A collection holds at least one capture.");
            else if (images.Count > CollectionRules.MaximumImages)
                problems.Add($"A collection holds at most {CollectionRules.MaximumImages} captures - this one has {images.Count}.");

            var seen = new HashSet<string>(StringComparer.Ordinal);

            foreach (CollectionImageEntry? image in images)
            {
                // "images": [null] in the body - said, never a crash the server answers 500 for (code review, 2026-11-27).
                if (image is null)
                {
                    problems.Add("A capture is empty.");
                    continue;
                }

                string where = CollectionRules.WhereOf(image);

                if (!DataTreeLayout.IsSha256(image.Sha256))
                    problems.Add($"{where}: the capture's checksum is not one.");
                else if (!seen.Add(image.Sha256))
                    problems.Add($"{where}: the same capture is in the collection twice.");

                if (string.IsNullOrWhiteSpace(image.ComponentLabel))
                    problems.Add("A capture names no component.");

                CollectionRules.AtMost(problems, image.ComponentLabel, CollectionRules.MaximumNameLength, $"{where}: the component");
                CollectionRules.AtMost(problems, image.Pin, CollectionRules.MaximumNameLength, $"{where}: the pin");
                CollectionRules.AtMost(problems, image.SignalName, CollectionRules.MaximumNameLength, $"{where}: the signal name");
                CollectionRules.AtMost(problems, image.ExpectedReading, CollectionRules.MaximumNameLength, $"{where}: the expected reading");
                CollectionRules.AtMost(problems, image.TimePerDivision, CollectionRules.MaximumSettingLength, $"{where}: T/DIV");
                CollectionRules.AtMost(problems, image.VoltsPerDivision, CollectionRules.MaximumSettingLength, $"{where}: V/DIV");
                CollectionRules.AtMost(problems, image.TriggerLevel, CollectionRules.MaximumSettingLength, $"{where}: the trigger level");
                CollectionRules.AtMost(problems, image.Caption, CollectionRules.MaximumCaptionLength, $"{where}: the note");

                if (image.PinoutSha256 is not null && !DataTreeLayout.IsSha256(image.PinoutSha256))
                    problems.Add($"{where}: the pinout drawing's checksum is not one.");
            }

            CollectionRules.ComponentProblems(problems, content.Components ?? []);

            IReadOnlyList<string?> oscilloscopes = content.Oscilloscopes ?? [];

            if (oscilloscopes.Count > CollectionRules.MaximumOscilloscopes)
                problems.Add($"A collection names at most {CollectionRules.MaximumOscilloscopes} oscilloscopes.");

            foreach (string? oscilloscope in oscilloscopes)
            {
                if (string.IsNullOrWhiteSpace(oscilloscope))
                    problems.Add("An oscilloscope's name is empty.");
                else
                    CollectionRules.AtMost(problems, oscilloscope, CollectionRules.MaximumNameLength, "An oscilloscope's name");
            }

            return problems;
        }

        // ---------------------------------------------------------------------------------------
        // THE COMPONENTS' FACTS (server 16.0.0): one per component captured at most, each as CRT's component info window shows it.
        // ---------------------------------------------------------------------------------------
        public const int MaximumComponentTextLength = 300;

        public const int MaximumPurposeLength = 4000;

        public const int MaximumPinsPerComponent = 300;

        public const int MaximumReadingLength = 1000;

        private static void ComponentProblems(List<string> problems, IReadOnlyList<CollectionComponent?> components)
        {
            if (components.Count > CollectionRules.MaximumImages)
                problems.Add($"A collection tells of at most {CollectionRules.MaximumImages} components.");

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (CollectionComponent? component in components)
            {
                if (component is null || string.IsNullOrWhiteSpace(component.Label))
                {
                    problems.Add("A component's facts name no component.");
                    continue;
                }

                string label = component.Label.Trim();

                if (!seen.Add(label))
                    problems.Add($"{CollectionRules.Shortened(label)}: its facts are in the collection twice.");

                CollectionRules.AtMost(problems, label, CollectionRules.MaximumNameLength, "A component's label");
                CollectionRules.AtMost(problems, component.Title, CollectionRules.MaximumComponentTextLength, $"{label}: the title");
                CollectionRules.AtMost(problems, component.CategoryLine, CollectionRules.MaximumComponentTextLength, $"{label}: the category");
                CollectionRules.AtMost(problems, component.Description, CollectionRules.MaximumComponentTextLength, $"{label}: the description");
                CollectionRules.AtMost(problems, component.Purpose, CollectionRules.MaximumPurposeLength, $"{label}: what it is there for");

                IReadOnlyList<CollectionPin?> pins = component.Pins ?? [];

                if (pins.Count > CollectionRules.MaximumPinsPerComponent)
                    problems.Add($"{label}: a component lists at most {CollectionRules.MaximumPinsPerComponent} pins.");

                foreach (CollectionPin? pin in pins)
                {
                    if (pin is null || string.IsNullOrWhiteSpace(pin.Pin))
                    {
                        problems.Add($"{label}: a pin is empty.");
                        continue;
                    }

                    CollectionRules.AtMost(problems, pin.Pin, CollectionRules.MaximumSettingLength, $"{label}: a pin");
                    CollectionRules.AtMost(problems, pin.Name, CollectionRules.MaximumNameLength, $"{label}, pin {pin.Pin}: the name");
                    CollectionRules.AtMost(problems, pin.Direction, CollectionRules.MaximumSettingLength, $"{label}, pin {pin.Pin}: the direction");
                    CollectionRules.AtMost(problems, pin.Net, CollectionRules.MaximumNameLength, $"{label}, pin {pin.Pin}: the net");
                    CollectionRules.AtMost(problems, pin.Reading, CollectionRules.MaximumReadingLength, $"{label}, pin {pin.Pin}: the expected reading");
                }
            }
        }

        // "U7, pin 12" - a capture as a person names it.
        public static string WhereOf(CollectionImageEntry image)
        {
            string label = string.IsNullOrWhiteSpace(image.ComponentLabel) ? "A capture" : image.ComponentLabel.Trim();

            return string.IsNullOrWhiteSpace(image.Pin) ? label : $"{label}, pin {image.Pin.Trim()}";
        }

        // ###########################################################################################
        // Each capture's anchor on the page, in the page's order: "U7-12" - letters, digits and "-" only, made unique with
        // "-2", "-3" when two share one. The page gives each capture its own (CollectionPages), and CRT's "Copy for forum"
        // links straight to one, so a forum reply can point at pin 12 - one rule for both.
        // ###########################################################################################
        public static IReadOnlyList<string> AnchorsOf(IReadOnlyList<CollectionImageEntry> images)
        {
            var anchors = new List<string>(images.Count);
            var taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "top" };

            foreach (CollectionImageEntry image in images)
            {
                string raw = string.IsNullOrWhiteSpace(image.Pin) ? image.ComponentLabel : $"{image.ComponentLabel}-{image.Pin}";
                string anchor = string.Concat((raw ?? string.Empty).Select(c => char.IsAsciiLetterOrDigit(c) ? c : '-')).Trim('-');

                // The page's own ids are "top", "locked" and those of the comments - "comment-<id>", "comments-<anchor>" and
                // "comment-on-<anchor>": a capture's never starts so (code review, 2026-11-27: a component labelled "COMMENT",
                // pin 12, took comment #12's place, and a mail's link to that comment jumped to the capture).
                if (anchor.Length == 0 || !char.IsAsciiLetter(anchor[0]) ||
                    anchor.StartsWith("comment", StringComparison.OrdinalIgnoreCase) || string.Equals(anchor, "locked", StringComparison.OrdinalIgnoreCase))
                {
                    anchor = "capture-" + anchor;
                }

                string unique = anchor;

                for (int n = 2; !taken.Add(unique); n++)
                    unique = $"{anchor}-{n}";

                anchors.Add(unique);
            }

            return anchors;
        }

        // The date a person reads, as CRT writes dates: "2026-November-26".
        public static string DateText(DateTimeOffset when) =>
            when.ToUniversalTime().ToString("yyyy-MMMM-d", CultureInfo.InvariantCulture);

        // A collection's page, under the site's address (the public API's address less its "/api").
        public static string PageUrl(string siteBaseUrl, string id) => $"{siteBaseUrl.TrimEnd('/')}{CollectionRoutes.PagePath}/{id}";

        private static void Required(List<string> problems, string? value, string what, string name)
        {
            if (string.IsNullOrWhiteSpace(value))
                problems.Add($"The collection does not say {what} it is about.");
            else
                CollectionRules.AtMost(problems, value, CollectionRules.MaximumNameLength, name);
        }

        private static void AtMost(List<string> problems, string? value, int limit, string what)
        {
            if (value is not null && value.Trim().Length > limit)
                problems.Add($"{what} is at most {limit} characters.");
        }

        private static string Shortened(string? text)
        {
            string value = (text ?? string.Empty).Trim();

            return value.Length <= 60 ? value : value[..57] + "...";
        }
    }
}
