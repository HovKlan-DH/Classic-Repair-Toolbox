using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.Collections
{
    // ###########################################################################################
    // ONE COLLECTION AS THE "Collections" LIST SHOWS IT (owner request, 2026-11-27: "In the new "Collections" tab we will show
    // all collections and have a management interface for them" - the "Oscilloscope" tab, beside its communication log;
    // TabOscilloscope.Collections.cs draws it). Every word a row says, worked out from the collection as kept - a record, so a
    // row changed in nothing is the same row (InPlaceList keeps it).
    //
    // `FirstCaptureSha256`: the capture its thumbnail shows - the first in the page's order, none with no capture.
    // `IsShared`: its page is up for anybody with the link - "Open shared page" is offered then alone (owner request, 2026-11-28:
    // "but only if it is shared").
    // ###########################################################################################
    public sealed record CollectionListItem(
        string LocalId,
        string Name,
        string Board,
        string Facts,
        string State,
        string? PageUrl,
        string? FirstCaptureSha256,
        bool IsShared = false);

    public static class CollectionListing
    {
        // ###########################################################################################
        // Every collection on this computer, the newest made first - each board's together reads less well than "the one I made
        // yesterday at the top", and the board is said on every row.
        // ###########################################################################################
        public static IReadOnlyList<CollectionListItem> ItemsOf(IEnumerable<LocalCollection> collections)
        {
            ArgumentNullException.ThrowIfNull(collections);

            return collections
                .OrderByDescending(collection => collection.CreatedUtc)
                .ThenBy(collection => CollectionSharing.NameOf(collection), StringComparer.CurrentCultureIgnoreCase)
                .Select(CollectionListing.ItemOf)
                .ToList();
        }

        public static CollectionListItem ItemOf(LocalCollection collection)
        {
            ArgumentNullException.ThrowIfNull(collection);

            return new CollectionListItem(
                collection.LocalId,
                CollectionSharing.NameOf(collection),
                CollectionListing.BoardOf(collection),
                $"{CollectionListing.CapturesOf(collection.Captures.Count)} | Made on {CollectionRules.DateText(collection.CreatedUtc)}",
                CollectionListing.StateOf(collection),
                collection.PageUrl,
                CollectionSharing.InPageOrder(collection.Captures).FirstOrDefault(capture => DataTreeLayout.IsSha256(capture.Sha256))?.Sha256,
                collection.IsOnline && collection.PageUrl is not null);
        }

        // The board it is of - "Commodore 64 - 250407 (PAL)" - as the collection's window says it too.
        public static string BoardOf(LocalCollection collection)
        {
            ArgumentNullException.ThrowIfNull(collection);

            return $"{collection.HardwareName} - {collection.BoardName}" + (collection.Region is { } region ? $" ({region})" : string.Empty);
        }

        public static string CapturesOf(int count) => count switch
        {
            0 => "No captures yet",
            1 => "1 capture",
            _ => $"{count} captures",
        };

        // ###########################################################################################
        // Where it stands, in one line - kept here only, its page offline, or shared, with whether it changed since and when its
        // page goes - no "Shared on" date (owner request, 2026-11-28, of the collection's window: "no need for that"). The window
        // says the same at length (CollectionWindow.SharedStateLine, DeleteDateLine).
        // ###########################################################################################
        public static string StateOf(LocalCollection collection)
        {
            ArgumentNullException.ThrowIfNull(collection);

            if (!collection.IsShared)
                return "Not shared - it is on this computer only.";

            if (!collection.IsOnline)
                return "Not shared at the moment - its page is offline.";

            string state = "Shared";

            if (CollectionSharing.HasChangesNotShared(collection))
                state += ", and changed since";

            if (collection.DeleteAtUtc is { } deleteAt)
                state += $" - its page is deleted on {CollectionRules.DateText(deleteAt)}";

            return state + ".";
        }

        // The words over the list - how many - and what it says with none.
        public static string HeadingOf(IReadOnlyCollection<CollectionListItem> items)
        {
            ArgumentNullException.ThrowIfNull(items);

            return items.Count == 1 ? "1 collection on this computer" : $"{items.Count} collections on this computer";
        }

        public const string Intro =
            "Your oscilloscope collections. Open one to name it, share it or delete it - your oscilloscope does not need to be " +
            "connected for that. A shared page is changed or deleted from this computer only.";

        public const string NoneYet =
            "No collections on this computer yet. A collection is made in a component's information window: with your " +
            "oscilloscope connected, pick \"New collection\" under \"Collect captures into\", and capture.";
    }
}
