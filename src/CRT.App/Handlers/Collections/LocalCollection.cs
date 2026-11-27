using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using Handlers.DataHandling;

namespace Handlers.Collections
{
    // ###########################################################################################
    // AN OSCILLOSCOPE COLLECTION ON THIS COMPUTER (2026-11-26; CRT.Data's CollectionRules says what one is) - made from the
    // component info window's captures, kept here until it is shared, and after: the page on the server is made from it, so
    // a change here is shared again with "Share" (CollectionSharing). One board's, as every collection is.
    //
    // Kept by CollectionStore as "collection.json" in a folder of its own, beside its captures. A field a later CRT writes is
    // kept and written back (JsonExtensionData), as a device's and a workbook's are.
    // ###########################################################################################
    public sealed record LocalCollection
    {
        [JsonPropertyName("localId")] public string LocalId { get; init; } = string.Empty;

        [JsonPropertyName("hardwareName")] public string HardwareName { get; init; } = string.Empty;

        [JsonPropertyName("boardName")] public string BoardName { get; init; } = string.Empty;

        [JsonPropertyName("region")] public string? Region { get; init; }

        [JsonPropertyName("title")] public string? Title { get; init; }

        [JsonPropertyName("description")] public string? Description { get; init; }

        [JsonPropertyName("forumLinks")] public IReadOnlyList<string> ForumLinks { get; init; } = [];

        // The "Delete in" chosen (CollectionRules.Lifetimes) - counted from the day it is shared, or kept longer.
        [JsonPropertyName("lifetime")] public string Lifetime { get; init; } = CollectionRules.DefaultLifetimeKey;

        // When the server deletes it - set as it is shared, moved by "Keep it longer".
        [JsonPropertyName("deleteAtUtc")] public DateTimeOffset? DeleteAtUtc { get; init; }

        [JsonPropertyName("createdUtc")] public DateTimeOffset CreatedUtc { get; init; }

        // Once shared: the server's id for it, its page, when, and what was sent (CollectionSharing.ContentHash) - so a change
        // since is known.
        [JsonPropertyName("serverId")] public string? ServerId { get; init; }

        [JsonPropertyName("pageUrl")] public string? PageUrl { get; init; }

        [JsonPropertyName("sharedUtc")] public DateTimeOffset? SharedUtc { get; init; }

        [JsonPropertyName("sharedContentHash")] public string? SharedContentHash { get; init; }

        // The account it was shared with, and its name then (code review, 2026-11-27) - only that account shares, deletes or
        // answers it again (CollectionSharing.OtherAccountLine). Null for one shared before CRT kept it.
        [JsonPropertyName("sharedByAccountId")] public long? SharedByAccountId { get; init; }

        [JsonPropertyName("sharedByName")] public string? SharedByName { get; init; }

        // When its owner pressed "Stop sharing" (owner decision, 2026-11-27: "Offline, same link back") - its page offline, kept
        // on the server under its link, until "Share" brings it back. Null while it is shared, or never was.
        [JsonPropertyName("sharingStoppedUtc")] public DateTimeOffset? SharingStoppedUtc { get; init; }

        // The newest comment its owner has read in CRT (step 3) - comments newer than it are marked new.
        [JsonPropertyName("seenCommentId")] public long SeenCommentId { get; init; }

        [JsonPropertyName("captures")] public IReadOnlyList<LocalCapture> Captures { get; init; } = [];

        [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; init; }

        // It has a page on the server - shown, or offline with its sharing stopped.
        [JsonIgnore] public bool IsShared => this.ServerId is not null;

        // Its page is up for anybody with the link.
        [JsonIgnore] public bool IsOnline => this.IsShared && this.SharingStoppedUtc is null;
    }

    // ###########################################################################################
    // One capture: its PNG by checksum (beside the collection's file), the pin it is of, the oscilloscope's settings read back
    // as it was taken, and the board's baseline for that pin - as the page shows them.
    // ###########################################################################################
    public sealed record LocalCapture
    {
        [JsonPropertyName("sha256")] public string Sha256 { get; init; } = string.Empty;

        [JsonPropertyName("componentLabel")] public string ComponentLabel { get; init; } = string.Empty;

        [JsonPropertyName("pin")] public string? Pin { get; init; }

        [JsonPropertyName("signalName")] public string? SignalName { get; init; }

        [JsonPropertyName("expectedReading")] public string? ExpectedReading { get; init; }

        [JsonPropertyName("timePerDivision")] public string? TimePerDivision { get; init; }

        [JsonPropertyName("voltsPerDivision")] public string? VoltsPerDivision { get; init; }

        [JsonPropertyName("triggerLevel")] public string? TriggerLevel { get; init; }

        [JsonPropertyName("baselineSha256")] public string? BaselineSha256 { get; init; }

        [JsonPropertyName("baselineFileName")] public string? BaselineFileName { get; init; }

        [JsonPropertyName("baselineTimePerDivision")] public string? BaselineTimePerDivision { get; init; }

        [JsonPropertyName("baselineVoltsPerDivision")] public string? BaselineVoltsPerDivision { get; init; }

        [JsonPropertyName("baselineTriggerLevel")] public string? BaselineTriggerLevel { get; init; }

        [JsonPropertyName("caption")] public string? Caption { get; init; }

        [JsonPropertyName("capturedUtc")] public DateTimeOffset CapturedUtc { get; init; }

        [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; init; }
    }
}
