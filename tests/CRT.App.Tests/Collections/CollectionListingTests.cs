using Handlers.Collections;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// The "Oscilloscope" tab's "Collections" list (CollectionListing): every collection, the newest made first, and what each
// card says - its board, its captures, where it stands, and comments not read.
// ###########################################################################################
public sealed class CollectionListingTests
{
    private static readonly DateTimeOffset Made = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private static LocalCapture Capture(char sha, string label, string? pin) =>
        new() { Sha256 = new string(sha, 64), ComponentLabel = label, Pin = pin, CapturedUtc = Made };

    private static LocalCollection Collection(string title, DateTimeOffset made, params LocalCapture[] captures) =>
        new()
        {
            LocalId = new string((char)('0' + (title.Length % 10)), 32),
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Region = "PAL",
            Title = title,
            CreatedUtc = made,
            Captures = captures
        };

    // Shared as it is now - its page up, nothing changed since.
    private static LocalCollection Shared(LocalCollection collection, DateTimeOffset deleteAt, long seenCommentId = 0) =>
        collection with
        {
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p",
            SharedUtc = Made.AddDays(1),
            DeleteAtUtc = deleteAt,
            SharedContentHash = CollectionSharing.ContentHash(CollectionSharing.RequestOf(collection, deleteAt)),
            SeenCommentId = seenCommentId
        };

    [Fact]
    public void The_newest_made_is_listed_first_whatever_its_board()
    {
        IReadOnlyList<CollectionListItem> items = CollectionListing.ItemsOf(
        [
            Collection("Older", Made),
            Collection("Newest", Made.AddDays(2)) with { HardwareName = "Commodore 128", BoardName = "310378" },
            Collection("Middle", Made.AddDays(1))
        ],
        _ => 0);

        Assert.Equal(["Newest", "Middle", "Older"], items.Select(item => item.Name));
        Assert.Equal("Commodore 128 - 310378 (PAL)", items[0].Board);
    }

    [Fact]
    public void A_card_says_its_board_its_captures_and_the_day_it_was_made()
    {
        CollectionListItem item = CollectionListing.ItemOf(Collection("No picture", Made, Capture('a', "U7", "12"), Capture('b', "U8", "1")), 0);

        Assert.Equal("No picture", item.Name);
        Assert.Equal("Commodore 64 - 250407 (PAL)", item.Board);
        Assert.Equal("2 captures | Made on 2026-November-26", item.Facts);
    }

    [Fact]
    public void A_board_without_regions_names_none() =>
        Assert.Equal("Commodore 64 - 250407", CollectionListing.BoardOf(Collection("x", Made) with { Region = null }));

    [Theory]
    [InlineData(0, "No captures yet")]
    [InlineData(1, "1 capture")]
    [InlineData(3, "3 captures")]
    public void The_captures_are_counted_in_words(int count, string words) => Assert.Equal(words, CollectionListing.CapturesOf(count));

    // The thumbnail is the first capture as the page orders them - by component, then pin - and none with no capture, or a
    // capture of no checksum (a file edited by hand), which has no picture to show.
    [Fact]
    public void The_thumbnail_is_the_page_s_first_capture()
    {
        CollectionListItem item = CollectionListing.ItemOf(Collection("x", Made, Capture('b', "U10", "1"), Capture('a', "U2", "12")), 0);

        Assert.Equal(new string('a', 64), item.FirstCaptureSha256);
        Assert.Null(CollectionListing.ItemOf(Collection("x", Made), 0).FirstCaptureSha256);
        Assert.Null(CollectionListing.ItemOf(Collection("x", Made, Capture('a', "U2", "12") with { Sha256 = "not one" }), 0).FirstCaptureSha256);
    }

    [Fact]
    public void A_collection_never_shared_is_on_this_computer_only()
    {
        CollectionListItem item = CollectionListing.ItemOf(Collection("x", Made, Capture('a', "U7", "12")), 0);

        Assert.Equal("Not shared - it is on this computer only.", item.State);
        Assert.Null(item.PageUrl);
        Assert.False(item.IsShared);
    }

    // No "Shared on" date (owner request, 2026-11-28: "no need for that") - when its page goes.
    [Fact]
    public void A_shared_collection_says_when_its_page_goes()
    {
        CollectionListItem item = CollectionListing.ItemOf(Shared(Collection("x", Made, Capture('a', "U7", "12")), Made.AddYears(1)), 0);

        Assert.Equal("Shared - its page is deleted on 2027-November-26.", item.State);
        Assert.Equal("https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p", item.PageUrl);
        Assert.True(item.IsShared);
    }

    [Fact]
    public void A_shared_collection_changed_since_says_so()
    {
        LocalCollection shared = Shared(Collection("x", Made, Capture('a', "U7", "12")), Made.AddYears(1));

        CollectionListItem item = CollectionListing.ItemOf(shared with { Description = "Typed after it was shared." }, 0);

        Assert.Equal("Shared, and changed since - its page is deleted on 2027-November-26.", item.State);
    }

    // Its page offline ("Stop sharing") is not shared: "Open shared page" is not offered (owner request, 2026-11-28: "only if it
    // is shared").
    [Fact]
    public void A_collection_whose_sharing_stopped_says_its_page_is_offline()
    {
        LocalCollection stopped = Shared(Collection("x", Made, Capture('a', "U7", "12")), Made.AddYears(1)) with { SharingStoppedUtc = Made.AddDays(3) };

        CollectionListItem item = CollectionListing.ItemOf(stopped, 0);

        Assert.Equal("Not shared at the moment - its page is offline.", item.State);
        Assert.False(item.IsShared);
    }

    // New comments: a shared page with a comment newer than its owner read - asked of the server by its id, never of a
    // collection with no page.
    [Fact]
    public void New_comments_are_those_newer_than_its_owner_read()
    {
        LocalCollection shared = Shared(Collection("Shared", Made, Capture('a', "U7", "12")), Made.AddYears(1), seenCommentId: 4);
        LocalCollection kept = Collection("Kept here", Made.AddDays(1));
        List<string> asked = [];

        IReadOnlyList<CollectionListItem> items = CollectionListing.ItemsOf([shared, kept], serverId =>
        {
            asked.Add(serverId);
            return 5;
        });

        Assert.Equal(["k7mq3x9vbn2p"], asked);
        Assert.False(items.Single(item => item.Name == "Kept here").HasNewComments);
        Assert.True(items.Single(item => item.Name == "Shared").HasNewComments);
        Assert.False(CollectionListing.ItemOf(shared, 4).HasNewComments);
    }

    [Fact]
    public void The_heading_counts_them_and_those_with_new_comments()
    {
        CollectionListItem quiet = CollectionListing.ItemOf(Collection("a", Made), 0);
        CollectionListItem talked = quiet with { HasNewComments = true };

        Assert.Equal("1 collection on this computer", CollectionListing.HeadingOf([quiet]));
        Assert.Equal("3 collections on this computer - 2 with new comments", CollectionListing.HeadingOf([quiet, talked, talked]));
    }

    // A row unchanged is the same row - the list keeps it (InPlaceList), so a press on it is never thrown away.
    [Fact]
    public void The_same_collection_gives_the_same_item()
    {
        LocalCollection collection = Collection("x", Made, Capture('a', "U7", "12"));

        Assert.Equal(CollectionListing.ItemOf(collection, 0), CollectionListing.ItemOf(collection, 0));
        Assert.NotEqual(CollectionListing.ItemOf(collection, 0), CollectionListing.ItemOf(collection with { Title = "y" }, 0));
    }
}
