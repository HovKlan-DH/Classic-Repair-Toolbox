using Handlers.Collections;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// A collection as it is shared (CollectionSharing): the request sent, the page's order, a change since it was shared, and
// what "Copy for forum" puts on the clipboard.
// ###########################################################################################
public sealed class CollectionSharingTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private static LocalCapture Capture(string label, string? pin, int minute = 0, bool looksWrong = false, string? baselineTdiv = null) =>
        new()
        {
            Sha256 = new string((char)('a' + (label.Length + (pin?.Length ?? 0) + minute) % 6), 64),
            ComponentLabel = label,
            Pin = pin,
            SignalName = "/RAS",
            TimePerDivision = "50us",
            BaselineTimePerDivision = baselineTdiv,
            LooksWrong = looksWrong,
            CapturedUtc = Now.AddMinutes(minute)
        };

    private static LocalCollection Collection(params LocalCapture[] captures) =>
        new()
        {
            LocalId = new string('0', 32),
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Region = "PAL",
            Title = "No picture",
            CreatedUtc = Now,
            Captures = captures
        };

    [Fact]
    public void Captures_are_shown_by_component_then_pin_then_as_taken()
    {
        IReadOnlyList<LocalCapture> ordered = CollectionSharing.InPageOrder(
        [
            Capture("U10", "1"),
            Capture("U2", "12", minute: 5),
            Capture("U2", "2"),
            Capture("U2", "12", minute: 1)
        ]);

        Assert.Equal(
            ["U2 2", "U2 12 1", "U2 12 5", "U10 1"],
            ordered.Select(capture => $"{capture.ComponentLabel} {capture.Pin}" + (capture.Pin == "12" ? $" {capture.CapturedUtc.Minute}" : string.Empty)));
    }

    [Fact]
    public void The_request_carries_every_capture_in_the_page_s_order_and_a_baseline_only_where_there_is_one()
    {
        CollectionContentRequest request = CollectionSharing.RequestOf(
            Collection(Capture("U8", "1"), Capture("U7", "12", baselineTdiv: "100us")),
            Now.AddYears(1));

        Assert.Equal(("Commodore 64", "250407", "PAL", "No picture"), (request.HardwareName, request.BoardName, request.Region, request.Title));
        Assert.Equal(Now.AddYears(1), request.DeleteAtUtc);
        Assert.Equal(["U7", "U8"], request.Images.Select(image => image.ComponentLabel));
        Assert.Equal("100us", request.Images[0].Baseline?.TimePerDivision);
        Assert.Null(request.Images[1].Baseline);
    }

    [Fact]
    public void The_deletion_date_is_the_one_kept_else_the_choice_counted_from_now()
    {
        LocalCollection collection = Collection() with { Lifetime = "6m" };

        Assert.Equal(Now.AddMonths(6), CollectionSharing.DeleteAtFor(collection, Now));
        Assert.Equal(Now.AddYears(3), CollectionSharing.DeleteAtFor(collection with { DeleteAtUtc = Now.AddYears(3) }, Now));
    }

    [Fact]
    public void A_change_since_it_was_shared_is_known_by_what_was_sent()
    {
        LocalCollection collection = Collection(Capture("U7", "12")) with { ServerId = "k7mq3x9vbn2p", DeleteAtUtc = Now.AddYears(1) };

        string sent = CollectionSharing.ContentHash(CollectionSharing.RequestOf(collection, Now.AddYears(1)));
        LocalCollection shared = collection with { SharedContentHash = sent };

        Assert.False(CollectionSharing.HasChangesNotShared(shared));
        Assert.True(CollectionSharing.HasChangesNotShared(shared with { Description = "Found the bad trace" }));
        Assert.True(CollectionSharing.HasChangesNotShared(shared with { DeleteAtUtc = Now.AddYears(5) }));

        // Never shared: nothing to send a change to.
        Assert.False(CollectionSharing.HasChangesNotShared(shared with { ServerId = null, Description = "x" }));
    }

    [Fact]
    public void Copy_for_forum_links_the_page_and_each_capture_that_looks_wrong()
    {
        LocalCollection shared = Collection(Capture("U7", "12", looksWrong: true), Capture("U8", "1")) with
        {
            PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p"
        };

        string bbcode = CollectionSharing.ForumText(shared, CollectionSharing.ForumFormat.BbCode);
        string markdown = CollectionSharing.ForumText(shared, CollectionSharing.ForumFormat.Markdown);

        Assert.StartsWith(
            "[url=https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p]No picture[/url] - 2 oscilloscope captures of Commodore 64 250407 (PAL), shared with Classic Repair Toolbox",
            bbcode,
            StringComparison.Ordinal);
        Assert.Contains("[list]\n[*] [url=https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p#U7-12]U7, pin 12 (/RAS)[/url]\n[/list]", bbcode, StringComparison.Ordinal);

        Assert.StartsWith("[No picture](https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p) - 2 oscilloscope captures", markdown, StringComparison.Ordinal);
        Assert.Contains("\n- [U7, pin 12 (/RAS)](https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p#U7-12)", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("U8", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Nothing_looking_wrong_is_the_link_and_its_line_alone()
    {
        LocalCollection shared = Collection(Capture("U7", "12")) with { PageUrl = "https://x/collections/a" };

        Assert.Equal(
            "[No picture](https://x/collections/a) - 1 oscilloscope capture of Commodore 64 250407 (PAL), shared with Classic Repair Toolbox",
            CollectionSharing.ForumText(shared, CollectionSharing.ForumFormat.Markdown));
    }

    [Fact]
    public void A_collection_not_shared_has_nothing_to_copy() =>
        Assert.Equal(string.Empty, CollectionSharing.ForumText(Collection(Capture("U7", "12")), CollectionSharing.ForumFormat.BbCode));

    // Comments newer than the owner read in CRT mark the collection in the popups' drop-downs (step 3).
    [Fact]
    public void A_shared_collection_with_comments_not_read_says_so()
    {
        LocalCollection shared = Collection(Capture("U7", "1")) with { ServerId = "k7mq3x9vbn2p", SeenCommentId = 5 };

        Assert.Equal("No picture (1 capture, new comments)", CollectionSharing.LineOf(shared, latestCommentId: 6));
        Assert.Equal("No picture (1 capture)", CollectionSharing.LineOf(shared, latestCommentId: 5));
        Assert.False(CollectionSharing.HasNewComments(shared with { ServerId = null }, 6));
    }

    [Fact]
    public void A_collection_is_named_by_its_title_and_listed_with_its_count()
    {
        Assert.Equal("No picture (2 captures)", CollectionSharing.LineOf(Collection(Capture("U7", "1"), Capture("U7", "2"))));
        Assert.Equal("Untitled collection (1 capture)", CollectionSharing.LineOf(Collection(Capture("U7", "1")) with { Title = " " }));
        Assert.Equal("U7 captures, 2026-November-26", CollectionSharing.DefaultTitle("U7", Now));
    }
}
