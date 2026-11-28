using Handlers.Collections;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// A collection as it is shared (CollectionSharing): the request sent, the page's order, a change since it was shared, and
// what "Copy link" puts on the clipboard.
// ###########################################################################################
public sealed class CollectionSharingTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private static LocalCapture Capture(string label, string? pin, int minute = 0, CollectionComponent? component = null) =>
        new()
        {
            Sha256 = new string((char)('a' + (label.Length + (pin?.Length ?? 0) + minute) % 6), 64),
            ComponentLabel = label,
            Pin = pin,
            SignalName = "/RAS",
            TimePerDivision = "50us",
            Component = component,
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

    // The oscilloscopes the captures were taken with, each once as first taken (owner request, 2026-11-28) - none known, none sent,
    // so a collection shared before CRT kept them reads as unchanged.
    [Fact]
    public void The_request_names_each_oscilloscope_once_as_first_taken()
    {
        CollectionContentRequest request = CollectionSharing.RequestOf(
            Collection(
                Capture("U2", "1", minute: 3) with { Oscilloscope = "Siglent SDS1104X-E" },
                Capture("U1", "1", minute: 1) with { Oscilloscope = "Rigol DS1054Z" },
                Capture("U3", "1", minute: 5) with { Oscilloscope = "rigol ds1054z" },
                Capture("U4", "1", minute: 7)),
            Now.AddYears(1));

        Assert.Equal(["Rigol DS1054Z", "Siglent SDS1104X-E"], request.Oscilloscopes);
        Assert.Null(CollectionSharing.RequestOf(Collection(Capture("U1", "1")), Now.AddYears(1)).Oscilloscopes);
    }

    [Fact]
    public void The_request_carries_every_capture_in_the_page_s_order_and_its_pinout_drawing()
    {
        CollectionContentRequest request = CollectionSharing.RequestOf(
            Collection(Capture("U8", "1"), Capture("U7", "12") with { PinoutSha256 = new string('d', 64) }),
            Now.AddYears(1));

        Assert.Equal(("Commodore 64", "250407", "PAL", "No picture"), (request.HardwareName, request.BoardName, request.Region, request.Title));
        Assert.Equal(Now.AddYears(1), request.DeleteAtUtc);
        Assert.Equal(["U7", "U8"], request.Images.Select(image => image.ComponentLabel));
        Assert.Equal(new string('d', 64), request.Images[0].PinoutSha256);
        Assert.Null(request.Images[1].PinoutSha256);
        Assert.Null(request.Components);
    }

    // Each component captured once, as its LATEST capture told of it - named as its captures name it, in the page's order.
    [Fact]
    public void The_request_tells_of_each_component_once_as_last_captured()
    {
        CollectionComponent before = new("U7", "U7 | old", null, null, null, null);
        CollectionComponent after = new("U7", "U7 | VIC-II", null, null, null, null);
        CollectionComponent cpu = new("u2", "U2 | CPU", null, null, null, null);

        CollectionContentRequest request = CollectionSharing.RequestOf(
            Collection(Capture("U7", "12", minute: 1, component: before), Capture("U7", "13", minute: 9, component: after), Capture("U2", "1", component: cpu), Capture("U8", "1")),
            Now.AddYears(1));

        Assert.Equal([("U2", "U2 | CPU"), ("U7", "U7 | VIC-II")], request.Components!.Select(component => (component.Label, component.Title)));
    }

    // Owner decision, 2026-11-27: "every time the user shares it again, it resets its time" - counted from every share, whatever
    // the last share sent.
    [Fact]
    public void The_deletion_date_is_the_choice_counted_from_the_share()
    {
        LocalCollection collection = Collection() with { Lifetime = "1w" };

        Assert.Equal(Now.AddDays(7), CollectionSharing.DeleteAtFor(collection, Now));
        Assert.Equal(Now.AddDays(7), CollectionSharing.DeleteAtFor(collection with { DeleteAtUtc = Now.AddMonths(1) }, Now));
        Assert.Equal(Now.AddYears(1), CollectionSharing.DeleteAtFor(collection with { Lifetime = "5y" }, Now));

        // "6 months" is no choice since 2026-11-28 (owner request: "1 day, 1 week, 1 month, 1 year") - kept, it reads as a year.
        Assert.Equal(Now.AddYears(1), CollectionSharing.DeleteAtFor(collection with { Lifetime = "6m" }, Now));
    }

    [Fact]
    public void A_change_since_it_was_shared_is_known_by_what_was_sent()
    {
        LocalCollection collection = Collection(Capture("U7", "12")) with
        {
            OwnerKey = new string('c', CollectionRules.OwnerKeyLength),
            ServerId = "k7mq3x9vbn2p",
            DeleteAtUtc = Now.AddYears(1)
        };

        string sent = CollectionSharing.ContentHash(CollectionSharing.RequestOf(collection, Now.AddYears(1)));
        LocalCollection shared = collection with { SharedContentHash = sent };

        Assert.False(CollectionSharing.HasChangesNotShared(shared));
        Assert.True(CollectionSharing.HasChangesNotShared(shared with { Description = "Found the bad trace" }));
        Assert.True(CollectionSharing.HasChangesNotShared(shared with { DeleteAtUtc = Now.AddYears(5) }));

        // Never shared: nothing to send a change to - nor shared with an account before 2026-11-28, which no key here owns.
        Assert.False(CollectionSharing.HasChangesNotShared(shared with { ServerId = null, Description = "x" }));
        Assert.False(CollectionSharing.HasChangesNotShared(shared with { OwnerKey = null, Description = "x" }));
    }

    // The link and a line saying what it is, in each format (owner decision, 2026-11-28) - each capture's own link went with the
    // "Looks wrong" mark (owner request, 2026-11-27).
    [Fact]
    public void Copy_link_is_the_link_and_its_line_in_each_format()
    {
        LocalCollection shared = Collection(Capture("U7", "12"), Capture("U8", "1")) with
        {
            PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p"
        };

        Assert.Equal(
            "[url=https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p]No picture[/url] - 2 oscilloscope captures of Commodore 64 250407 (PAL), shared with Classic Repair Toolbox",
            CollectionSharing.ForumText(shared, CollectionLinkFormat.BbCode));

        Assert.Equal(
            "[No picture](https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p) - 2 oscilloscope captures of Commodore 64 250407 (PAL), shared with Classic Repair Toolbox",
            CollectionSharing.ForumText(shared, CollectionLinkFormat.Markdown));

        Assert.Equal(
            "No picture - 2 oscilloscope captures of Commodore 64 250407 (PAL), shared with Classic Repair Toolbox: https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p",
            CollectionSharing.ForumText(shared, CollectionLinkFormat.Plain));

        Assert.Equal(
            "<a href=\"https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p\">No picture</a> - 2 oscilloscope captures of Commodore 64 250407 (PAL), shared with Classic Repair Toolbox",
            CollectionSharing.ForumText(shared, CollectionLinkFormat.Html));

        // The line after an HTML link is HTML too.
        Assert.EndsWith(
            "of Amstrad &amp; Schneider CPC 664 (PAL), shared with Classic Repair Toolbox",
            CollectionSharing.ForumText(shared with { HardwareName = "Amstrad & Schneider", BoardName = "CPC 664" }, CollectionLinkFormat.Html),
            StringComparison.Ordinal);
    }

    // A capture copied into a worklog as a photo (owner decision, 2026-11-27: "A copy of its captures") says where it was taken,
    // the settings, what it should read and its note - the values bold, as a note with formatting.
    [Fact]
    public void A_capture_s_worklog_comment_says_its_place_settings_reading_and_note()
    {
        LocalCapture capture = Capture("U7", "12") with
        {
            VoltsPerDivision = "2V",
            ExpectedReading = "Pulsing",
            Caption = "Stays low\nafter reset"
        };

        string comment = CollectionSharing.WorklogCommentOf(capture);

        Assert.Equal(
            "**U7, pin 12** - /RAS\n" +
            "This capture: T/DIV **50us** | V/DIV **2V** | Trigger level not read\n" +
            "Expected reading: **Pulsing**\n" +
            "Stays low\nafter reset",
            comment);
        Assert.Equal(
            "U7, pin 12 - /RAS\nThis capture: T/DIV 50us | V/DIV 2V | Trigger level not read\nExpected reading: Pulsing\nStays low\nafter reset",
            RichNote.PlainOf(comment));
    }

    // A character the note format reads stays a character.
    [Fact]
    public void A_capture_s_worklog_comment_keeps_a_star_in_its_note_a_star()
    {
        string comment = CollectionSharing.WorklogCommentOf(Capture("U7", null) with { Caption = "*very* odd" });

        Assert.EndsWith("*very* odd", RichNote.PlainOf(comment), StringComparison.Ordinal);
        Assert.StartsWith("**U7** - /RAS\n", comment, StringComparison.Ordinal);
    }

    [Fact]
    public void A_collection_not_shared_has_nothing_to_copy() =>
        Assert.Equal(string.Empty, CollectionSharing.ForumText(Collection(Capture("U7", "12")), CollectionLinkFormat.BbCode));

    [Fact]
    public void A_collection_is_named_by_its_title_and_listed_with_its_count()
    {
        Assert.Equal("No picture (2 captures)", CollectionSharing.LineOf(Collection(Capture("U7", "1"), Capture("U7", "2"))));
        Assert.Equal("Untitled collection (1 capture)", CollectionSharing.LineOf(Collection(Capture("U7", "1")) with { Title = " " }));
        Assert.Equal("U7 captures, 2026-November-26", CollectionSharing.DefaultTitle("U7", Now));
    }
}
