using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// BOTH ENDS OF THE COLLECTION ROUTES, IN ONE TEST (code review, 2026-11-27: CollectionContract's header promised this test,
// and none was there). As ReviewWireContractTests for the review API: a REQUEST written as CollectionClient writes it
// (ReviewApiContract.WireSettings) and read with the SERVER's settings (JsonSerializerDefaults.Web plus ApplyWireSettings,
// what Program.cs configures) into the record the server binds; an ANSWER written with the server's settings and read as
// CollectionClient reads it. A field renamed on one side only fails here.
// ###########################################################################################
public sealed class CollectionWireContractTests
{
    private static readonly JsonSerializerOptions Server = CollectionWireContractTests.ServerSettings();

    private static readonly DateTimeOffset When = new(2026, 11, 27, 12, 0, 0, TimeSpan.Zero);

    private static readonly string Sha = new('a', 64);

    private static JsonSerializerOptions ServerSettings()
    {
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
        ReviewApiContract.ApplyWireSettings(options);
        return options;
    }

    // What the server binds from what CRT sends.
    private static T Received<T>(T request) =>
        JsonSerializer.Deserialize<T>(JsonSerializer.Serialize(request, ReviewApiContract.WireSettings), CollectionWireContractTests.Server)!;

    // What CRT reads of what the server answers.
    private static T Read<T>(T answer) =>
        JsonSerializer.Deserialize<T>(JsonSerializer.Serialize(answer, CollectionWireContractTests.Server), ReviewApiContract.WireSettings)!;

    [Fact]
    public void A_collection_arrives_whole_its_captures_and_their_components_with_it()
    {
        var sent = new CollectionContentRequest(
            "Commodore 64", "250407", "PAL", "No picture", "It went black.", ["https://www.lemon64.com/forum/viewtopic.php?t=1"], When.AddYears(1),
            [new CollectionImageEntry(Sha, "U7", "12", "/RAS", "Pulsing", "50us", "2V", "1V", "Stays low", When, PinoutSha256: new string('b', 64))],
            [new CollectionComponent("U7", "U7 | VIC-II | 6569", "IC | DIP-40", "Video chip", "Makes the picture", [new CollectionPin("12", "/RAS", "Output", "RAS", "Pulsing\nLow [64K]")])]);

        CollectionContentRequest received = CollectionWireContractTests.Received(sent);

        Assert.Equal((sent.HardwareName, sent.BoardName, sent.Region, sent.Title, sent.Description, sent.DeleteAtUtc), (received.HardwareName, received.BoardName, received.Region, received.Title, received.Description, received.DeleteAtUtc));
        Assert.Equal(sent.ForumLinks, received.ForumLinks);
        Assert.Equal(sent.Images[0], received.Images[0]);

        CollectionComponent component = Assert.Single(received.Components!);

        Assert.Equal(sent.Components![0] with { Pins = null }, component with { Pins = null });
        Assert.Equal(sent.Components[0].Pins, component.Pins);
    }

    // "Stop sharing" and sharing again with nothing changed.
    [Fact]
    public void A_sharing_change_arrives_as_sent()
    {
        Assert.False(CollectionWireContractTests.Received(new CollectionSharingRequest(false)).Shared);
        Assert.True(CollectionWireContractTests.Received(new CollectionSharingRequest(true)).Shared);
    }

    [Fact]
    public void A_share_s_answers_are_read_as_the_server_writes_them()
    {
        Assert.Equal(new CollectionAnswer("k7mq3x9vbn2p", "https://x/collections/k7mq3x9vbn2p", When, When.AddYears(1), 3), CollectionWireContractTests.Read(new CollectionAnswer("k7mq3x9vbn2p", "https://x/collections/k7mq3x9vbn2p", When, When.AddYears(1), 3)));
        Assert.Equal(new CollectionImageUploadAnswer(Sha, 800, 480), CollectionWireContractTests.Read(new CollectionImageUploadAnswer(Sha, 800, 480)));

        CollectionRefusalAnswer refusal = CollectionWireContractTests.Read(new CollectionRefusalAnswer(["The server does not hold 1 of the captures yet."], [Sha]));

        Assert.Equal(["The server does not hold 1 of the captures yet."], refusal.Errors);
        Assert.Equal([Sha], refusal.MissingImages);
    }
}
