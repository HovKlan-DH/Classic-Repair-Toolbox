using System.Security.Cryptography;
using System.Text.Json.Nodes;
using Handlers.Collections;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// The oscilloscope collections kept on this computer (CollectionStore) - in a folder of the test's own, never the user's.
// ###########################################################################################
public sealed class CollectionStoreTests : IDisposable
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private readonly TempWorkspace thisWorkspace = new();

    private readonly CollectionStore thisStore;

    public CollectionStoreTests()
    {
        this.thisStore = new CollectionStore(this.thisWorkspace.Root);
    }

    public void Dispose() => this.thisWorkspace.Dispose();

    private static LocalCapture Capture(string label = "U7", string pin = "12") =>
        new() { ComponentLabel = label, Pin = pin, SignalName = "/RAS", TimePerDivision = "50us", CapturedUtc = Now };

    private static string ShaOf(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

    [Fact]
    public void A_new_collection_is_kept_in_a_folder_of_its_own_and_read_back()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", "PAL", "U7 captures", Now)!;

        Assert.True(CollectionStore.IsLocalId(made.LocalId));
        Assert.True(File.Exists(Path.Combine(this.thisWorkspace.Root, made.LocalId, CollectionStore.FileName)));

        LocalCollection read = this.thisStore.Find(made.LocalId)!;

        Assert.Equal(("Commodore 64", "250407", "PAL", "U7 captures"), (read.HardwareName, read.BoardName, read.Region, read.Title));
        Assert.Equal("1y", read.Lifetime);
        Assert.False(read.IsShared);
        Assert.Empty(read.Captures);
    }

    [Fact]
    public void A_board_lists_its_own_collections_newest_first_capitals_aside()
    {
        LocalCollection older = this.thisStore.Create("Commodore 64", "250407", null, "Older", Now)!;
        LocalCollection newer = this.thisStore.Create("Commodore 64", "250407", null, "Newer", Now.AddHours(1))!;
        _ = this.thisStore.Create("Commodore 64", "250425", null, "Another board", Now);

        IReadOnlyList<LocalCollection> listed = this.thisStore.OfBoard("COMMODORE 64", "250407");

        Assert.Equal([newer.LocalId, older.LocalId], listed.Select(collection => collection.LocalId));
    }

    [Fact]
    public void A_capture_is_kept_by_its_checksum_and_the_same_picture_twice_is_one()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        byte[] png = [1, 2, 3, 4];

        LocalCollection once = this.thisStore.AddCapture(made.LocalId, png, Capture())!;
        LocalCollection twice = this.thisStore.AddCapture(made.LocalId, png, Capture("U8", "3"))!;

        LocalCapture kept = Assert.Single(twice.Captures);

        Assert.Equal(ShaOf(png), kept.Sha256);
        Assert.Equal("U7", kept.ComponentLabel);
        Assert.Equal(png, File.ReadAllBytes(this.thisStore.CapturePath(made.LocalId, ShaOf(png))));
        Assert.Equal(png, this.thisStore.ReadCapture(made.LocalId, ShaOf(png)));
        Assert.Single(once.Captures);
    }

    [Fact]
    public void A_capture_taken_out_goes_with_its_file()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        byte[] png = [9, 9, 9];

        this.thisStore.AddCapture(made.LocalId, png, Capture());

        LocalCollection changed = this.thisStore.RemoveCapture(made.LocalId, ShaOf(png))!;

        Assert.Empty(changed.Captures);
        Assert.False(File.Exists(this.thisStore.CapturePath(made.LocalId, ShaOf(png))));
        Assert.Null(this.thisStore.ReadCapture(made.LocalId, ShaOf(png)));
    }

    [Fact]
    public void A_collection_deleted_goes_whole()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        this.thisStore.AddCapture(made.LocalId, [1], Capture());

        Assert.True(this.thisStore.Delete(made.LocalId));
        Assert.False(Directory.Exists(Path.Combine(this.thisWorkspace.Root, made.LocalId)));
        Assert.Null(this.thisStore.Find(made.LocalId));
    }

    // Every path is made from an id of the right shape - nothing leads out of the folder.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("..")]
    [InlineData("../../Windows")]
    [InlineData("0123456789ABCDEF0123456789ABCDEF")]
    public void Anything_but_a_local_id_finds_and_deletes_nothing(string? id)
    {
        Assert.False(CollectionStore.IsLocalId(id));
        Assert.Null(this.thisStore.Find(id));
        Assert.False(this.thisStore.Delete(id ?? string.Empty));
    }

    [Fact]
    public void A_collection_whose_file_cannot_be_read_is_left_out_and_the_rest_listed()
    {
        LocalCollection good = this.thisStore.Create("Commodore 64", "250407", null, "Good", Now)!;
        string broken = Path.Combine(this.thisWorkspace.Root, Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(broken);
        File.WriteAllText(Path.Combine(broken, CollectionStore.FileName), "{ not json");

        Assert.Equal([good.LocalId], this.thisStore.All().Select(collection => collection.LocalId));
    }

    // A field a later CRT writes is kept and written back.
    [Fact]
    public void A_field_this_CRT_does_not_know_is_kept()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        string path = Path.Combine(this.thisWorkspace.Root, made.LocalId, CollectionStore.FileName);

        JsonObject json = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        json["laterField"] = "kept";
        File.WriteAllText(path, json.ToJsonString());

        this.thisStore.Save(this.thisStore.Find(made.LocalId)! with { Title = "Renamed" });

        JsonObject saved = JsonNode.Parse(File.ReadAllText(path))!.AsObject();

        Assert.Equal("kept", (string?)saved["laterField"]);
        Assert.Equal("Renamed", (string?)saved["title"]);
    }
}
