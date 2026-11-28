using Avalonia;
using Handlers.Collections;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// A collection attached to a worklog (CollectionWorklogAttach; owner decision, 2026-11-27: "A copy of its captures"): every
// capture copied in as a photo with its comment, in the page's order, and the page's link while it is up - into an existing
// worklog through the worklog's own writer, or into one being made through its editor's.
//
// Touches WorklogManager's static root, so this joins the "Worklog" collection.
// ###########################################################################################
[Collection("Worklog")]
public sealed class CollectionWorklogAttachTests : IDisposable
{
    // The key a shared collection's page is owned by - shared means shared under one (LocalCollection.IsShared).
    private static readonly string Key = new('c', CollectionRules.OwnerKeyLength);

    private readonly TempWorkspace thisWorkspace = new();

    private readonly CollectionStore thisStore;

    public CollectionWorklogAttachTests()
    {
        this.thisStore = new CollectionStore(this.thisWorkspace.Path_("collections"));
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
    }

    public void Dispose()
    {
        // Detach from the temp folder so nothing written later can reach the user's real one.
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    private (int WorkbookId, int EntryId) Worklog()
    {
        WorkbookRecord workbook = WorklogManager.CreateWorkbook("Commodore 64|250407", "No picture", "")!;
        WorklogEntryRecord entry = WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad PLA", "", "Issue", "Open", ["U17"])!;

        return (workbook.Id, entry.Id);
    }

    // U8 pin 3 and U7 pin 12, taken in that order - the page shows U7 first.
    private LocalCollection Collection()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", "PAL", "No picture", DateTimeOffset.UtcNow)!;

        this.thisStore.AddCapture(made.LocalId, CRT.Data.Tests.TestPng.Rgb(16, 10, 1), new LocalCapture { ComponentLabel = "U8", Pin = "3", TimePerDivision = "50us", CapturedUtc = DateTimeOffset.UtcNow });

        return this.thisStore.AddCapture(made.LocalId, CRT.Data.Tests.TestPng.Rgb(16, 10, 2), new LocalCapture { ComponentLabel = "U7", Pin = "12", Caption = "Stays low", CapturedUtc = DateTimeOffset.UtcNow }).Collection!;
    }

    private string CapturePath(LocalCollection collection, string sha256) => this.thisStore.CapturePath(collection.LocalId, sha256);

    [Fact]
    public void Every_capture_goes_into_the_worklog_as_a_photo_with_its_comment_in_the_page_s_order()
    {
        (int workbookId, int entryId) = this.Worklog();
        LocalCollection collection = this.Collection();

        CollectionWorklogAttach.Outcome outcome = CollectionWorklogAttach.ToEntry(workbookId, entryId, collection, sha => this.CapturePath(collection, sha));

        Assert.Equal(new CollectionWorklogAttach.Outcome(2, 0, false), outcome);

        WorklogEntryRecord stored = WorklogManager.GetEntries(workbookId).Single(entry => entry.Id == entryId);
        IReadOnlyList<LocalCapture> inOrder = CollectionSharing.InPageOrder(collection.Captures);

        Assert.Equal(inOrder.Select(CollectionSharing.WorklogCommentOf), stored.Photos.OrderBy(photo => photo.DisplayOrder).Select(photo => photo.Comment));
        Assert.All(stored.Photos, photo => Assert.True(File.Exists(Path.Combine(WorklogManager.GetEntryAttachmentsFolderPath(workbookId, entryId)!, photo.FileName))));
        Assert.Empty(stored.Links);

        // A copy: the collection keeps its own captures.
        Assert.All(collection.Captures, capture => Assert.True(File.Exists(this.CapturePath(collection, capture.Sha256))));
    }

    // Shared, its page goes in as a link of interest - once, however often the collection is attached.
    [Fact]
    public void A_shared_collection_s_page_goes_in_as_a_link_of_interest_once()
    {
        (int workbookId, int entryId) = this.Worklog();
        LocalCollection shared = this.Collection() with { OwnerKey = Key, ServerId = "k7mq3x9vbn2p", PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p" };

        Assert.True(CollectionWorklogAttach.ToEntry(workbookId, entryId, shared, sha => this.CapturePath(shared, sha)).Linked);
        CollectionWorklogAttach.ToEntry(workbookId, entryId, shared, sha => this.CapturePath(shared, sha));

        WorklogLinkRecord link = Assert.Single(WorklogManager.GetEntries(workbookId).Single(entry => entry.Id == entryId).Links);

        Assert.Equal(("Oscilloscope collection: No picture", "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p"), (link.Headline, link.Url));
    }

    // Its sharing stopped, its link says only that - so no link goes in.
    [Fact]
    public void A_collection_whose_sharing_stopped_goes_in_without_its_link()
    {
        (int workbookId, int entryId) = this.Worklog();
        LocalCollection stopped = this.Collection() with
        {
            OwnerKey = Key,
            ServerId = "k7mq3x9vbn2p",
            PageUrl = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p",
            SharingStoppedUtc = DateTimeOffset.UtcNow
        };

        Assert.False(CollectionWorklogAttach.ToEntry(workbookId, entryId, stopped, sha => this.CapturePath(stopped, sha)).Linked);
        Assert.Empty(WorklogManager.GetEntries(workbookId).Single(entry => entry.Id == entryId).Links);
    }

    // A capture whose file is gone from this computer is left out, and said.
    [Fact]
    public void A_capture_whose_file_is_gone_is_left_out_and_counted()
    {
        (int workbookId, int entryId) = this.Worklog();
        LocalCollection collection = this.Collection();
        File.Delete(this.CapturePath(collection, collection.Captures[0].Sha256));

        CollectionWorklogAttach.Outcome outcome = CollectionWorklogAttach.ToEntry(workbookId, entryId, collection, sha => this.CapturePath(collection, sha));

        Assert.Equal((1, 1), (outcome.Added, outcome.NotAdded));
        Assert.Single(WorklogManager.GetEntries(workbookId).Single(entry => entry.Id == entryId).Photos);
        Assert.EndsWith("- 1 could not be, their files gone from this computer.", CollectionWorklogAttach.Said(outcome, "worklog #1"), StringComparison.Ordinal);
    }

    // A new worklog: the captures and the link through its editor's own attach and link, in the page's order.
    [Fact]
    public void A_new_worklog_gets_the_captures_and_the_link_through_its_editor()
    {
        LocalCollection shared = this.Collection() with { OwnerKey = Key, ServerId = "k7mq3x9vbn2p", PageUrl = "https://x/collections/k7mq3x9vbn2p" };
        var attached = new List<string>();
        var linked = new List<(string, string)>();

        CollectionWorklogAttach.Outcome outcome = CollectionWorklogAttach.ToEditor(
            shared,
            sha => this.CapturePath(shared, sha),
            (_, comment) =>
            {
                attached.Add(comment);
                return true;
            },
            (headline, url) => linked.Add((headline, url)));

        Assert.Equal(new CollectionWorklogAttach.Outcome(2, 0, true), outcome);
        Assert.Equal(CollectionSharing.InPageOrder(shared.Captures).Select(CollectionSharing.WorklogCommentOf), attached);
        Assert.Equal([("Oscilloscope collection: No picture", "https://x/collections/k7mq3x9vbn2p")], linked);
    }

    [Fact]
    public void What_came_of_it_is_said_in_a_sentence()
    {
        Assert.Equal(
            "2 captures attached to worklog #3 as photos, with the page's link under \"Links of interest\".",
            CollectionWorklogAttach.Said(new CollectionWorklogAttach.Outcome(2, 0, true), "worklog #3"));
        Assert.Equal("1 capture attached to a new worklog as photos.", CollectionWorklogAttach.Said(new CollectionWorklogAttach.Outcome(1, 0, false), "a new worklog"));
        Assert.Equal("No capture could be attached to worklog #3.", CollectionWorklogAttach.Said(new CollectionWorklogAttach.Outcome(0, 2, false), "worklog #3"));
    }
}
