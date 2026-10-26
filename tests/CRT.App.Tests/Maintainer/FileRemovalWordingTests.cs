using System.Text.Json;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers FileRemovalWording and the parsing behind it - the files a publish REMOVES, shown before
// anyone approves (owner, 2026-09-25: "It must be, so it is clear what will happen").
//
// The parsing tests serialise CRT.Data's own records the way the server does (ASP.NET's web
// defaults, camelCase) and read them back through the Maintainer tab's parser. The records are shared,
// so a renamed property fails here rather than as a silently empty list in the maintainer's hands.
public sealed class FileRemovalWordingTests
{
    private static string AsServer(object value) => JsonSerializer.Serialize(value, JsonSerializerOptions.Web);

    // ---- the wording ---------------------------------------------------------------------

    [Fact]
    public void The_headline_says_how_many_files_go_and_from_where()
    {
        Assert.Equal(
            "Publishing REMOVES 2 files from the BETA data - nothing uses them any more:",
            FileRemovalWording.Headline(new FileRemovalPreview(["a", "b"], null), "the BETA data"));

        Assert.Equal(
            "Publishing REMOVES 1 file from the stable data - nothing uses it any more:",
            FileRemovalWording.Headline(new FileRemovalPreview(["a"], null), "the stable data"));
    }

    // Said even when it is nothing, so the maintainer knows it was checked.
    [Fact]
    public void Nothing_to_remove_is_said_as_such()
    {
        Assert.Equal("No files are removed from the stable data.", FileRemovalWording.Headline(FileRemovalPreview.Nothing, "the stable data"));
        Assert.Null(FileRemovalWording.ApproveNote(FileRemovalPreview.Nothing));
    }

    [Fact]
    public void A_blocked_preview_says_why_nothing_is_removed()
    {
        var blocked = new FileRemovalPreview([], "Nothing is removed, because the data could not be read completely: x");

        Assert.Equal(blocked.BlockedBecause, FileRemovalWording.Headline(blocked, "the BETA data"));
        Assert.Null(FileRemovalWording.ApproveNote(blocked));
    }

    // An older server sends no list and removed nothing - there is nothing to say.
    [Fact]
    public void An_older_server_gets_no_headline()
    {
        Assert.Null(FileRemovalWording.Headline(null, "the BETA data"));
        Assert.Null(FileRemovalWording.ApproveNote(null));
    }

    // Beside the button, where it is read at the moment of pressing.
    [Fact]
    public void The_note_beside_the_button_says_approving_removes_files()
    {
        Assert.StartsWith("Approving also REMOVES 3 files", FileRemovalWording.ApproveNote(new FileRemovalPreview(["a", "b", "c"], null)));
    }

    // ###########################################################################################
    // *** A REPLACED FILE'S EARLIER BYTES ARE NO REMOVAL TO SPEAK OF (CRT 4.0 tree, 2026-10-25). ***
    // The publish removes them - stored under their own checksum, nothing lists them afterwards -
    // but under the very name the board keeps listing, so saying "manual.pdf - nothing uses it any
    // more" would announce a file going that is only replaced. A removed file is named by the name
    // the board knew it by, never by its stored path.
    // ###########################################################################################
    [Fact]
    public void The_earlier_bytes_of_a_replaced_file_are_not_counted_or_listed_as_a_file_nothing_uses()
    {
        var preview = new FileRemovalPreview(
            ["Files/3f/3f7a.pdf", "Files/9c/9c01.png"],
            null,
            new Dictionary<string, string> { ["Files/3f/3f7a.pdf"] = "manual.pdf", ["Files/9c/9c01.png"] = "Images/old.png" });

        string[] kept = ["Manual.PDF", "Images/sheet.png"];

        Assert.Equal(["Images/old.png"], FileRemovalWording.Unused(preview, kept));
        Assert.Equal(
            "Publishing REMOVES 1 file from the BETA data - nothing uses it any more:",
            FileRemovalWording.Headline(preview, "the BETA data", kept));
        Assert.StartsWith("Approving also REMOVES 1 file", FileRemovalWording.ApproveNote(preview, kept));

        // Nothing kept: both are files nothing uses.
        Assert.Equal(["manual.pdf", "Images/old.png"], FileRemovalWording.Unused(preview));

        // Only a replacement: nothing is said to go.
        var replacing = new FileRemovalPreview(["Files/3f/3f7a.pdf"], null, new Dictionary<string, string> { ["Files/3f/3f7a.pdf"] = "manual.pdf" });

        Assert.Equal("No files are removed from the stable data.", FileRemovalWording.Headline(replacing, "the stable data", kept));
        Assert.Null(FileRemovalWording.ApproveNote(replacing, kept));
    }

    [Fact]
    public void What_was_removed_is_added_to_the_message_after_publishing()
    {
        Assert.Equal(string.Empty, FileRemovalWording.Done(null));
        Assert.Equal(" 1 unused file was removed.", FileRemovalWording.Done(["a"]));
        Assert.Equal(" 2 unused files were removed.", FileRemovalWording.Done(["a", "b"]));
    }

    [Fact]
    public void A_path_is_on_the_list_whatever_its_case()
    {
        var preview = new FileRemovalPreview(["Commodore/C64/250407/Old.pdf"], null);

        Assert.True(FileRemovalWording.IsRemoved(preview, "commodore/c64/250407/old.PDF"));
        Assert.False(FileRemovalWording.IsRemoved(preview, "Commodore/C64/250407/other.pdf"));
        Assert.False(FileRemovalWording.IsRemoved(null, "Commodore/C64/250407/Old.pdf"));
    }

    // ---- reading it off the server -------------------------------------------------------

    [Fact]
    public void A_submission_detail_carries_the_removal_list_as_the_same_record()
    {
        string json = FileRemovalWordingTests.AsServer(new
        {
            canPublish = true,
            submission = new { id = 42, boardId = "Commodore/C64/250407", state = "pending", summary = "x" },
            findings = Array.Empty<object>(),
            removals = new FileRemovalPreview(["Commodore/C64/250407/old.pdf"], null)
        });

        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(json);

        Assert.Equal(["Commodore/C64/250407/old.pdf"], detail!.Removals!.Files);
        Assert.False(detail.Removals.IsBlocked);
    }

    [Fact]
    public void A_production_plan_carries_the_removal_list()
    {
        string json = FileRemovalWordingTests.AsServer(new
        {
            boardId = "Commodore/C64/250407",
            betaContentHash = "h1",
            canPublish = true,
            files = Array.Empty<object>(),
            problems = Array.Empty<object>(),
            removals = new FileRemovalPreview([], "Nothing is removed, because x")
        });

        ProductionPlanView? plan = ReviewApiParser.ParseProductionPlan(json);

        Assert.True(plan!.Removals!.IsBlocked);
        Assert.Empty(plan.Removals.Files);
    }

    [Fact]
    public void What_a_publish_removed_is_read_from_both_publish_answers()
    {
        ReviewDecisionResult? decision = ReviewApiParser.ParseDecision(
            FileRemovalWordingTests.AsServer(new { state = "merged", revision = "r2", removedFiles = new[] { "a.pdf" } }));
        ProductionPublishResult? promotion = ReviewApiParser.ParseProductionPublish(
            FileRemovalWordingTests.AsServer(new { boardId = "Commodore/C64/250407", state = "published", filesCopied = 3, removedFiles = new[] { "a.pdf", "b.pdf" } }));

        Assert.Equal(["a.pdf"], decision!.RemovedFiles);
        Assert.Equal(["a.pdf", "b.pdf"], promotion!.RemovedFiles);
        Assert.Contains("1 unused file was removed.", ReviewDecisionWording.Describe(ReviewDecisionKind.Approve, decision));
    }

    // An answer from before removals existed: none listed, none removed, and still a usable answer.
    [Fact]
    public void An_older_servers_answers_parse_with_nothing_removed()
    {
        ReviewDecisionResult? decision = ReviewApiParser.ParseDecision("""{"state":"merged","revision":"r2"}""");

        Assert.Empty(decision!.RemovedFiles!);
        Assert.DoesNotContain("removed", ReviewDecisionWording.Describe(ReviewDecisionKind.Approve, decision), StringComparison.Ordinal);
    }
}
