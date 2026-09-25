using System.Text.Json;
using CRT.Review.Handlers;
using Handlers.DataHandling;

namespace CRT.Review.Tests;

// Covers FileRemovalWording and the parsing behind it - the files a publish REMOVES, shown before
// anyone approves (maintainer, 2026-09-25: "It must be, so it is clear what will happen").
//
// The parsing tests serialise CRT.Data's own records the way the server does (ASP.NET's web
// defaults, camelCase) and read them back through the review app's parser. The records are shared,
// so a renamed property fails here rather than as a silently empty list in the reviewer's hands.
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
            "Publishing REMOVES 1 file from production - nothing uses it any more:",
            FileRemovalWording.Headline(new FileRemovalPreview(["a"], null), "production"));
    }

    // Said even when it is nothing, so the reviewer knows it was checked.
    [Fact]
    public void Nothing_to_remove_is_said_as_such()
    {
        Assert.Equal("No files are removed from production.", FileRemovalWording.Headline(FileRemovalPreview.Nothing, "production"));
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
            submission = new { id = 42, systemId = "Commodore/C64/250407", state = "pending", summary = "x" },
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
            systemId = "Commodore/C64/250407",
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
            FileRemovalWordingTests.AsServer(new { systemId = "Commodore/C64/250407", state = "published", filesCopied = 3, removedFiles = new[] { "a.pdf", "b.pdf" } }));

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
