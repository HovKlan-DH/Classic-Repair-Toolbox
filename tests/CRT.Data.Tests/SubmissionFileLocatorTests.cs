using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Tests for SubmissionFileLocator - which of the TWO roots a submitted file's bytes come from.
//
// THE BUG THIS EXISTS TO STOP. The submission path originally resolved every referenced file
// against a single root, the system's draft folder. That is correct only for a system created
// from nothing: for the ordinary case - a draft over a published board - the officially
// published files live under Data/ and are not in the draft folder at all. Submitting a
// one-line typo fix to a board with 240 images would have reported all 240 as "not on disk"
// and refused to send, and the first person to hit it would have been a real contributor
// rather than this suite.
//
// So the first test below is the regression test, and it fails against a single-root
// implementation whichever single root is chosen.
// ###########################################################################################
public sealed class SubmissionFileLocatorTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DataRoot => System.IO.Path.Combine(this.thisWorkspace.Root, "Data");

    private string DraftFolder =>
        System.IO.Path.Combine(this.thisWorkspace.Root, "Drafts", "Commodore", "C64", "250407");

    // ###########################################################################################
    // THE REGRESSION TEST: one submission, both roots, in a single resolve pass.
    // ###########################################################################################
    [Fact]
    public void A_submission_resolves_published_files_and_drafted_files_in_the_same_pass()
    {
        // The published board image, living where the sync put it.
        this.thisWorkspace.WriteFile(
            System.IO.Path.Combine("Data", "Commodore", "C64", "250407", "board.png"), "official");

        // The one image the contributor just added, which can only live in the draft - Data/ is
        // sync-owned and would overwrite it.
        this.thisWorkspace.WriteFile(
            System.IO.Path.Combine("Drafts", "Commodore", "C64", "250407", "new-photo.png"),
            "drafted");

        bool foundOfficial = SubmissionFileLocator.TryLocate(
            this.DataRoot, this.DraftFolder,
            "Commodore/C64/250407/board.png", out string officialPath, out _);

        bool foundDrafted = SubmissionFileLocator.TryLocate(
            this.DataRoot, this.DraftFolder,
            "Commodore/C64/250407/new-photo.png", out string draftedPath, out _);

        Assert.True(foundOfficial, "the published file must be found under the data root");
        Assert.True(foundDrafted, "the drafted file must be found under the draft folder");

        Assert.Equal("official", File.ReadAllText(officialPath));
        Assert.Equal("drafted", File.ReadAllText(draftedPath));
    }

    // ###########################################################################################
    // *** THE DRAFTED COPY WINS, AND THIS ASSERTION WAS FLIPPED IN PHASE 6 (2026-09-23). ***
    //
    // The REASON is unchanged and is what matters: what gets uploaded has to be what the app has
    // been showing the contributor on screen, so this must match DraftFileResolver's precedence
    // exactly. Both flipped together.
    //
    // They flipped because a draft stopped being an overlay. It is now a full copy of the board,
    // so a drafted file is a REPLACEMENT for the published one - uploading the published bytes
    // would send the old picture for an image the contributor has already replaced, silently,
    // since both files exist and both are valid.
    // ###########################################################################################
    [Fact]
    public void The_DRAFTED_copy_is_preferred_when_both_roots_hold_the_file()
    {
        this.thisWorkspace.WriteFile(
            System.IO.Path.Combine("Data", "Commodore", "C64", "250407", "board.png"), "official");

        this.thisWorkspace.WriteFile(
            System.IO.Path.Combine("Drafts", "Commodore", "C64", "250407", "board.png"),
            "drafted");

        Assert.True(SubmissionFileLocator.TryLocate(
            this.DataRoot, this.DraftFolder,
            "Commodore/C64/250407/board.png", out string path, out _));

        Assert.Equal("drafted", File.ReadAllText(path));
    }

    // ###########################################################################################
    // CONTAINMENT IS STILL ENFORCED, on both roots. This is the difference between this class and
    // DraftFileResolver: the path resolved here is about to be read and its bytes posted to a
    // public server, so a stored path escaping its root must be refused rather than merely
    // failing to exist. A traversal that happened to land on a real file would otherwise upload
    // it.
    // ###########################################################################################
    [Theory]
    [InlineData("../../../secret.txt")]
    [InlineData("Commodore/../../secret.txt")]
    public void A_path_escaping_its_root_is_refused_rather_than_read(string escaping)
    {
        // A real file at the place the traversal aims for, so the test fails for the right reason:
        // refusal, not absence.
        this.thisWorkspace.WriteFile("secret.txt", "private");

        Assert.False(SubmissionFileLocator.TryLocate(
            this.DataRoot, this.DraftFolder, escaping, out string path, out string reason));

        Assert.Equal(string.Empty, path);
        Assert.NotEqual(string.Empty, reason);
    }

    // A missing file and a refused path are different problems and get different words. Telling
    // someone their file is invalid when it has merely been moved sends them to fix the wrong
    // thing.
    [Fact]
    public void A_file_in_neither_root_is_reported_as_missing_not_as_invalid()
    {
        Assert.False(SubmissionFileLocator.TryLocate(
            this.DataRoot, this.DraftFolder,
            "Commodore/C64/250407/gone.png", out _, out string reason));

        Assert.Contains("not on disk", reason);
    }

    // A draft-only system has no published half at all, so the data root is simply empty. That is
    // the normal state for a brand-new system rather than a failure.
    [Fact]
    public void A_system_with_no_data_root_still_resolves_its_drafted_files()
    {
        this.thisWorkspace.WriteFile(
            System.IO.Path.Combine("Drafts", "Commodore", "C64", "250407", "new-photo.png"),
            "drafted");

        Assert.True(SubmissionFileLocator.TryLocate(
            string.Empty, this.DraftFolder,
            "Commodore/C64/250407/new-photo.png", out string path, out _));

        Assert.Equal("drafted", File.ReadAllText(path));
    }

    // With nowhere to look, the message says so rather than claiming the file is missing - the
    // user's file is not the problem, the app's state is.
    [Fact]
    public void With_neither_root_known_the_reason_says_there_is_nowhere_to_look()
    {
        Assert.False(SubmissionFileLocator.TryLocate(
            string.Empty, string.Empty, "Commodore/C64/250407/board.png", out _, out string reason));

        Assert.Contains("nowhere to look", reason);
    }

    // ###########################################################################################
    // *** THE WRITE SIDE AND THE READ SIDE MUST AGREE, which is the whole point of this test. ***
    //
    // The drafted half used to live under a "Files" subfolder; since Phase 6 it sits at the
    // draft folder's own root with the system's prefix stripped, because a draft folder IS a
    // board folder. The layout moved, the requirement did not: a file written by
    // DraftFileResolver has to be findable by this locator, or a contributor's newly attached
    // photo is silently absent from their submission.
    // ###########################################################################################
    [Fact]
    public void The_drafted_half_is_read_from_exactly_where_DraftFileResolver_writes_it()
    {
        string relative = "Commodore/C64/250407/new-photo.png";

        string writtenTo = DraftFileResolver.BuildDraftFileDestination(this.DraftFolder, relative);

        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(writtenTo)!);
        File.WriteAllText(writtenTo, "drafted");

        Assert.True(SubmissionFileLocator.TryLocate(
            string.Empty, this.DraftFolder, relative, out string locatedAt, out _));

        Assert.Equal(
            System.IO.Path.GetFullPath(writtenTo),
            System.IO.Path.GetFullPath(locatedAt));
    }
    // ###########################################################################################
    // *** THE REPORTED BUG (maintainer, 2026-09-25): "Some files are missing ... HotCPU.png". ***
    // A new image filed in a SHARED folder is written inside the draft under its whole path, and
    // the locator only looked in the draft for the board's own files - so the contributor's
    // attachment was on disk and the submission called it missing.
    // ###########################################################################################
    [Fact]
    public void A_SHARED_file_the_contributor_attached_is_submitted_from_where_it_was_written()
    {
        string relative = "Commodore/Shared files/Component images/HotCPU.png";

        string writtenTo = DraftFileResolver.BuildDraftFileDestination(this.DraftFolder, relative);

        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(writtenTo)!);
        File.WriteAllText(writtenTo, "drafted");

        Assert.True(SubmissionFileLocator.TryLocate(
            this.DataRoot, this.DraftFolder, relative, out string locatedAt, out string reason), reason);

        Assert.Equal(
            System.IO.Path.GetFullPath(writtenTo),
            System.IO.Path.GetFullPath(locatedAt));
    }

    // Containment still holds for the drafted shared copy: a stored path climbing out of the draft
    // is refused, whatever sits where it points.
    [Fact]
    public void A_shared_path_climbing_out_of_the_draft_is_still_refused()
    {
        Assert.False(SubmissionFileLocator.TryLocate(
            this.DataRoot, this.DraftFolder, "Commodore/Shared files/../../../../secret.txt", out _, out string reason));

        Assert.NotEmpty(reason);
    }
}
