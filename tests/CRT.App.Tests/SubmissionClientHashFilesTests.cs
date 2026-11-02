using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What hashing a submission's files says of each, before anything is sent. No network, as hashing goes
// to none (test rule 6).
//
// *** A FILE OF NO BYTES STOPS HERE (owner decision, 2026-11-02: "Refuse empty files"), *** per file, never
// as the server's refusal after an upload began. Through the manifest builder as the Submit dialog uses it
// (SubmitDraftWindow.SubmitAsync; SubmitDraftWindowOutcomeTests holds that wiring), and compared with what
// the server would say: moved on one side only, the last test here fails.
// ###########################################################################################
public sealed class SubmissionClientHashFilesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DraftFolder => Directory.CreateDirectory(Path.Combine(this.thisWorkspace.Root, "Drafts")).FullName;

    private string DataRoot => Directory.CreateDirectory(Path.Combine(this.thisWorkspace.Root, "Data")).FullName;

    private void InDraft(string name, string content) => this.thisWorkspace.WriteFile(Path.Combine("Drafts", name), content);

    private Task<FileHashResult> HashAsync(params string[] names) =>
        SubmissionClient.HashFilesAsync(this.DataRoot, this.DraftFolder, names);

    private static BoardData BoardCiting(params string[] names) => new()
    {
        BoardLocalFiles = names.Select(name => new BoardLocalFileEntry { Category = "Docs", Name = name, File = name }).ToList(),
    };

    private static SubmissionIdentity Identity() => new()
    {
        BoardId = "Commodore/C64/250407",
        Manufacturer = "Commodore",
        Hardware = "C64",
        Board = "250407",
    };

    [Fact]
    public async Task A_file_of_no_bytes_has_no_hash_and_says_why()
    {
        this.InDraft("notes.txt", string.Empty);

        FileHashResult hashed = await this.HashAsync("notes.txt");

        Assert.Empty(hashed.Hashes);
        Assert.Equal(SubmissionFileRules.EmptyFileMessage("notes.txt"), hashed.Problems["notes.txt"]);
    }

    [Fact]
    public async Task A_file_with_bytes_has_its_hash_and_no_problem()
    {
        this.InDraft("notes.txt", "x");

        FileHashResult hashed = await this.HashAsync("notes.txt");

        Assert.Empty(hashed.Problems);
        Assert.Equal(1, hashed.Hashes["notes.txt"].SizeBytes);
    }

    // The manifest is not complete, so nothing is created on the server, and the empty file is said by name
    // while the good one is listed.
    [Fact]
    public async Task A_manifest_with_an_empty_file_is_not_complete_and_says_so_per_file()
    {
        this.InDraft("notes.txt", string.Empty);
        this.InDraft("good.txt", "x");

        FileHashResult hashed = await this.HashAsync("notes.txt", "good.txt");

        SubmissionManifestBuildResult built = SubmissionManifestBuilder.Build(
            BoardCiting("notes.txt", "good.txt"), Identity(), hashed.Hashes, fileProblems: hashed.Problems);

        Assert.False(built.IsComplete);
        Assert.Equal(SubmissionFileRules.EmptyFileMessage("notes.txt"), Assert.Single(built.Problems));
        Assert.Equal("good.txt", Assert.Single(built.Manifest.Files).Path);
    }

    // CRT says of an empty file what the server would refuse it with, each asked for itself.
    [Fact]
    public async Task CRT_says_of_an_empty_file_what_the_server_would_refuse_it_with()
    {
        this.InDraft("notes.txt", string.Empty);

        FileHashResult hashed = await this.HashAsync("notes.txt");

        var manifest = new SubmissionManifest
        {
            BoardId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
            Rows = new SubmissionRows { BoardLocalFiles = BoardCiting("notes.txt").BoardLocalFiles },
            Files = [new SubmissionFile { Path = "notes.txt", Sha256 = SubmissionFileRules.EmptySha256, SizeBytes = 0 }],
        };

        ValidationFinding refused = Assert.Single(SubmissionFileRules.ValidateManifestFiles(manifest, board: null));

        Assert.Equal(refused.Message, hashed.Problems["notes.txt"]);
    }
}
