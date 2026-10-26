using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// Covers FileReplacer - writing a board file by replacing it, never by opening the old one
// (owner report, 2026-09-28: production's files copied into BETA by hand as root, and the
// approval that followed was refused half-way through the board at the one file it opened for
// writing).
// ###########################################################################################
public sealed class FileReplacerTests : IDisposable
{
    private readonly string thisRoot;

    public FileReplacerTests()
    {
        this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-file-replacer", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(this.thisRoot);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(this.thisRoot, recursive: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A leftover temp folder is harmless.
        }
    }

    private string Target => Path.Combine(this.thisRoot, "Data C128 310378 v2.0.0.json");

    // ###########################################################################################
    // *** A TARGET HELD OPEN FOR A MOMENT IS WAITED FOR (2026-10-25). *** On Windows a file open for
    // reading cannot be replaced, and CRT reads its drafts' files on the UI thread (the Contribute
    // tab's list, CRT's window coming to the front) while a draft saves itself off it - so a save now
    // and then failed with the file "being used by another process" (found by a test that failed
    // in some orders only). On Linux the rename succeeds anyway; the test then simply passes.
    // ###########################################################################################
    [Fact]
    public async Task A_target_held_open_for_a_moment_is_replaced_once_it_is_let_go()
    {
        FileReplacer.WriteAllText(this.Target, "first");

        var reader = new FileStream(this.Target, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        Task release = Task.Run(async () =>
        {
            await Task.Delay(60);
            await reader.DisposeAsync();
        });

        try
        {
            // On a thread of its own, as a draft's save runs - the reader lets go meanwhile.
            await Task.Run(() => FileReplacer.WriteAllText(this.Target, "second"));
        }
        finally
        {
            await release;
        }

        Assert.Equal("second", await File.ReadAllTextAsync(this.Target));
    }

    // The rename is tried again while it is refused, a short pause apart - and gives up with what it
    // threw after the last try, so a folder that may not be written still fails.
    [Fact]
    public void The_rename_is_tried_again_while_refused_and_gives_up_after_the_last_try()
    {
        int moves = 0;
        var pauses = new List<TimeSpan>();

        FileReplacer.MoveOver("from", "to", (_, _) =>
        {
            if (++moves < 3)
                throw new IOException("used by another process");
        }, pauses.Add);

        Assert.Equal(3, moves);
        Assert.Equal([FileReplacer.MovePause, FileReplacer.MovePause], pauses);

        moves = 0;
        Assert.Throws<UnauthorizedAccessException>(() =>
            FileReplacer.MoveOver("from", "to", (_, _) => { moves++; throw new UnauthorizedAccessException("denied"); }, _ => { }));
        Assert.Equal(FileReplacer.MoveAttempts, moves);
    }

    // Anything else is not waited for.
    [Fact]
    public void Another_failure_is_not_tried_again()
    {
        int moves = 0;

        Assert.Throws<ArgumentException>(() =>
            FileReplacer.MoveOver("from", "to", (_, _) => { moves++; throw new ArgumentException("bad path"); }, _ => { }));
        Assert.Equal(1, moves);
    }

    [Fact]
    public void A_new_file_is_written_and_an_existing_one_replaced()
    {
        FileReplacer.WriteAllText(this.Target, "first");
        Assert.Equal("first", File.ReadAllText(this.Target));

        FileReplacer.WriteAllText(this.Target, "second");
        Assert.Equal("second", File.ReadAllText(this.Target));
    }

    // The temporary is renamed over the target, never left beside it - a dot-name is never
    // published, but litter is litter.
    [Fact]
    public void No_temporary_file_is_left_behind()
    {
        File.WriteAllText(this.Target, "old");

        FileReplacer.WriteAllText(this.Target, "new");

        Assert.Equal([this.Target], Directory.GetFiles(this.thisRoot));
    }

    // A write that fails leaves the old file exactly as it was, and no temporary either - the
    // point of writing beside the target first.
    [Fact]
    public void A_write_that_fails_leaves_the_old_file_untouched()
    {
        File.WriteAllText(this.Target, "old");

        Assert.Throws<IOException>(() => FileReplacer.Replace(this.Target, temporary =>
        {
            File.WriteAllText(temporary, "half");
            throw new IOException("disk full");
        }));

        Assert.Equal("old", File.ReadAllText(this.Target));
        Assert.Equal([this.Target], Directory.GetFiles(this.thisRoot));
    }

    // Beside the target (a rename never crosses a file system) and a dot-name (never published,
    // promoted or listed - the checksum manifest skips dot-segments).
    [Fact]
    public void The_temporary_is_a_dot_name_beside_the_target()
    {
        string temporary = FileReplacer.TemporaryPathFor(this.Target);

        Assert.Equal(this.thisRoot, Path.GetDirectoryName(temporary));
        Assert.StartsWith(".", Path.GetFileName(temporary), StringComparison.Ordinal);
        Assert.NotEqual(temporary, FileReplacer.TemporaryPathFor(this.Target));
    }

    // ###########################################################################################
    // *** THE REASON IT EXISTS. *** A file copied in by hand as root may be replaced (its
    // folder's permission) but not opened for writing (its own). Made here as a file with no write
    // permission for its owner, which refuses an open the same way. Not on Windows (no such mode;
    // a read-only file there refuses a rename too), and not as root (which ignores it). Fails
    // against File.WriteAllText.
    // ###########################################################################################
    [Fact]
    public void A_file_that_may_not_be_opened_for_writing_is_still_replaced()
    {
        // The return is for the platform analyzer, which cannot see that Skip throws.
        if (OperatingSystem.IsWindows())
        {
            Assert.Skip("Unix permissions only.");
            return;
        }
        Assert.SkipWhen(Environment.UserName == "root", "root may write anything.");

        File.WriteAllText(this.Target, "copied in by hand");
        File.SetUnixFileMode(this.Target, UnixFileMode.UserRead | UnixFileMode.GroupRead | UnixFileMode.OtherRead);

        FileReplacer.WriteAllText(this.Target, "published");

        Assert.Equal("published", File.ReadAllText(this.Target));
    }
}
