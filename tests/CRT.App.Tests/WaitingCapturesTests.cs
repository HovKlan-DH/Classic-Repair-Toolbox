using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// WaitingCaptures - the oscilloscope captures an earlier run left waiting go, once a run (code review, 2026-11-21: one CRT ending
// with captures not accepted left them in the user's folder for good). Only a capture's own name, only one older than the age
// given - another CRT running may have some waiting now - and nothing else in the folder.
// ###########################################################################################
public sealed class WaitingCapturesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string Captures => Path.Combine(this.thisWorkspace.Root, "Scope captures waiting");

    private string Write(string name, DateTime writtenUtc)
    {
        Directory.CreateDirectory(this.Captures);
        string path = Path.Combine(this.Captures, name);
        File.WriteAllBytes(path, [1, 2, 3]);
        File.SetLastWriteTimeUtc(path, writtenUtc);
        return path;
    }

    [Fact]
    public void Only_captures_older_than_the_age_go_and_nothing_else_in_the_folder()
    {
        var now = new DateTime(2026, 11, 21, 12, 0, 0, DateTimeKind.Utc);
        string old = this.Write("0123456789abcdef0123456789abcdef.png", now.AddDays(-2));
        string recent = this.Write("fedcba9876543210fedcba9876543210.png", now.AddHours(-1));
        string other = this.Write("notes.png", now.AddDays(-30));
        string notPng = this.Write("0123456789abcdef0123456789abcdee.txt", now.AddDays(-30));

        Assert.Equal(1, WaitingCaptures.Sweep(this.Captures, now, WaitingCaptures.SweptAfter));

        Assert.False(File.Exists(old));
        Assert.True(File.Exists(recent));
        Assert.True(File.Exists(other));
        Assert.True(File.Exists(notPng));
    }

    [Fact]
    public void A_folder_not_there_or_blank_sweeps_nothing()
    {
        Assert.Equal(0, WaitingCaptures.Sweep(this.Captures, DateTime.UtcNow, WaitingCaptures.SweptAfter));
        Assert.Equal(0, WaitingCaptures.Sweep(" ", DateTime.UtcNow, WaitingCaptures.SweptAfter));
    }
}
