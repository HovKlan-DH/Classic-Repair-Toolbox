using Avalonia;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The "Schematics" tab's pictures, read once and the last few kept (PictureReads, 2026-11-16, performance walkthrough):
// a picture was decoded whole every time it was shown - back and forth between a board's top and bottom, each switch -
// and on a board switch only after every thumbnail. Tested on a plain picture: what matters is which reads are made,
// shared, kept and let go of.
// ###########################################################################################
public sealed class PictureReadsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private sealed class Picture(string path)
    {
        public string Path { get; } = path;

        public bool Disposed { get; set; }
    }

    private readonly List<Picture> thisDisposed = [];

    private PictureReads<Picture> NewReads(int kept = 3) =>
        new(kept, path => (new Picture(path), new Thickness(1)), picture =>
        {
            lock (this.thisDisposed)
            {
                picture.Disposed = true;
                this.thisDisposed.Add(picture);
            }
        });

    private string File_(string name) => this.thisWorkspace.WriteFile(name, name);

    private static Picture Done(Task<(Picture? Picture, Thickness Border)> read)
    {
        Assert.True(read.Wait(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken), "Not read in time");
        return read.Result.Picture!;
    }

    // The board load begins reading the picture it will show; the selection takes that very read.
    [Fact]
    public async Task A_picture_asked_for_again_while_its_file_stays_as_it_is_is_the_same_read()
    {
        var reads = NewReads();
        string top = File_("top.png");

        Task<(Picture? Picture, Thickness Border)> first = reads.Read(top);

        Assert.Same(first, reads.Read(top));
        Assert.Equal(1, reads.ReadsStartedForTests);
        Assert.Equal(new Thickness(1), (await first).Border);
    }

    [Fact]
    public void A_picture_whose_file_changed_is_read_again()
    {
        var reads = NewReads();
        string top = File_("top.png");
        Picture before = Done(reads.Read(top));

        File.WriteAllText(top, "replaced, longer");
        File.SetLastWriteTimeUtc(top, File.GetLastWriteTimeUtc(top).AddMinutes(1));

        Assert.NotSame(before, Done(reads.Read(top)));
        Assert.Equal(2, reads.ReadsStartedForTests);
    }

    // Going back and forth decodes nothing; a picture past the newest kept is let go of - and disposed - once read.
    [Fact]
    public void The_newest_reads_are_kept_and_one_past_them_is_disposed()
    {
        var reads = NewReads(kept: 2);
        string a = File_("a.png"), b = File_("b.png"), c = File_("c.png");

        Picture pictureA = Done(reads.Read(a));
        reads.KeepOnly(pictureA);
        Picture pictureB = Done(reads.Read(b));
        reads.KeepOnly(pictureB);
        Done(reads.Read(a));
        reads.KeepOnly(pictureA);

        Assert.Equal(2, reads.ReadsStartedForTests);
        Assert.Empty(this.thisDisposed);

        Picture pictureC = Done(reads.Read(c));
        reads.KeepOnly(pictureC);

        SpinWait.SpinUntil(() => pictureB.Disposed, TimeSpan.FromSeconds(10));

        Assert.True(pictureB.Disposed);
        Assert.False(pictureA.Disposed);
        Assert.False(pictureC.Disposed);
        Assert.Equal(2, reads.HeldForTests);
    }

    // ###########################################################################################
    // *** A read that came to nothing is not kept (code review, 2026-11-16). *** A picture held a moment by a virus scanner as it
    // was shown stayed "missing" until three others were read; the next ask reads it again.
    // ###########################################################################################
    [Fact]
    public async Task A_read_that_came_to_nothing_is_read_again_at_the_next_ask()
    {
        int asked = 0;
        var reads = new PictureReads<Picture>(3, path => Interlocked.Increment(ref asked) == 1 ? (null, default) : (new Picture(path), default), _ => { });
        string top = File_("top.png");

        (Picture? Picture, Thickness Border) first = await reads.Read(top).WaitAsync(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken);
        Assert.Null(first.Picture);
        SpinWait.SpinUntil(() => reads.HeldForTests == 0, TimeSpan.FromSeconds(10));

        Assert.NotNull(Done(reads.Read(top)));
        Assert.Equal(2, reads.DecodesMadeForTests);
    }

    // ###########################################################################################
    // *** Boards picked one after another decode at most two pictures at once, the newest first, and one overtaken before its
    // turn is never decoded (code review, 2026-11-16). *** Every board load begins its picture's read; eight boards picked in a
    // row began eight decodes of some 50 MB, all at once, sharing the CPU with the board actually chosen.
    // ###########################################################################################
    [Fact]
    public async Task Reads_overtaken_before_their_turn_are_never_decoded_and_two_decode_at_once()
    {
        using var hold = new ManualResetEventSlim(false);
        var decoded = new List<string>();
        int atOnce = 0, mostAtOnce = 0;

        var reads = new PictureReads<Picture>(3, path =>
        {
            int now = Interlocked.Increment(ref atOnce);
            lock (decoded)
            {
                mostAtOnce = Math.Max(mostAtOnce, now);
                decoded.Add(Path.GetFileName(path));
            }

            hold.Wait(TimeSpan.FromSeconds(10));
            Interlocked.Decrement(ref atOnce);
            return (new Picture(path), default);
        }, _ => { });

        List<Task<(Picture? Picture, Thickness Border)>> asked = Enumerable.Range(1, 8).Select(index => reads.Read(File_($"p{index}.png"))).ToList();

        // p3, p4 and p5 waited, and were pushed past the newest three before their turn: come to nothing at once.
        foreach (Task<(Picture? Picture, Thickness Border)> overtaken in asked.Skip(2).Take(3))
        {
            Assert.True(overtaken.IsCompleted);
            Assert.Null((await overtaken).Picture);
        }

        hold.Set();
        (Picture? Picture, Thickness Border)[] all = await Task.WhenAll(asked).WaitAsync(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken);

        Assert.Equal(["p1.png", "p2.png", "p6.png", "p7.png", "p8.png"], decoded.Order(StringComparer.Ordinal));
        Assert.Equal(PictureReads<Picture>.DecodedAtOnce, mostAtOnce);
        Assert.NotNull(all[^1].Picture);
    }

    // The picture shown is never let go of, however old its read.
    [Fact]
    public void The_picture_shown_is_never_let_go_of()
    {
        var reads = NewReads(kept: 1);
        string a = File_("a.png"), b = File_("b.png");

        Picture pictureA = Done(reads.Read(a));
        Done(reads.Read(b));

        reads.KeepOnly(pictureA);

        Assert.False(pictureA.Disposed);
        Assert.Equal(2, reads.HeldForTests);
    }
}
