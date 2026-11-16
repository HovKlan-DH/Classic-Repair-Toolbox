using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DeferredTextFile - the settings file written a moment after it changes, off the UI thread
// (2026-11-16, performance walkthrough: a board switch wrote it whole one to four times on the UI
// thread). The writes are counted by the write the tests hand in, with a delay of an hour where a
// test must see what waits, so no slow machine writes it early.
// ###########################################################################################
public sealed class DeferredTextFileTests
{
    private readonly List<(string Path, string Text)> thisWritten = [];

    private DeferredTextFile Writer(TimeSpan delay) =>
        new(delay, (path, text) =>
        {
            lock (this.thisWritten)
            {
                this.thisWritten.Add((path, text));
            }

            return true;
        });

    [Fact]
    public void Texts_queued_in_the_moment_are_written_once_the_latest_of_them()
    {
        DeferredTextFile writer = this.Writer(TimeSpan.FromHours(1));

        writer.Queue("settings.json", "first");
        writer.Queue("settings.json", "second");
        writer.Queue("settings.json", "third");

        Assert.True(writer.IsPending);
        Assert.Empty(this.thisWritten);

        Assert.True(writer.Flush());

        Assert.Equal([("settings.json", "third")], this.thisWritten);
        Assert.False(writer.IsPending);
    }

    [Fact]
    public void The_text_written_last_is_not_written_again()
    {
        DeferredTextFile writer = this.Writer(TimeSpan.FromHours(1));

        writer.Queue("settings.json", "same");
        writer.Flush();
        writer.Queue("settings.json", "same");
        writer.Flush();

        Assert.Single(this.thisWritten);
        Assert.Equal(1, writer.WritesForTests);

        // Unless the file was read again since: it may hold what somebody typed into it.
        writer.ForgetWhatWasWritten();
        writer.Queue("settings.json", "same");
        writer.Flush();

        Assert.Equal(2, this.thisWritten.Count);
    }

    [Fact]
    public void The_same_text_for_another_file_is_written()
    {
        DeferredTextFile writer = this.Writer(TimeSpan.FromHours(1));

        writer.Queue("one.json", "text");
        writer.Flush();
        writer.Queue("two.json", "text");
        writer.Flush();

        Assert.Equal([("one.json", "text"), ("two.json", "text")], this.thisWritten);
    }

    // A write that did not reach the disk is no "written last": the same text queued again is tried again.
    [Fact]
    public void A_text_that_could_not_be_written_is_tried_again_when_queued_again()
    {
        int tries = 0;
        var writer = new DeferredTextFile(TimeSpan.FromHours(1), (_, _) => ++tries > 1);

        writer.Queue("settings.json", "text");
        Assert.False(writer.Flush());

        writer.Queue("settings.json", "text");
        Assert.True(writer.Flush());

        Assert.Equal(2, tries);
    }

    [Fact]
    public void Nothing_waiting_writes_nothing()
    {
        DeferredTextFile writer = this.Writer(TimeSpan.FromHours(1));

        Assert.True(writer.Flush());
        Assert.Empty(this.thisWritten);
    }

    [Fact]
    public async Task What_waits_is_written_by_itself_after_the_delay()
    {
        DeferredTextFile writer = this.Writer(TimeSpan.FromMilliseconds(20));

        writer.Queue("settings.json", "text");

        DateTime giveUp = DateTime.UtcNow.AddSeconds(10);

        while (writer.IsPending && DateTime.UtcNow < giveUp)
        {
            await Task.Delay(10);
        }

        // Taken off the queue just before it is written: wait for the write itself.
        while (writer.WritesForTests == 0 && DateTime.UtcNow < giveUp)
        {
            await Task.Delay(10);
        }

        Assert.Equal(1, writer.WritesForTests);
        Assert.Equal([("settings.json", "text")], this.thisWritten);
    }

    // ###########################################################################################
    // *** WRITES NEVER GO BACKWARDS. *** Each write takes the latest text under the one write lock, so
    // texts queued from many threads while the timer and Flush both write end with the last one queued.
    // ###########################################################################################
    [Fact]
    public async Task Texts_queued_from_many_threads_end_with_the_last_one_queued()
    {
        DeferredTextFile writer = this.Writer(TimeSpan.FromMilliseconds(1));

        await Task.WhenAll(Enumerable.Range(0, 8).Select(thread => Task.Run(() =>
        {
            for (int i = 0; i < 200; i++)
            {
                writer.Queue("settings.json", $"{thread}-{i}");

                if (i % 17 == 0)
                {
                    writer.Flush();
                }
            }
        })));

        writer.Queue("settings.json", "last");
        writer.Flush();

        lock (this.thisWritten)
        {
            Assert.Equal("last", this.thisWritten[^1].Text);
        }
    }
}
