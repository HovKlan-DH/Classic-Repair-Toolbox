using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Handlers.DataHandling;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardViewReporter - keeping board views in a file and sending them home (owner: "must work
// robust, and take care of troublesome or no network/internet").
//
// DRIVES THE REAL STATIC STATE through LoadFrom, its test seam, with the SENDING handed in - no
// network (rule 6). NEVER call BoardViewReporter.Load() from a test: it resolves the user's real
// AppData folder. Static state, so the class has a collection of its own.
// ###########################################################################################
[Collection("BoardViewReporter")]
public sealed class BoardViewReporterTests : IDisposable
{
    private static readonly DateTimeOffset Now = DateTimeOffset.UtcNow;

    private static readonly BoardViewMachine Machine = new("CRT 2026.10.0", "Windows", "Microsoft Windows 10.0.19045", "64-bit");

    private readonly TempWorkspace thisWorkspace = new();

    private readonly string thisFile;

    public BoardViewReporterTests()
    {
        this.thisFile = this.thisWorkspace.Path_(BoardViewReporter.FileName);
        BoardViewReporter.LoadFrom(this.thisFile);
    }

    public void Dispose()
    {
        // Leave the static state pointing nowhere, so nothing later writes into a deleted folder.
        BoardViewReporter.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    // A sender answering from a script, and remembering every report it was given.
    private sealed class ScriptedSender(params BoardViewDelivery[] answers)
    {
        private int thisNext;

        public List<BoardViewReport> Sent { get; } = [];

        public Task<BoardViewDelivery> SendAsync(BoardViewReport report, CancellationToken cancellationToken)
        {
            this.Sent.Add(report);
            BoardViewDelivery answer = this.thisNext < answers.Length ? answers[this.thisNext] : BoardViewDelivery.Done;
            this.thisNext++;
            return Task.FromResult(answer);
        }
    }

    private static Task<int> SendAsync(ScriptedSender sender) =>
        BoardViewReporter.SendWaitingAsync(sender.SendAsync, BoardViewReporterTests.Machine, BoardViewReporterTests.Now, CancellationToken.None);

    // A view is in the file the moment it is counted - a CRT closed straight after loses nothing.
    [Fact]
    public void A_recorded_view_is_in_the_file_at_once()
    {
        BoardViewReporter.Record("Commodore/C64/250407", fromBeta: true, BoardViewReporterTests.Now);

        BoardViewOutbox onDisk = BoardViewOutbox.FromJson(File.ReadAllText(this.thisFile));

        BoardView view = Assert.Single(onDisk.Waiting);
        Assert.Equal(("Commodore/C64/250407", true), (view.BoardId, view.FromBeta));
    }

    [Fact]
    public async Task Sent_views_are_finished_and_leave_the_file()
    {
        BoardViewReporter.Record("Commodore/C64/250407", false, BoardViewReporterTests.Now);
        BoardViewReporter.Record("Commodore/VIC-20/250403", false, BoardViewReporterTests.Now);
        var sender = new ScriptedSender(BoardViewDelivery.Done);

        Assert.Equal(1, await BoardViewReporterTests.SendAsync(sender));

        Assert.Equal(2, Assert.Single(sender.Sent).Views!.Count);
        Assert.False(BoardViewReporter.OutboxForTests.HasAnything);
        Assert.False(BoardViewOutbox.FromJson(File.ReadAllText(this.thisFile)).HasAnything);
    }

    // ###########################################################################################
    // *** NO NETWORK: NOTHING IS LOST, AND THE RETRY IS THE SAME BATCH. *** The report stays - in
    // the FILE, so it survives a restart - and the next send carries the same batch id, which the
    // server recognises if the first copy did arrive.
    // ###########################################################################################
    [Fact]
    public async Task A_report_that_could_not_be_sent_is_kept_and_sent_again_as_the_same_batch()
    {
        BoardViewReporter.Record("Commodore/C64/250407", false, BoardViewReporterTests.Now);

        var offline = new ScriptedSender(BoardViewDelivery.TryLater);
        Assert.Equal(0, await BoardViewReporterTests.SendAsync(offline));

        // A restart: read back from the file.
        BoardViewReporter.LoadFrom(this.thisFile);

        var online = new ScriptedSender(BoardViewDelivery.Done);
        Assert.Equal(1, await BoardViewReporterTests.SendAsync(online));

        Assert.Equal(Assert.Single(offline.Sent).BatchId, Assert.Single(online.Sent).BatchId);
        Assert.False(BoardViewReporter.OutboxForTests.HasAnything);
    }

    // ###########################################################################################
    // HasWaiting is what Main's retry timer asks every BoardViewRetryInterval: true while a view
    // waits - counted and not yet sent, or sent without an answer that finished it - and false
    // once it is delivered, so a CRT with nothing to send makes no request at all.
    // ###########################################################################################
    [Fact]
    public async Task Views_wait_until_a_send_delivers_them()
    {
        Assert.False(BoardViewReporter.HasWaiting);

        BoardViewReporter.Record("Commodore/C64/250407", false, BoardViewReporterTests.Now);
        Assert.True(BoardViewReporter.HasWaiting);

        await BoardViewReporterTests.SendAsync(new ScriptedSender(BoardViewDelivery.TryLater));
        Assert.True(BoardViewReporter.HasWaiting);

        await BoardViewReporterTests.SendAsync(new ScriptedSender(BoardViewDelivery.Done));
        Assert.False(BoardViewReporter.HasWaiting);
    }

    // A backlog goes as several reports in one send, and the first "not now" stops the rest.
    [Fact]
    public async Task A_backlog_goes_report_after_report_until_one_is_not_accepted()
    {
        for (int index = 0; index < (2 * BoardViewRules.MaxViewsPerReport) + 1; index++)
            BoardViewReporter.Record($"Commodore/C64/{index}", false, BoardViewReporterTests.Now);

        var sender = new ScriptedSender(BoardViewDelivery.Done, BoardViewDelivery.TryLater);

        Assert.Equal(1, await BoardViewReporterTests.SendAsync(sender));
        Assert.Equal(2, sender.Sent.Count);

        BoardViewOutbox left = BoardViewReporter.OutboxForTests;
        Assert.Equal(sender.Sent[1].BatchId, left.Sending!.BatchId);
        Assert.Single(left.Waiting);
    }

    // Nothing waiting: nothing sent.
    [Fact]
    public async Task With_nothing_waiting_nothing_is_sent()
    {
        var sender = new ScriptedSender();

        Assert.Equal(0, await BoardViewReporterTests.SendAsync(sender));
        Assert.Empty(sender.Sent);
    }

    // A blank board is not a view.
    [Fact]
    public void A_blank_board_is_not_recorded()
    {
        BoardViewReporter.Record("   ", false, BoardViewReporterTests.Now);

        Assert.False(BoardViewReporter.OutboxForTests.HasAnything);
    }

    // ###########################################################################################
    // One send at a time: a second send starting while one waits on the network does nothing - it
    // would otherwise send the same report twice at once.
    // ###########################################################################################
    [Fact]
    public async Task A_send_started_while_one_is_running_does_nothing()
    {
        BoardViewReporter.Record("Commodore/C64/250407", false, BoardViewReporterTests.Now);

        var gate = new TaskCompletionSource<BoardViewDelivery>();
        int calls = 0;

        Task<int> first = BoardViewReporter.SendWaitingAsync(
            (report, _) => { Interlocked.Increment(ref calls); return gate.Task; },
            BoardViewReporterTests.Machine, BoardViewReporterTests.Now, CancellationToken.None);

        int second = await BoardViewReporter.SendWaitingAsync(
            (report, _) => { Interlocked.Increment(ref calls); return Task.FromResult(BoardViewDelivery.Done); },
            BoardViewReporterTests.Machine, BoardViewReporterTests.Now, CancellationToken.None);

        gate.SetResult(BoardViewDelivery.Done);

        Assert.Equal(0, second);
        Assert.Equal(1, await first);
        Assert.Equal(1, calls);
    }

    // A file that will not read starts empty, and the application carries on.
    [Fact]
    public void An_unreadable_file_starts_empty()
    {
        File.WriteAllText(this.thisFile, "{ not json");

        BoardViewReporter.LoadFrom(this.thisFile);

        Assert.False(BoardViewReporter.OutboxForTests.HasAnything);
    }
}
