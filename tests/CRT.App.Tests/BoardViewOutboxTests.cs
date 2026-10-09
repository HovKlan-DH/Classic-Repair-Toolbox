using System;
using System.Linq;
using Handlers.DataHandling;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardViewOutbox - the views waiting to be sent, and the one report being sent (owner: "must work
// robust, and take care of troublesome or no network/internet").
//
// THE RULE THAT MATTERS MOST: a report is sent again UNCHANGED until an answer finishes it - the
// same batch id and the same views - because that is what lets the server recognise a retry and
// never count a view twice. Views counted meanwhile wait behind it.
// ###########################################################################################
public sealed class BoardViewOutboxTests
{
    private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

    private static readonly BoardViewMachine Machine = new("CRT 2026.10.0", "Windows", "Microsoft Windows 10.0.19045", "64-bit");

    private static BoardView View(string boardId = "Commodore/C64/250407", double minutesAgo = 1) =>
        new(boardId, BoardViewOutboxTests.Now.AddMinutes(-minutesAgo), false);

    [Fact]
    public void Waiting_views_become_one_report_carrying_the_machine()
    {
        var outbox = new BoardViewOutbox();
        outbox.Add(BoardViewOutboxTests.View(minutesAgo: 3));
        outbox.Add(BoardViewOutboxTests.View("Commodore/VIC-20/250403", minutesAgo: 2));

        BoardViewReport report = outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine)!;

        Assert.Equal(2, report.Views!.Count);
        Assert.Equal(("CRT 2026.10.0", "Windows", "Microsoft Windows 10.0.19045", "64-bit"), (report.Version, report.OsHighlevel, report.OsVersion, report.Cpu));
        Assert.True(Guid.TryParse(report.BatchId, out _));
        Assert.Empty(outbox.Waiting);
        Assert.Same(report, outbox.Sending);
    }

    // ###########################################################################################
    // *** UNANSWERED, IT IS SENT AGAIN UNCHANGED. *** Same batch id, same views - new views wait
    // behind it and go in the next report, with a new id.
    // ###########################################################################################
    [Fact]
    public void An_unanswered_report_is_sent_again_unchanged_and_new_views_wait_behind_it()
    {
        var outbox = new BoardViewOutbox();
        outbox.Add(BoardViewOutboxTests.View());

        BoardViewReport first = outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine)!;
        outbox.Add(BoardViewOutboxTests.View("Commodore/VIC-20/250403"));

        BoardViewReport again = outbox.NextReport(BoardViewOutboxTests.Now.AddMinutes(5), BoardViewOutboxTests.Machine)!;

        Assert.Equal(first.BatchId, again.BatchId);
        Assert.Equal(first.Views!, again.Views!);
        Assert.Single(outbox.Waiting);

        outbox.Finished(first.BatchId);
        BoardViewReport next = outbox.NextReport(BoardViewOutboxTests.Now.AddMinutes(6), BoardViewOutboxTests.Machine)!;

        Assert.NotEqual(first.BatchId, next.BatchId);
        Assert.Equal("Commodore/VIC-20/250403", Assert.Single(next.Views!).BoardId);
    }

    // Finished only by its own batch id - a late answer to an older report finishes nothing.
    [Fact]
    public void Only_the_answer_to_the_report_being_sent_finishes_it()
    {
        var outbox = new BoardViewOutbox();
        outbox.Add(BoardViewOutboxTests.View());
        BoardViewReport report = outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine)!;

        outbox.Finished(Guid.NewGuid().ToString("N"));
        outbox.Finished(null);
        Assert.Same(report, outbox.Sending);

        outbox.Finished(report.BatchId);
        Assert.Null(outbox.Sending);
        Assert.False(outbox.HasAnything);
    }

    // A backlog goes as several reports, the oldest views first.
    [Fact]
    public void A_backlog_goes_in_reports_of_the_most_one_may_carry_oldest_first()
    {
        var outbox = new BoardViewOutbox();

        for (int index = 0; index < BoardViewRules.MaxViewsPerReport + 5; index++)
            outbox.Add(new BoardView($"Commodore/C64/{index}", BoardViewOutboxTests.Now.AddMinutes(-1000 + index), false));

        BoardViewReport first = outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine)!;
        Assert.Equal(BoardViewRules.MaxViewsPerReport, first.Views!.Count);
        Assert.Equal("Commodore/C64/0", first.Views[0].BoardId);

        outbox.Finished(first.BatchId);
        Assert.Equal(5, outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine)!.Views!.Count);
    }

    // ###########################################################################################
    // What the server would not count is let go: waiting views older than MaxAge, and a report being
    // sent whose views have all grown too old - never sent again only to be refused.
    // ###########################################################################################
    [Fact]
    public void Views_too_old_to_count_are_let_go()
    {
        var outbox = new BoardViewOutbox();
        outbox.Add(new BoardView("Commodore/C64/250407", BoardViewOutboxTests.Now - BoardViewRules.MaxAge - TimeSpan.FromHours(1), false));

        Assert.Null(outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine));
        Assert.Empty(outbox.Waiting);

        outbox.Add(BoardViewOutboxTests.View());
        BoardViewReport report = outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine)!;

        Assert.Null(outbox.NextReport(BoardViewOutboxTests.Now + BoardViewRules.MaxAge + TimeSpan.FromDays(1), BoardViewOutboxTests.Machine));
        Assert.Null(outbox.Sending);
        Assert.NotNull(report);
    }

    // A machine offline for months keeps at most MaxWaitingViews - the newest.
    [Fact]
    public void At_most_the_newest_max_waiting_views_are_kept()
    {
        var outbox = new BoardViewOutbox();

        for (int index = 0; index < BoardViewRules.MaxWaitingViews + 3; index++)
            outbox.Add(new BoardView($"Commodore/C64/{index}", BoardViewOutboxTests.Now, false));

        Assert.Equal(BoardViewRules.MaxWaitingViews, outbox.Waiting.Count);
        Assert.Equal("Commodore/C64/3", outbox.Waiting[0].BoardId);
    }

    // The file: a report being sent survives a restart WITH its batch id, so the retry after a
    // restart is still recognised by the server.
    [Fact]
    public void The_file_keeps_the_waiting_views_and_the_report_being_sent_with_its_batch_id()
    {
        var outbox = new BoardViewOutbox();
        outbox.Add(BoardViewOutboxTests.View());
        BoardViewReport sending = outbox.NextReport(BoardViewOutboxTests.Now, BoardViewOutboxTests.Machine)!;
        outbox.Add(BoardViewOutboxTests.View("Commodore/VIC-20/250403"));

        BoardViewOutbox read = BoardViewOutbox.FromJson(outbox.ToJson());

        Assert.Equal(sending.BatchId, read.Sending!.BatchId);
        Assert.Equal(sending.Views!, read.Sending.Views!);
        Assert.Equal(outbox.Waiting, read.Waiting);
        Assert.DoesNotContain("hasAnything", outbox.ToJson(), StringComparison.OrdinalIgnoreCase);
    }

    // ###########################################################################################
    // *** A VIEW AN OLDER CRT WROTE IS STILL SENT (owner decision, 2026-10-09). *** Until "system"
    // became "board" a view named its board "systemId", and an outbox file written then may still
    // be waiting on disk. Its views carry on under the new name.
    // ###########################################################################################
    [Fact]
    public void A_view_written_as_systemId_before_the_rename_is_read_and_sent_as_boardId()
    {
        BoardViewOutbox read = BoardViewOutbox.FromJson(
            """{"waiting":[{"systemId":"Commodore/C64/250407","viewedUtc":"2026-09-27T12:00:00+00:00","fromBeta":false}]}""");

        BoardView view = Assert.Single(read.Waiting);
        Assert.Equal("Commodore/C64/250407", view.BoardId);
        Assert.Null(view.SystemId);
        Assert.DoesNotContain("systemId", read.ToJson(), StringComparison.Ordinal);
    }

    // A file that will not read costs its views, never the application.
    [Theory]
    [InlineData("")]
    [InlineData("not json")]
    [InlineData("[]")]
    [InlineData("null")]
    [InlineData("""{"waiting":[{"boardId":null,"viewedUtc":"2026-09-27T12:00:00+00:00","fromBeta":false}]}""")]
    public void A_file_that_will_not_read_is_an_empty_outbox(string json)
    {
        BoardViewOutbox read = BoardViewOutbox.FromJson(json);

        Assert.False(read.HasAnything);
    }
}
