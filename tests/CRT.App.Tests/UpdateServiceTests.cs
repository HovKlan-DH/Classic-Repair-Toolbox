using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// UpdateService's one pure decision (the class itself is an untested network boundary - CLAUDE.md,
// "Deliberately not covered"): whether a failed update download was the caller stopping it.
//
// The "please wait" overlay cancels a download that makes no progress for two minutes, and the
// banner then says so. That is an expected path, and it was logged as a CRITICAL "Update install
// failed" - so a log sent in support showed a crash where none occurred (code review, 2026-09-29).
// ###########################################################################################
public sealed class UpdateServiceTests
{
    [Fact]
    public void A_cancellation_the_caller_asked_for_is_the_caller_stopping_it()
    {
        using var stop = new CancellationTokenSource();
        stop.Cancel();

        Assert.True(UpdateService.WasStoppedByCaller(new OperationCanceledException(stop.Token), stop.Token));
        Assert.True(UpdateService.WasStoppedByCaller(new TaskCanceledException(), stop.Token));
    }

    // HttpClient reports its OWN timeout as a TaskCanceledException, with the caller's token never
    // cancelled - that is a real failure and stays one.
    [Fact]
    public void A_timeout_the_caller_did_not_ask_for_is_a_failure()
    {
        Assert.False(UpdateService.WasStoppedByCaller(new TaskCanceledException(), CancellationToken.None));
    }

    // Any other fault is a failure, whatever the token says.
    [Fact]
    public void Any_other_exception_is_a_failure_even_after_the_caller_stopped()
    {
        using var stop = new CancellationTokenSource();
        stop.Cancel();

        Assert.False(UpdateService.WasStoppedByCaller(new IOException("disk full"), stop.Token));
    }
}
