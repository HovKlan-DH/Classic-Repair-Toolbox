using System;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // Server calls the maintainer waits for, under the window's BusyOverlay (owner decision,
    // 2026-09-28: "I want this method everywhere in the entire project where there is a Wait").
    //
    // *** THE TWO-MINUTE LIMIT COMES BACK AS AN ORDINARY FAILED RESULT *** (ReviewApiFailure
    // .TimedOut, with WaitWording's sentence), so every existing `if (!result.IsOk)` path already
    // handles it. A caller whose call CHANGED something tests for TimedOut as well and looks at the
    // server again before saying what happened - see WaitWording.AfterTimeout.
    // ###########################################################################################
    internal static class ServerWait
    {
        // One request. The overlay's token is the request's, so the request stops at the limit.
        public static async Task<ReviewApiResult<T>> CallAsync<T>(
            StyledElement anchor,
            string message,
            Func<CancellationToken, Task<ReviewApiResult<T>>> call)
            where T : class
        {
            ArgumentNullException.ThrowIfNull(call);

            WaitResult<ReviewApiResult<T>> waited =
                await BusyOverlay.RunAsync(anchor, message, context => call(context.Token));

            return waited.IsTimedOut || waited.Value is null
                ? ReviewApiResult<T>.Failed(ReviewApiFailure.TimedOut, WaitWording.NoAnswer)
                : waited.Value;
        }

        // Anything else - several requests, a list read again. False when the limit passed.
        public static async Task<bool> RunAsync(StyledElement anchor, string message, Func<Task> work)
        {
            ArgumentNullException.ThrowIfNull(work);

            WaitResult<bool> waited = await BusyOverlay.RunAsync(anchor, message, _ => work());

            return !waited.IsTimedOut;
        }
    }
}
