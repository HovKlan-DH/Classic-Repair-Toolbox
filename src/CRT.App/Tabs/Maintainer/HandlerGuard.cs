using System;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // *** AN EVENT HANDLER OF THE MAINTAINER TAB NEVER ENDS CRT (code review, 2026-10-13). ***
    //
    // Every click, key and selection on this tab is an `async void` handler, and an exception
    // thrown after its first await has nowhere to go but the dispatcher - where App lets it end
    // CRT, on purpose (a crash in the UI is not to be papered over in general). The server's
    // answers are values, never exceptions (ReviewApiClient), but the code around them is not
    // proof against everything: a control in an unexpected state, a dialog, a file. So every async
    // handler under Tabs/Maintainer runs its work through RunAsync: anything thrown is written to
    // CRT's log with the handler's name and the whole exception, and said in a sentence where that
    // screen shows its errors (`show`) - the tab stays, and so does CRT.
    //
    // UI code, beside the handlers it guards: it writes CRT's log, which the shared controls in
    // Controls/ may not. Its sentence is MaintainerWaitWording.Unexpected, tested there.
    // ###########################################################################################
    internal static class HandlerGuard
    {
        public static async Task RunAsync(Func<Task> work, Action<string>? show, [CallerMemberName] string handler = "")
        {
            try
            {
                await work();
            }
            catch (Exception exception)
            {
                Logger.Critical($"Maintainer tab: {handler} failed - {exception}");

                try
                {
                    show?.Invoke(MaintainerWaitWording.Unexpected(exception.Message));
                }
                catch (Exception)
                {
                    // Saying it failed must not be the next thing to end CRT - it is in the log.
                }
            }
        }
    }
}
