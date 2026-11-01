using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // EVERY CHANGE THIS PANEL MAKES TO THE DRAFT: ONE AT A TIME, OFF THE UI THREAD, UNDER THE
    // "PLEASE WAIT" (code review, 2026-10-10).
    //
    // An import - the folder walked, its files copied and taken into the draft, the report read -
    // and a removal each write the draft, up to a couple of seconds on the largest board, and they
    // did it on the UI thread: CRT froze with nothing to say why, against the rule that every wait
    // the user watches runs under BusyOverlay. Now each runs on the pool while the overlay holds the
    // window.
    //
    // *** ONE AT A TIME. *** The overlay takes every NEW click, but a click whose press came first
    // still lands. Two writes of one draft at once would lose one of them, so each change waits for
    // the one before it.
    //
    // *** EACH IS TOLD AS IT ENDS (DraftChanged, 2026-10-30). *** A window told the Contribute tab
    // once it closed (code review, 2026-10-28: closing waited for the queue, or a change still being
    // written was missed). A panel can stay open while the contributor goes to the Schematics tab, so
    // a change that moved the draft is said as soon as it is in - the tab reads the board on screen
    // again then - and closing waits for the queue (DraftChanging) before the panel goes.
    //
    // Part of the KiCadDataPanel partial class - see KiCadDataPanel.axaml.cs for the file map.
    // ###########################################################################################
    public partial class KiCadDataPanel
    {
        // The last change queued; each runs once the one before it has finished.
        private Task thisDraftChanges = Task.CompletedTask;

        // Completes when every change queued so far has finished - what closing waits for, and tests,
        // which wait for it as the user waits for the overlay.
        internal Task DraftChanging => this.thisDraftChanges;

        // A change still running, for a test - the queue waits for `running` as for one of its own.
        internal void HoldChangesForTests(Task running) => this.thisDraftChanges = running;

        // ###########################################################################################
        // Queues one change: it runs once every change queued before it has finished.
        // ###########################################################################################
        private Task<T> ChangeDraftAsync<T>(Func<Task<T>> change)
        {
            Task before = this.thisDraftChanges;
            Task<T> next = RunAfterAsync();
            this.thisDraftChanges = next;

            return next;

            async Task<T> RunAfterAsync()
            {
                try
                {
                    await before;
                }
                catch (Exception ex)
                {
                    // Already told to whoever queued it; it is no reason to drop this one.
                    Logger.Warning($"A change to the draft failed before the next one ran - [{ex.Message}]");
                }

                int madeBefore = this.thisChangesMade;

                try
                {
                    return await change();
                }
                finally
                {
                    // The draft moved: said now, on the UI thread this was queued on.
                    if (this.thisChangesMade != madeBefore)
                    {
                        this.DraftChanged?.Invoke(this, EventArgs.Empty);
                    }
                }
            }
        }

        private Task ChangeDraftAsync(Func<Task> change) =>
            this.ChangeDraftAsync(async () =>
            {
                await change();
                return true;
            });

        // One piece of file work, on the pool, under CRT's "please wait".
        private Task<T> WriteDraftAsync<T>(string message, Func<T> write) =>
            BusyOverlay.RunLocalAsync(this, message, () => Task.Run(write));

        // ###########################################################################################
        // Every change queued so far, ended - however it ended: what closing waits for. A failure was
        // said where the change was made.
        // ###########################################################################################
        internal async Task WaitForChangesAsync()
        {
            try
            {
                await this.thisDraftChanges;
            }
            catch (Exception ex)
            {
                Logger.Warning($"A change to the draft failed while the KiCad data panel was closing - [{ex.Message}]");
            }
        }
    }
}
