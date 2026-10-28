using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // EVERY CHANGE THIS WINDOW MAKES TO THE DRAFT: ONE AT A TIME, OFF THE UI THREAD, UNDER THE
    // "PLEASE WAIT" (code review, 2026-10-10).
    //
    // An import - the folder walked, its files copied and taken into the draft, the report read -
    // and a removal each write the draft, up to a couple of seconds on the largest board, and they
    // did it on the UI thread: the window froze with nothing to say why, against the rule that every
    // wait the user watches runs under BusyOverlay. Now each runs on the pool while the overlay holds
    // the window.
    //
    // *** ONE AT A TIME. *** The overlay takes every NEW click, but a click whose press came first
    // still lands. Two writes of one draft at once would lose one of them, so each change waits for
    // the one before it.
    //
    // *** CLOSING WAITS FOR THEM (OnClosing). *** The Contribute tab looks at ChangedTheDraft once the
    // window has closed, so a change still being written then would be missed - and its files would
    // go into the draft with nothing on screen reading them (code review, 2026-10-28).
    //
    // Part of the BoardFilesWindow partial class - see BoardFilesWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardFilesWindow
    {
        // The last change queued; each runs once the one before it has finished.
        private Task thisDraftChanges = Task.CompletedTask;

        // Set once the window may close: everything queued is done.
        private bool thisReadyToClose;

        // The wait-then-close under way, if any - a second close waits for the same one.
        private Task? thisClosingAfterChanges;

        // Completes when every change queued so far has finished - for tests, which wait for it as
        // the user waits for the overlay.
        internal Task DraftChanging => this.thisDraftChanges;

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

                return await change();
            }
        }

        private Task ChangeDraftAsync(Func<Task> change) =>
            this.ChangeDraftAsync(async () =>
            {
                await change();
                return true;
            });

        // One piece of file work, on the pool, under the window's "please wait".
        private Task<T> WriteDraftAsync<T>(string message, Func<T> write) =>
            BusyOverlay.RunLocalAsync(this, message, () => Task.Run(write));

        // ###########################################################################################
        // Closing: once every change has been written. Not on a shutdown of CRT or the computer, which
        // nothing should hold up.
        // ###########################################################################################
        protected override void OnClosing(WindowClosingEventArgs e)
        {
            if (!this.thisDraftChanges.IsCompleted && !this.thisReadyToClose &&
                e.CloseReason is WindowCloseReason.WindowClosing or WindowCloseReason.Undefined)
            {
                e.Cancel = true;

                if (this.thisClosingAfterChanges is null || this.thisClosingAfterChanges.IsCompleted)
                {
                    this.thisClosingAfterChanges = this.CloseAfterChangesAsync();
                }
            }

            base.OnClosing(e);
        }

        private async Task CloseAfterChangesAsync()
        {
            try
            {
                await this.thisDraftChanges;
            }
            catch (Exception ex)
            {
                Logger.Warning($"A change to the draft failed while the window was closing - [{ex.Message}]");
            }

            this.thisReadyToClose = true;
            this.Close();
        }
    }
}
