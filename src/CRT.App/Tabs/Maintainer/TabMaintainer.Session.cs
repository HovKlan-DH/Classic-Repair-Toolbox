using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Threading;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE REMEMBERED SIGN-IN, restored the FIRST time the tab is shown (2026-09-29).
    //
    // While this was the the Maintainer tab application, its window restored the session in OnOpened.
    // A tab has no OnOpened; the moment it is first attached to the visual tree - which a
    // TabControl does only when the tab is SELECTED - plays that part. So a maintainer who never
    // opens the tab in a session sends nothing to the server, and one who does is signed in
    // without the password box, exactly as before.
    //
    // *** ONCE, NOT ON EVERY ATTACH. *** A TabControl detaches a tab's content on every switch away
    // and attaches it again on the way back; restoring each time would re-read the stored file and
    // re-fetch the lists behind a session that is already live. thisSessionRestoreTried says the
    // question has been asked; signing out does not reset it (the stored session is forgotten
    // then anyway).
    //
    // *** NEVER IN THE CONSTRUCTOR. *** Restoring means fetching the queue, and Main builds every
    // tab while CRT starts - a slow or unreachable server would then hold up CRT's own window. It
    // also keeps the tests safe: ReviewSessionStore is pointed at the real file only by CRT's
    // start-up (App.StartApplicationAsync), which no test runs, so a tab built in a test finds no
    // stored session and never reaches the live server.
    //
    // *** THE STORED SESSION IS NOT TRUSTED, ONLY OFFERED. *** Recall refuses an expired one, and
    // RefreshQueueAsync then asks the SERVER, the only authority on whether the session is still
    // live: signed out, revoked, or the account locked all come back as a 401 and land in
    // ShowSignInPanel, which clears the file. So the worst case for a stale token is one failed
    // request and a normal sign-in screen.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private bool thisSessionRestoreTried;

        protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnAttachedToVisualTree(e);

            if (!this.thisSessionRestoreTried)
            {
                this.thisSessionRestoreTried = true;

                // Posted, not run inside the attach: the tab is still being laid out, and the
                // restore reads the lists under CRT's window's "please wait" overlay.
                Dispatcher.UIThread.Post(async () => await this.RestoreRememberedSessionAsync());
            }

            this.AttachQueueChecks();
        }

        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            this.DetachQueueChecks();

            base.OnDetachedFromVisualTree(e);
        }

        // A remembered, still-usable session goes straight to the queue; none leaves the sign-in
        // panel as it is. Nothing to do when somebody is already signed in.
        private async Task RestoreRememberedSessionAsync()
        {
            if (this.thisSession is not null)
                return;

            ReviewSession? remembered = ReviewSessionStore.Recall(DateTimeOffset.UtcNow);

            if (remembered is null)
                return;

            this.thisSession = remembered;
            this.thisClient = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress);

            this.ShowQueuePanel();
            await this.ReadListsAsync();
        }
    }
}
