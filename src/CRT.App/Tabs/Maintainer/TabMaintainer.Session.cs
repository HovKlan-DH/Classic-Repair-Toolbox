using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Threading;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE REMEMBERED SIGN-IN, restored AT LAUNCH for the tab's badge - or, failing that, the first
    // time the tab is shown (2026-09-29, 2026-09-30).
    //
    // While this was the CRT Maintainer application, its window restored the session in OnOpened.
    // A tab has no OnOpened; the moment it is first attached to the visual tree - which a
    // TabControl does only when the tab is SELECTED - played that part. Since the tab carries a
    // badge in CRT's row of tabs (owner request, 2026-09-30), that is too late: a maintainer who
    // never opened the tab would never be told anything was waiting. So Main calls
    // RestoreInBackgroundAsync once CRT's window is up, when "Enable Maintainer tab" is on - QUIETLY,
    // with no "please wait": nobody pressed anything. With the setting off it calls
    // RestoreSignInQuietly instead (2026-10-02): the sign-in alone, with no request, so the rest of
    // CRT uses the account's address until the maintainer signs out. Attaching still restores it,
    // under the overlay, when the launch did not.
    //
    // *** ONCE, NOT ON EVERY ATTACH. *** A TabControl detaches a tab's content on every switch away
    // and attaches it again on the way back; restoring each time would re-read the stored file and
    // re-fetch the lists behind a session that is already live. thisSessionRestoreTried says the
    // question has been asked, by either path; signing out does not reset it (the stored session is
    // forgotten then anyway).
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

        // ###########################################################################################
        // *** THE SIGN-IN, FOR THE REST OF CRT (owner request, 2026-10-01: "When I am a maintainer,
        // and I have logged in, then I want to use that email address everywhere in the CRT app" -
        // and 2026-10-02, whatever the Configuration tab says, until signing out).
        // *** The session in use, or null when nobody is signed in; SignedInChanged is raised each
        // time it changes - signing in, the remembered one restored, signing out, or a 401 - and
        // Main hands it to the Feedback and Drafts tabs (Main.ShareMaintainerSignIn). Every change
        // of thisSession goes through UseSession, so none can be missed.
        // ###########################################################################################
        internal ReviewSession? SignedIn => this.thisSession;

        internal event Action? SignedInChanged;

        private void UseSession(ReviewSession? session)
        {
            if (Equals(this.thisSession, session))
                return;

            this.thisSession = session;

            // What was read ahead was read - and judged - for the session before.
            this.thisPrefetch.Forget();

            this.SignedInChanged?.Invoke();
        }

        // Signs in (or out) without the server, for a test of what the rest of CRT does with it.
        internal void UseSessionForTests(ReviewSession? session) => this.UseSession(session);

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

        // ###########################################################################################
        // *** AT LAUNCH, WHATEVER THE CONFIGURATION TAB SAYS: THE REMEMBERED SIGN-IN (owner request,
        // 2026-10-02: "As long as the maintainer is logged in, then use email from that, no matter
        // what is checked in "Configuration" tab. The maintainer will need to logoff to be
        // forgotten"). *** Put in place with NO request - the Feedback tab and the Submit dialog
        // use the account's address from it - also with "Enable Maintainer tab" off, when nothing
        // else of the tab runs. Only signing out (or the session expiring) forgets it. Called by
        // Main once its window is up; never from the constructor (see the header).
        // ###########################################################################################
        internal void RestoreSignInQuietly()
        {
            if (this.thisSessionRestoreTried)
                return;

            this.thisSessionRestoreTried = true;
            this.TryRestoreRememberedSession();
        }

        // ###########################################################################################
        // For the badge, with the tab turned on - at launch, or when it is ticked later: the
        // remembered session, then the two lists the badge counts - with no overlay, since CRT's
        // window must stay usable while the server answers, or does not. The rest (the Systems
        // overview, the drop-down listing) is read when the tab is first shown, which is when
        // anything reads them. Once: Main calls this on every change of the tab's visibility.
        //
        // LAST, the submission the tab will open on (owner request, 2026-10-02 - TabMaintainer
        // .Prefetch.cs), so the first look at the tab has no "please wait".
        // ###########################################################################################
        internal async Task RestoreInBackgroundAsync()
        {
            this.RestoreSignInQuietly();

            if (this.thisSession is null || this.thisQueueKnown || this.thisBadgeListsAsked)
                return;

            this.thisBadgeListsAsked = true;

            await this.RefreshBadgesInBackgroundAsync();

            // The name and address as the server holds them now (TabMaintainer.Account.cs) -
            // before the read-ahead, since a changed session forgets what was read ahead for it.
            await this.ReadAccountQuietlyAsync();

            await this.PrefetchEntryToOpenAsync();
        }

        private bool thisBadgeListsAsked;

        // A remembered, still-usable session goes straight to the queue; none leaves the sign-in
        // panel as it is. Nothing to do when somebody is already signed in.
        private async Task RestoreRememberedSessionAsync()
        {
            if (this.TryRestoreRememberedSession())
            {
                await this.ReadListsAsync();
                await this.ReadAccountQuietlyAsync();
            }
        }

        // True when a remembered session was put in place, and the lists need reading.
        private bool TryRestoreRememberedSession()
        {
            if (this.thisSession is not null)
                return false;

            ReviewSession? remembered = (this.RecallSessionOverrideForTests ?? ReviewSessionStore.Recall)(DateTimeOffset.UtcNow);

            if (remembered is null)
                return false;

            this.thisClient = this.ClientOverrideForTests?.Invoke() ?? new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress);
            this.UseSession(remembered);

            this.ShowQueuePanel();
            return true;
        }

        // A remembered session and a client for a test - never the user's real file, never the real
        // server (ReviewSessionStore is pointed at nothing in a test anyway; see the header).
        internal Func<DateTimeOffset, ReviewSession?>? RecallSessionOverrideForTests { get; set; }

        internal Func<ReviewApiClient>? ClientOverrideForTests { get; set; }
    }
}
