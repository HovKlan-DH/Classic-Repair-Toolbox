using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW LONG A MAINTAINER'S SIGN-IN LIVES AFTER IT WAS LAST USED - written once, for both sides
    // (code review, 2026-10-13).
    //
    // The server slides a session's expiry forward as it is used (CRT.Server's
    // SessionExtensionRules, with ServerOptions.RefreshTokenDays, whose default is this). CRT
    // remembers the sign-in with the expiry the LOGIN answered, and never learns the slid one - so
    // its remembered copy died 30 days after signing in however often it was used, and a maintainer
    // using the tab every day was asked for the password once a month. CRT now slides its copy the
    // way the server does after every request the server accepted (ReviewSession.AfterUseAt), by
    // this lifetime.
    //
    // *** THE SERVER'S SETTING CAN DIFFER FROM THIS DEFAULT. *** CRT cannot read it. Its copy is a
    // hint, never the authority: a server set shorter answers 401 and CRT forgets the sign-in; one
    // set longer only means CRT asks for the password sooner than it had to.
    // ###########################################################################################
    public static class SessionLifetime
    {
        public const int DefaultDays = 30;

        public static TimeSpan Default => TimeSpan.FromDays(SessionLifetime.DefaultDays);
    }
}
