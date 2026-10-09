using System;

namespace Handlers.Online
{
    // ###########################################################################################
    // "THE SERVER SAID UPDATE CRT" - raised by the two clients of the gated API areas whenever an
    // answer is the 426 ClientOutdatedAnswer (owner request, 2026-10-09; see AppUpdateRequirement):
    // SubmissionClient for the Drafts tab, ReviewApiClient for the Maintainer tab.
    //
    // *** WHY A STATIC EVENT. *** Both clients are made wherever they are needed - a SubmissionClient
    // by the launch check, the minute check, the Submit dialog and "My submissions"; a
    // ReviewApiClient by the Maintainer tab and its sign-in screen - so there is no one owner to hand
    // a callback to, and a refusal met by any of them has to reach Main, which covers the tabs.
    //
    // *** IT HOLDS NO STATE, AND ONLY MAIN'S StartAsync LISTENS. *** A test that drives a client
    // into a 426 raises it with nobody listening (tests build Main but never start it), so nothing
    // leaks from one test into the next. Raised on whatever thread the answer arrived on - a
    // listener posts to the UI thread itself.
    // ###########################################################################################
    public static class ApiOutdatedSignal
    {
        // The area that was refused, and the server's sentence.
        public static event Action<AppUpdateArea, string>? Raised;

        public static void Raise(AppUpdateArea area, string serversWords) =>
            ApiOutdatedSignal.Raised?.Invoke(area, serversWords);
    }
}
