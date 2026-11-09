using System;
using System.IO;
using CRT;

namespace Handlers.Online
{
    // ###########################################################################################
    // THE CONTRIBUTION SERVICE'S TEST INSTANCE, AS CRT SEES IT (owner decisions, 2026-10-25: "We
    // need the test server"; and 2026-11-09: every pre-release talks to it, every release to the live
    // service). A CRT whose version is an alpha or a beta submits to it and signs in to it
    // (RootUrlFor, AppConfig.ContributionServerRootUrl); the check-in, feedback and board views go to
    // the live service all the same.
    //
    //   - WHAT CRT REMEMBERS OF A SERVICE IS KEPT PER SERVICE (FileNameFor). The Maintainer tab's
    //     sign-in (ReviewSessionStore) and the receipts of what was sent (SubmissionReceiptStore)
    //     mean nothing to another service: a token, and submission numbers, are one database's.
    //     Shared with the live files - which a release reads and writes too - a pre-release would send
    //     the live token to the test instance and forget it on the 401, signing the release out (and
    //     the release the same the other way), and would ask the test instance about live submissions.
    //   - A RELEASE NEVER TALKS TO IT: the version decides, so there is nothing to set back by hand.
    //   - THE SCREENS THAT TALK TO IT SAY SO: the Maintainer tab's sign-in (its accounts are its
    //     own) and the Submit dialog (what it publishes is test data).
    // ###########################################################################################
    public static class ContributionServer
    {
        // ###########################################################################################
        // Whether a CRT built as `informationalVersion` is a release: a version with no "-" pre-release
        // tag before any "+" build metadata. "4.0.0" is; "4.0.0-alpha.4", "4.1.0-beta.2+7f3a2c1" and no
        // version at all are not.
        // ###########################################################################################
        public static bool IsRelease(string? informationalVersion)
        {
            string version = (informationalVersion ?? string.Empty).Split('+')[0].Trim();
            return version.Length > 0 && !version.Contains('-');
        }

        // ###########################################################################################
        // The contribution service a CRT built as `informationalVersion` talks to (owner decision,
        // 2026-11-09): the live one (port 5199 behind Apache) for a release, the test instance (5198)
        // for an alpha or a beta.
        // ###########################################################################################
        public static string RootUrlFor(string? informationalVersion) =>
            ContributionServer.IsRelease(informationalVersion) ? AppConfig.CrtServerRootUrl : AppConfig.TestServerRootUrl;

        // What a file of the test instance's own carries before its extension.
        public const string TestFileSuffix = "-test";

        // ###########################################################################################
        // The file to keep `fileName` in for the service CRT talks to: as it is for the live
        // service, "submissions-test.json" for "submissions.json" while it is the test instance.
        // ###########################################################################################
        public static string FileNameFor(string fileName, bool testServer)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

            return testServer
                ? Path.GetFileNameWithoutExtension(fileName) + ContributionServer.TestFileSuffix + Path.GetExtension(fileName)
                : fileName;
        }

        // The same for the service this CRT is built to talk to.
        public static string FileNameFor(string fileName) =>
            ContributionServer.FileNameFor(fileName, AppConfig.UsesTestServer);

        // ###########################################################################################
        // Whether a draft whose work reached the stable data may be deleted by itself
        // (Main.RetirePublishedDraftsAsync): never while CRT talks to the test instance (code review,
        // 2026-10-25). What the test instance publishes never reaches the live data - it is replaced
        // whenever its trees are copied from the live ones again - so there the draft is the only
        // lasting copy of the contributor's work, kept to be sent from a released CRT.
        // ###########################################################################################
        public static bool RetiresPublishedDrafts(bool testServer) => !testServer;

        // Above the Maintainer tab's sign-in, while CRT talks to the test instance.
        public const string SignInNote =
            "This pre-release of CRT signs in to the test server. It has accounts of its own: " +
            "your account on the live server does not work here, and nothing done here reaches the live data.";

        // In the Submit dialog, while CRT talks to the test instance.
        public const string SubmitNote =
            "This pre-release of CRT sends your changes to the test server. What is published there is " +
            "test data: everybody else keeps the live data. Your draft is kept, so your changes can be " +
            "sent again from a released CRT.";
    }
}
