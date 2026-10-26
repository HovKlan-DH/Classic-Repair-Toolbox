using System;
using System.IO;
using CRT;

namespace Handlers.Online
{
    // ###########################################################################################
    // THE CONTRIBUTION SERVICE'S TEST INSTANCE, AS CRT SEES IT (owner decision, 2026-10-25: "We
    // need the test server"). CRT 4.0's pre-releases submit to it and sign in to it
    // (AppConfig.ContributionServerRootUrl), so the live service keeps serving CRT 3.x until the
    // release; the check-in, feedback and board views go to the live service all the same.
    //
    //   - WHAT CRT REMEMBERS OF A SERVICE IS KEPT PER SERVICE (FileNameFor). The Maintainer tab's
    //     sign-in (ReviewSessionStore) and the receipts of what was sent (SubmissionReceiptStore)
    //     mean nothing to another service: a token, and submission numbers, are one database's.
    //     Shared with the live files - which CRT 3.x reads and writes too - a pre-release would send
    //     the live token to the test instance and forget it on the 401, signing 3.x out (and 3.x the
    //     same the other way), and would ask the test instance about live submissions.
    //   - A RELEASE NEVER SHIPS POINTING AT IT (MayShip, held by ContributionServerTests, which the
    //     release workflow runs before it builds anything).
    //   - THE SCREENS THAT TALK TO IT SAY SO: the Maintainer tab's sign-in (its accounts are its
    //     own) and the Submit dialog (what it publishes is test data).
    // ###########################################################################################
    public static class ContributionServer
    {
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
        // Whether a CRT built as `informationalVersion` may ship talking to `contributionRootUrl`: a
        // RELEASE - no "-" pre-release tag before any "+" build metadata - never to the test instance.
        // A pre-release may talk to either.
        // ###########################################################################################
        public static bool MayShip(string? informationalVersion, string? contributionRootUrl)
        {
            string version = (informationalVersion ?? string.Empty).Split('+')[0].Trim();
            bool isRelease = version.Length > 0 && !version.Contains('-');
            bool isTestServer = string.Equals(
                (contributionRootUrl ?? string.Empty).Trim().TrimEnd('/'),
                AppConfig.TestServerRootUrl,
                StringComparison.OrdinalIgnoreCase);

            return !(isRelease && isTestServer);
        }

        // ###########################################################################################
        // Whether a draft whose work reached the stable data may be deleted by itself
        // (Main.RetirePublishedDraftsAsync): never while CRT talks to the test instance (code review,
        // 2026-10-25). What the test instance publishes is replaced by the release day's conversion,
        // so there the draft is the only lasting copy of the contributor's work - kept, to be sent to
        // the live service once CRT 4.0 is released.
        // ###########################################################################################
        public static bool RetiresPublishedDrafts(bool testServer) => !testServer;

        // Above the Maintainer tab's sign-in, while CRT talks to the test instance.
        public const string SignInNote =
            "This pre-release of CRT signs in to the test server for CRT 4.0. It has accounts of its own: " +
            "your account on the live server does not work here, and nothing done here reaches the live data.";

        // In the Submit dialog, while CRT talks to the test instance.
        public const string SubmitNote =
            "This pre-release of CRT sends your changes to the test server for CRT 4.0. What is published " +
            "there is test data: everybody else keeps the live data, and the test data is replaced when " +
            "CRT 4.0 is released. Your draft is kept, so your changes can be sent again then.";
    }
}
