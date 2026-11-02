using System;
using System.Collections.Generic;
using System.Globalization;
using CRT;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // "Account" > RESET CONTRIBUTION DATA (owner request, 2026-10-04: "When I go-live
    // with this, it should not have old data visible, so it should be deleted. Of course the real
    // sources of BETA and stable must not be touched, but all contributor and maintainer data should
    // go away"). Every word DataResetView shows, and the one rule it applies: the button works only
    // while the server allows a reset AND the confirmation word is typed exactly.
    //
    // The counts are the server's (CRT.Server's DataResetFlow), each line "[14] submissions - ..."
    // with the number bold, as every count in this tab.
    // ###########################################################################################
    public static class DataResetWording
    {
        public const string Heading = "Reset contribution data";

        public const string Explanation =
            "For going live: deletes everything people have sent and done through the contribution service, so " +
            "nothing from the test period is left to see. It cannot be undone. The BETA and stable data are not touched.";

        // What stays - said, so nobody has to wonder whether the boards go too.
        public const string Kept =
            "Not touched: the BETA and stable data, their drop-down lists and checksum manifests, the launch " +
            "check-ins, the saved feedback, and the administrators' accounts.";

        // The word typed to confirm. Exact, capitals included: it is there to be typed on purpose.
        public const string ConfirmWord = "RESET";

        public const string ConfirmPrompt = "Type " + DataResetWording.ConfirmWord + " to confirm:";

        public const string ResetButton = "Reset contribution data";

        // ###########################################################################################
        // WHETHER A RESET THE SERVER NEVER ANSWERED WENT THROUGH (code review, 2026-11-02), judged from
        // `now`, the counts read again - null when they could not be read: the submissions, the accounts
        // and the board records all at zero. Never every count: the reset's own history entry is one,
        // and the API usage and the board views grow again by themselves. Counts that had none of
        // those three to begin with (`before`, those the reset was sent for) cannot tell.
        // ###########################################################################################
        public static bool? HasLanded(DataResetPlanAnswer before, DataResetPlanAnswer? now)
        {
            ArgumentNullException.ThrowIfNull(before);

            return now is null || DataResetWording.HoldsNoneOfWhatIsJudged(before)
                ? null
                : DataResetWording.HoldsNoneOfWhatIsJudged(now);
        }

        private static bool HoldsNoneOfWhatIsJudged(DataResetPlanAnswer plan) =>
            plan is { Submissions: 0, Accounts: 0, Boards: 0 };

        public static string AfterTimeout(bool? landed) =>
            WaitWording.AfterTimeout(landed, "the contribution data is reset", "the contribution data is still there");

        public const string Changed =
            "Something arrived or went since the counts were shown - they were read again. Look at them and confirm again.";

        // ###########################################################################################
        // Whether the button may be pressed: counts to send back, a server that allows the reset, and
        // the word typed - trimmed, since a stray space is not a change of mind.
        // ###########################################################################################
        public static bool CanReset(DataResetPlanAnswer? plan, string? typed) =>
            plan is { IsEnabled: true } &&
            !string.IsNullOrWhiteSpace(plan.Fingerprint) &&
            string.Equals(typed?.Trim(), DataResetWording.ConfirmWord, StringComparison.Ordinal);

        // ###########################################################################################
        // One line per kind of thing deleted, in the order a reader thinks of them: people's work,
        // the people, then what was kept about them.
        // ###########################################################################################
        public static IReadOnlyList<IReadOnlyList<ReviewNoteRun>> Lines(DataResetPlanAnswer plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            return
            [
                DataResetWording.Line(plan.Submissions, "submission", "submissions", " - every one, whatever its state, with its uploaded files"),
                DataResetWording.Line(plan.OscilloscopeModels, "oscilloscope model sent in from CRT", "oscilloscope models sent in from CRT", " - waiting or decided"),
                DataResetWording.Line(plan.Accounts, "account", "accounts",
                    $" - every one but the administrators ({DataResetWording.Count(plan.Administrators, "administrator", "administrators")} kept)"),
                DataResetWording.Line(plan.Maintainers, "maintainer of a board", "maintainers of a board", string.Empty),
                DataResetWording.Line(plan.Invitations, "invitation to maintain", "invitations to maintain", string.Empty),
                DataResetWording.Line(plan.Boards, "board record", "board records",
                    " - made again as each board is next used; the boards themselves stay"),
                DataResetWording.Line(plan.HistoryEntries, "history entry", "history entries", " - the reset itself becomes the first"),
                DataResetWording.Line(plan.BoardViews, "board view", "board views", string.Empty),
                DataResetWording.Line(plan.ApiUsageRows, "API usage row", "API usage rows", string.Empty)
            ];
        }

        // What the reset deleted, in one line.
        public static IReadOnlyList<ReviewNoteRun> Done(DataResetAnswer answer)
        {
            ArgumentNullException.ThrowIfNull(answer);

            var runs = new List<ReviewNoteRun> { new("The contribution data was reset: ", IsCount: false) };

            (int Count, string One, string Many)[] parts =
            [
                (answer.SubmissionsDeleted, "submission", "submissions"),
                (answer.OscilloscopeModelsDeleted, "oscilloscope model", "oscilloscope models"),
                (answer.AccountsDeleted, "account", "accounts"),
                (answer.MaintainersDeleted, "maintainer", "maintainers"),
                (answer.InvitationsDeleted, "invitation", "invitations"),
                (answer.BoardsDeleted, "board record", "board records"),
                (answer.HistoryEntriesDeleted, "history entry", "history entries"),
                (answer.BoardViewsDeleted, "board view", "board views"),
                (answer.ApiUsageRowsDeleted, "API usage row", "API usage rows")
            ];

            for (int index = 0; index < parts.Length; index++)
            {
                if (index > 0)
                    runs.Add(new ReviewNoteRun(index == parts.Length - 1 ? " and " : ", ", IsCount: false));

                runs.AddRange(DataResetWording.Line(parts[index].Count, parts[index].One, parts[index].Many, string.Empty));
            }

            runs.Add(new ReviewNoteRun(" deleted, and ", IsCount: false));
            runs.AddRange(DataResetWording.Line(answer.StoredFilesRemoved, "stored file", "stored files", " removed from the disk."));

            return runs;
        }

        // "[14] submissions" + rest, the number bold.
        private static IReadOnlyList<ReviewNoteRun> Line(int count, string one, string many, string rest) =>
        [
            new ReviewNoteRun("[", IsCount: false),
            new ReviewNoteRun(count.ToString("N0", CultureInfo.InvariantCulture), IsCount: true),
            new ReviewNoteRun("] " + (count == 1 ? one : many) + rest, IsCount: false)
        ];

        private static string Count(int count, string one, string many) =>
            string.Create(CultureInfo.InvariantCulture, $"{count} {(count == 1 ? one : many)}");
    }
}
