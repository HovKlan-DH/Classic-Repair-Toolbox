using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The pure half of resetting the contribution data (owner request, 2026-10-04): the fingerprint
    // the administrator confirms, the answers sent back, and the words. DataResetFlow sequences.
    // ###########################################################################################
    public static class DataResetRules
    {
        // The audit action of the one row a reset leaves - the first row of the new history.
        public const string ResetAction = "data.reset";

        public const string NotEnabledMessage =
            "Resetting is switched off on this server. To use it, set CrtServer:AllowDataReset to true in " +
            "the service's appsettings.Production.json and restart the service - and set it back to false " +
            "once the reset is done.";

        public const string NotAdministratorMessage = "Only an administrator can reset the contribution data.";

        public const string ChangedMessage =
            "Something changed since the counts were shown - a submission or an account arrived or went. " +
            "Look at the counts again before resetting.";

        public const string NotDoneMessage =
            "The reset could not be done, and nothing was deleted - the database refused it. The service's log says why.";

        // ###########################################################################################
        // *** WHAT THE FINGERPRINT COVERS. *** The submissions and the accounts by count AND highest
        // id - a count alone stays the same when one goes and another arrives - and the maintainers,
        // invitations, production approvals and board records by count.
        //
        // NOT the history, the board views or the API usage counts: they grow by themselves (a
        // sign-in writes history, CRTs report board views every 15 minutes, API usage is written every
        // few minutes), so with them in it every reset would be refused as "changed". They are deleted
        // all the same - what they hold is never somebody's work in the queue.
        //
        // NOR the administrators: an administrator granted by hand meanwhile is kept either way.
        // ###########################################################################################
        public static string Fingerprint(DataResetCounts counts)
        {
            ArgumentNullException.ThrowIfNull(counts);

            string text = string.Join(
                "|",
                counts.Submissions,
                counts.LastSubmissionId,
                counts.Accounts,
                counts.LastAccountId,
                counts.Maintainers,
                counts.Invitations,
                counts.ProductionApprovals,
                counts.Boards);

            return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
        }

        public static DataResetPlanAnswer ToPlan(DataResetCounts counts, bool isEnabled)
        {
            ArgumentNullException.ThrowIfNull(counts);

            return new DataResetPlanAnswer(
                isEnabled,
                DataResetRules.Fingerprint(counts),
                counts.Submissions,
                counts.Accounts,
                counts.Administrators,
                counts.Maintainers,
                counts.Invitations,
                counts.Boards,
                counts.HistoryEntries,
                counts.BoardViews,
                counts.ApiUsageRows,
                isEnabled ? null : DataResetRules.NotEnabledMessage);
        }

        public static DataResetAnswer ToAnswer(DataResetCounts deleted, int storedFilesRemoved)
        {
            ArgumentNullException.ThrowIfNull(deleted);

            return new DataResetAnswer(
                deleted.Submissions,
                deleted.Accounts,
                deleted.Maintainers,
                deleted.Invitations,
                deleted.Boards,
                deleted.HistoryEntries,
                deleted.BoardViews,
                deleted.ApiUsageRows,
                storedFilesRemoved);
        }

        // The detail of the reset's own history row - what went, in one line.
        public static string Detail(DataResetCounts deleted)
        {
            ArgumentNullException.ThrowIfNull(deleted);

            return string.Create(
                CultureInfo.InvariantCulture,
                $"{deleted.Submissions} submission(s), {deleted.Accounts} account(s), {deleted.Maintainers} maintainer(s), " +
                $"{deleted.Invitations} invitation(s), {deleted.ProductionApprovals} production approval(s), " +
                $"{deleted.Boards} board record(s), {deleted.HistoryEntries} history row(s), " +
                $"{deleted.BoardViews} board view(s) and {deleted.ApiUsageRows} API usage row(s) deleted; " +
                $"{deleted.Administrators} administrator(s) kept");
        }
    }
}
