using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // RESETTING THE CONTRIBUTION DATA (owner request, 2026-10-04: "When I go-live with this, it
    // should not have old data visible, so it should be deleted. Of course the real sources of BETA
    // and stable must not be touched, but all contributor and maintainer data should go away").
    // Account > "Reset contribution data".
    //
    // *** WHAT GOES (owner decisions, 2026-10-04) ***
    //   - every submission, whatever its state, with everything the schema hangs off it (files,
    //     payloads, findings, approvals, amendments, discards, BETA returns, change summaries);
    //   - every account that is not an administrator, with its sessions and tokens ("yes, just
    //     delete them as I need real data anyway"), and so every maintainer and invitation;
    //   - every production approval and every SYSTEM RECORD - see below;
    //   - the whole history (audit), the board views ("yes, delete them") and the API usage counts;
    //   - and then, on disk, every stored file no submission needs any more (all of them) and every
    //     partial upload - the blob store's own collection, run at once rather than at the sweeper's
    //     next turn.
    //
    // *** WHAT STAYS ***
    //   - the BETA and stable data trees, their manifests and their main Excel data files - nothing
    //     here opens a file in them;
    //   - the administrators' accounts and sessions - somebody must sign in afterwards, and the one
    //     pressing the button stays signed in;
    //   - the launch check-ins (crt_update - CRT 2.x's history, no migration's), the board view batch
    //     ids, the saved feedback folders.
    //
    // *** THE SYSTEM RECORDS GO TOO, AND THAT IS SAFE. *** A record is made on demand by every flow
    // that needs one (EnsureSystemAsync: a submission, a maintainer or invitation, a placement) and
    // every flow reads a missing one as a system the pipeline never touched. What a record holds
    // about the trees - BETA's and the stable source's content hashes - describes the TEST period's
    // publishes; keeping it would leave a system "waiting for the stable source" with no submission
    // to say why. The owner copies the stable data over BETA and rebuilds the manifests around the
    // reset (owner, 2026-10-04: "I will anyway copy everything from stable to BETA and then update the
    // manifest from admin page"), so the trees are level and no record is the truth.
    //
    // *** NOTHING STANDS IN THE WAY BUT THE SWITCH. *** Unlike deleting a system, nothing waiting in
    // the queue or under BETA > Stable refuses it (owner answer: "yes, just delete them").
    //
    // *** ADMINISTRATOR ONLY, AND ONLY WHILE ServerOptions.AllowDataReset IS ON. *** Off by default;
    // switching it on needs a shell on the box, so a stolen administrator session cannot wipe the
    // database. The plan is answered either way - the screen shows the counts and how to switch it on.
    //
    // *** WHAT WAS SHOWN IS WHAT IS DELETED. *** The reset sends back the plan's fingerprint and is
    // refused when the database moved since (DataResetRules.Fingerprint) - checked against the counts
    // the reset's own transaction reads with those tables locked (IDataResetStore.ResetAsync), since
    // a contributor's create takes no lock this flow holds. It runs under the publish lock, so no
    // publish, promotion or push-back is half way through a submission it deletes.
    //
    // NOBODY IS MAILED: the people losing a submission or an account are the test period's testers,
    // and a mail to each of them is not what a go-live wants. Their CRTs ask about their submissions
    // and are told "not found" - which CRT now shows as "No longer on the server" (CRT.Data's
    // SubmissionReceiptPresenter), no longer blocking a new Submit of the same draft.
    // ###########################################################################################
    public static class DataResetFlow
    {
        public static async Task<DataResetOutcome> PlanAsync(
            ReviewAccess? access,
            ServerOptions options,
            IDataResetStore store,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(store);

            if (!ReviewAuthority.CanAdminister(access))
                return DataResetOutcome.Forbidden();

            DataResetCounts counts = await store.CountAsync(cancellationToken);

            return DataResetOutcome.Planned(DataResetRules.ToPlan(counts, options.AllowDataReset));
        }

        // ###########################################################################################
        // Performs it. `fingerprint` is the plan the administrator confirmed. `removeStoredFiles` is
        // given the deleted submissions' ids and clears their partial uploads and every stored file
        // nothing needs any more, answering how many files went - the endpoint's, so this tests with
        // no blob store. `pendingUsage` is the API usage counted in memory but not yet written.
        //
        // *** THE PENDING USAGE IS DROPPED ONLY ONCE THE RESET IS DONE, AND NO WRITE OF IT CAN LAND
        // AFTER (code review, 2026-10-04). *** The reset holds the counter's writes around its
        // transaction (ApiUsageCounter.HoldWritesAsync): a write already under way finishes first,
        // and its rows are then deleted with the rest, and none starts until the counter has been
        // cleared. It used to be cleared BEFORE the transaction - so a reset the database refused,
        // or one refused for counts that moved, lost up to five minutes of counting although
        // nothing was reset; and a write that had taken its tallies just before could insert them
        // after the reset, or put them back for the next write. What is counted while the
        // transaction runs goes with it - a few seconds, around the reset itself.
        // ###########################################################################################
        public static async Task<DataResetOutcome> ResetAsync(
            ReviewAccess? access,
            string? fingerprint,
            ServerOptions options,
            IDataResetStore store,
            PublishLock publishLock,
            Func<IReadOnlyList<long>, CancellationToken, Task<int>> removeStoredFiles,
            ApiUsageCounter pendingUsage,
            DateTimeOffset nowUtc,
            ILogger logger,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(publishLock);
            ArgumentNullException.ThrowIfNull(removeStoredFiles);
            ArgumentNullException.ThrowIfNull(pendingUsage);
            ArgumentNullException.ThrowIfNull(logger);

            if (!ReviewAuthority.CanAdminister(access))
                return DataResetOutcome.Forbidden();

            if (!options.AllowDataReset)
                return DataResetOutcome.NotEnabled();

            AccountRecord actor = access!.Account;
            string? shown = fingerprint?.Trim();
            DataResetResult? result;

            using (await publishLock.EnterAsync(cancellationToken))
            {
                // The early answer, before anything is touched. The transaction holds the counts to
                // the fingerprint again, under its locks; that is the check that decides.
                DataResetCounts current = await store.CountAsync(cancellationToken);

                if (!string.Equals(DataResetRules.Fingerprint(current), shown, StringComparison.Ordinal))
                    return DataResetOutcome.Conflict(DataResetRules.ChangedMessage);

                using IDisposable usageWrites = await pendingUsage.HoldWritesAsync(cancellationToken);

                try
                {
                    result = await store.ResetAsync(
                        counts => string.Equals(DataResetRules.Fingerprint(counts), shown, StringComparison.Ordinal),
                        deleted => new AuditEntry(
                            actor.Id,
                            actor.Email,
                            DataResetRules.ResetAction,
                            null,
                            DataResetRules.Detail(deleted),
                            nowUtc),
                        cancellationToken);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    logger.LogError(ex, "{Account} pressed reset, but the database refused it; nothing was deleted.", actor.Email);
                    return DataResetOutcome.Failed(DataResetRules.NotDoneMessage);
                }

                // Something arrived between the early count and the transaction's locks.
                if (result is null)
                    return DataResetOutcome.Conflict(DataResetRules.ChangedMessage);

                // Done: the calls counted before it go too - see the method's header.
                pendingUsage.Clear();
            }

            // ###########################################################################################
            // *** AFTER THE FACT, AND IT MAY NOT FAIL THE RESET. *** The rows are gone; a stored file
            // left behind is reclaimed by the abandoned-upload sweeper's next collection, since no
            // submission needs it any more. Not cancellable: the deleted submissions' PARTIAL uploads
            // are found by their ids, which nothing holds once this request is over.
            // ###########################################################################################
            int filesRemoved = 0;

            try
            {
                filesRemoved = await removeStoredFiles(result.SubmissionIds, CancellationToken.None);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "The contribution data was reset, but its stored files could not all be removed now; the sweeper will.");
            }

            logger.LogWarning(
                "{Account} reset the contribution data: {Detail}; {Files} stored file(s) removed.",
                actor.Email, DataResetRules.Detail(result.Deleted), filesRemoved);

            return DataResetOutcome.Done(DataResetRules.ToAnswer(result.Deleted, filesRemoved));
        }
    }

    // ###########################################################################################
    // What DataResetFlow answered. The endpoint maps it: 403 not an administrator, 503 switched off,
    // 409 changed since it was shown, 500 the database refused, 200 a plan or what was deleted.
    // ###########################################################################################
    public sealed record DataResetOutcome(
        DataResetPlanAnswer? Plan,
        DataResetAnswer? Answer,
        string? Error,
        bool IsForbidden = false,
        bool IsNotEnabled = false,
        bool IsConflict = false)
    {
        public static DataResetOutcome Planned(DataResetPlanAnswer plan) => new(plan, null, null);

        public static DataResetOutcome Done(DataResetAnswer answer) => new(null, answer, null);

        public static DataResetOutcome Forbidden() =>
            new(null, null, DataResetRules.NotAdministratorMessage, IsForbidden: true);

        public static DataResetOutcome NotEnabled() =>
            new(null, null, DataResetRules.NotEnabledMessage, IsNotEnabled: true);

        public static DataResetOutcome Conflict(string error) => new(null, null, error, IsConflict: true);

        public static DataResetOutcome Failed(string error) => new(null, null, error);
    }
}
