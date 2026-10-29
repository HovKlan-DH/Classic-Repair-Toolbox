using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // AN OSCILLOSCOPE MODEL SENT IN FROM CRT, AND THE ADMINISTRATOR'S DECISION ON IT (CRT 4.0 step 5
    // - the plan: "Submit this model sends it to the administrator's queue on the Maintainer tab,
    // which publishes it into the main file for everybody. This is a new kind of submission on the
    // server"; owner decision 2026-10-24: "as a submission to your queue, published with one click").
    //
    // *** SENDING ONE NEEDS NO ACCOUNT, as a board submission. *** The address typed is what the
    // contributor is told the outcome at; a signed-in sender is told at the account's. A model is
    // checked by OscilloscopeModelRules.SubmissionProblems, the rule CRT asks before it sends. Each
    // internet connection (ConnectionOf: an IPv4 address, an IPv6 /64) may send MaxPerAddress models a
    // day, counted and kept as one step (the request limiter on the route counts calls of any kind;
    // this counts what is kept), and at most MaxWaiting wait at once from everybody together. A newer
    // model of the same name from the same contributor - the same
    // account, or without one the same typed address AND connection - takes the place of the one still
    // pending (the store marks it replaced) - the board rule; one already in BETA keeps its place.
    //
    // *** THE ADMINISTRATOR DECIDES, ALONE. *** The routes are AdminEndpoints' (every one refused to
    // anybody else). A model is not a board: no maintainer's pool, no two stages.
    //
    // *** "PUBLISH" WRITES THE BETA AND THE STABLE MAIN FILES AT ONCE, as "Order of boards" does
    // (BoardOrderFlow) - a model is a fact of the list, not of any board waiting in BETA, and "one
    // click" was the decision. In place of the published model of its name (an update), else after
    // the last; each tree changed gets its checksum manifest rebuilt, so CRT downloads it at its next
    // launch. Under the one PublishLock. BETA first, never rolled back. A stable source that could
    // not take the model leaves it IN BETA (OscilloscopeModelStates.InBeta), still listed, the BETA
    // write audited, its contributor not told yet - Publish again finishes it, since a model already
    // in BETA writes nothing there (code reviews, 2026-10-29); a manifest that could not be rebuilt is
    // said (OscilloscopeModelDecisionAnswer.Problem) - "Rebuild checksum manifests" puts it right -
    // and neither the contributor's mail nor the answer then promises the next launch. A publish that
    // writes BETA settles every other copy of its name still in BETA (replaced): finished, its older
    // commands would be written back over the newer ones; a copy another contributor sent is mailed.
    //
    // "REJECT" needs a reason - the contributor is mailed it, and a rejection with none reads as
    // being ignored (SubmissionRejected's rule). It takes the publish lock too, so no decision lands
    // in the middle of a publish, and it refuses a model in BETA: that cannot be taken back out, and
    // "it will not be going in" would be untrue. Mails never fail what they follow.
    // ###########################################################################################
    public static class OscilloscopeModelFlows
    {
        public const int MaxPerAddress = 10;

        // ###########################################################################################
        // At most this many models wait for the administrator at once, from everybody together (code
        // review, 2026-10-29 - Security). The daily limit is per internet connection, and a sender
        // with a whole IPv6 prefix (a /48 from a tunnel broker is free) has thousands of /64s - each
        // model kept is mailed to every administrator. This bounds the rows and the mails until the
        // administrator decides; far more than a hobby project's models ever wait.
        // ###########################################################################################
        public const int MaxWaiting = 100;

        public const string TooManyWaiting =
            "A lot of oscilloscope models are waiting for the administrator already - please send yours again in a few days.";

        public static readonly TimeSpan Window = TimeSpan.FromHours(24);

        public const string PublishedAction = "oscilloscope_model.published";

        public const string RejectedAction = "oscilloscope_model.rejected";

        // BETA's main file took the model, the stable one's could not.
        public const string InBetaAction = "oscilloscope_model.in_beta";

        public const string SentMessage =
            "Sent. The administrator looks at it and publishes it for everybody - you get a mail either way.";

        public const string NoContact =
            "An email address is needed so you can be told whether your model was published. It is used for that and nothing else.";

        public const string BadContact = "That does not look like an email address.";

        // Counted by the internet connection, never the email address typed - and said so (code
        // review, 2026-10-29: "This address" read as the email address, and a fresh one was refused).
        public static string TooMany =>
            $"{OscilloscopeModelFlows.MaxPerAddress} oscilloscope models have been sent from your internet connection today - " +
            "please send the next one tomorrow.";

        public const string NoReason = "Say why it will not go in - the contributor is mailed it.";

        public const string NoSuchModel = "No such oscilloscope model is waiting - it may have been decided already.";

        public const string AlreadyDecided = "This model was decided meanwhile - look at the list again.";

        public const string PublishAgain = "Put that right, then press Publish again - it finishes the publish.";

        public static string InBetaNotRejected(string name) =>
            $"{name} is in BETA already - the stable source could not take it when it was published, and it cannot be taken " +
            "back out. Put the stable source right, then press Publish again to finish it.";

        private const string WhatChanged = "the model";

        // One count-and-keep at a time (SubmitAsync).
        private static readonly SemaphoreSlim SendGate = new(1, 1);

        // ###########################################################################################
        // THE INTERNET CONNECTION A MODEL CAME FROM, as the daily limit counts it and the record keeps
        // it: an IPv4 address as it is (one inside IPv6 too), an IPv6 address as its /64 network - a
        // home or a phone gets a whole /64, so counted by address a sender could take a new one for
        // every ten (code review, 2026-10-29 - Security). Null for none; text that is no address is
        // kept cut to the column.
        // ###########################################################################################
        public static string? ConnectionOf(string? address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return null;

            if (!IPAddress.TryParse(address.Trim(), out IPAddress? parsed))
                return BoardViewRules.Clip(address, 64);

            if (parsed.IsIPv4MappedToIPv6)
                return parsed.MapToIPv4().ToString();

            if (parsed.AddressFamily != AddressFamily.InterNetworkV6)
                return parsed.ToString();

            byte[] bytes = parsed.GetAddressBytes();
            Array.Clear(bytes, 8, 8);

            return $"{new IPAddress(bytes)}/64";
        }

        // ###########################################################################################
        // POST /api/submissions/oscilloscope-models.
        // ###########################################################################################
        public static async Task<OscilloscopeModelSubmitOutcome> SubmitAsync(
            OscilloscopeModelSubmissionRequest? request,
            AccountRecord? account,
            string? senderAddress,
            string? clientVersion,
            IOscilloscopeModelStore store,
            IAccountStore accounts,
            IEmailSender mailer,
            DateTimeOffset nowUtc,
            ILogger? logger = null,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(mailer);

            if (request is null)
                return OscilloscopeModelSubmitOutcome.Refused(["The request had no model in it."]);

            var problems = OscilloscopeModelRules.SubmissionProblems(request.Vendor, request.SeriesOrModel, request.Values, request.Message).ToList();
            string? contactEmail = request.ContactEmail?.Trim();

            if (account is null)
            {
                if (string.IsNullOrWhiteSpace(contactEmail))
                    problems.Add(OscilloscopeModelFlows.NoContact);
                else if (!EmailAddressRules.IsPlausible(contactEmail))
                    problems.Add(OscilloscopeModelFlows.BadContact);
            }

            if (problems.Count > 0)
                return OscilloscopeModelSubmitOutcome.Refused(problems);

            MainDocument.OscilloscopeRow model = OscilloscopeModelValues.ToModel(request.Vendor, request.SeriesOrModel, request.Values);
            string? message = string.IsNullOrWhiteSpace(request.Message) ? null : request.Message.Trim();
            string? connection = OscilloscopeModelFlows.ConnectionOf(senderAddress);
            long id;

            // *** COUNTED AND KEPT AS ONE STEP (code review, 2026-10-29 - Security). *** Counted, then
            // kept with nothing between, a burst from one connection all counted fewer than the limit -
            // each one kept and mailed to every administrator. The service is one process, so one
            // gate holds every count to the models kept before it.
            await OscilloscopeModelFlows.SendGate.WaitAsync(cancellationToken).ConfigureAwait(false);

            try
            {
                if (connection is not null &&
                    await store.CountFromAddressSinceAsync(connection, nowUtc - OscilloscopeModelFlows.Window, cancellationToken) >= OscilloscopeModelFlows.MaxPerAddress)
                {
                    return OscilloscopeModelSubmitOutcome.RateLimited(OscilloscopeModelFlows.TooMany);
                }

                if (await store.CountWaitingAsync(cancellationToken) >= OscilloscopeModelFlows.MaxWaiting)
                {
                    return OscilloscopeModelSubmitOutcome.RateLimited(OscilloscopeModelFlows.TooManyWaiting);
                }

                id = await store.AddAsync(
                    new NewOscilloscopeModelSubmission(
                        nowUtc,
                        model.Brand,
                        model.SeriesOrModel,
                        JsonSerializer.Serialize(OscilloscopeModelValues.Of(model)),
                        account is null ? contactEmail : null,
                        BoardViewRules.Clip(request.ContactName, 200),
                        account?.Id,
                        message,
                        request.FullTestSuitePassed,
                        BoardViewRules.Clip(clientVersion, 64),
                        connection),
                    cancellationToken);
            }
            finally
            {
                OscilloscopeModelFlows.SendGate.Release();
            }

            // The model is kept: from here nothing may turn that into an error the contributor sees.
            try
            {
                foreach (AccountRecord administrator in await accounts.GetAdministratorsAsync(CancellationToken.None))
                {
                    await OscilloscopeModelFlows.SendAsync(
                        mailer,
                        () => EmailTemplates.OscilloscopeModelWaiting(
                            administrator.Email,
                            OscilloscopeModelFlows.NameOf(model.Brand, model.SeriesOrModel),
                            id,
                            message,
                            request.FullTestSuitePassed,
                            administrator.DisplayName),
                        logger);
                }
            }
            catch (Exception ex)
            {
                logger?.LogWarning(ex, "The administrators could not be told of oscilloscope model {Id}.", id);
            }

            return OscilloscopeModelSubmitOutcome.Accepted(new OscilloscopeModelSubmissionAnswer(id, OscilloscopeModelFlows.SentMessage));
        }

        // ###########################################################################################
        // GET /api/admin/oscilloscope-models - every model waiting, oldest first, with whoever sent
        // it (the account's name and address when sent signed in) and the published model of its
        // name in BETA, when there is one, so the administrator sees what an update changes.
        // ###########################################################################################
        public static async Task<OscilloscopeModelQueueAnswer> ListAsync(
            IOscilloscopeModelStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(options);

            IReadOnlyList<OscilloscopeModelSubmissionRecord> waiting = await store.ListPendingAsync(cancellationToken);
            IReadOnlyList<MainDocument.OscilloscopeRow> published = OscilloscopeModelFlows.PublishedModels(OscilloscopeModelFlows.BetaRoot(options));

            IReadOnlyDictionary<long, AccountRecord> senders = await accounts.FindByIdsAsync(
                waiting.Where(record => record.AccountId is not null).Select(record => record.AccountId!.Value).Distinct().ToList(),
                cancellationToken);

            var entries = new List<OscilloscopeModelQueueEntry>(waiting.Count);

            foreach (OscilloscopeModelSubmissionRecord record in waiting)
            {
                AccountRecord? sender = record.AccountId is long accountId && senders.TryGetValue(accountId, out AccountRecord? found) ? found : null;
                MainDocument.OscilloscopeRow? same = published.FirstOrDefault(model => OscilloscopeModelRules.IsNamed(model, record.Vendor, record.SeriesOrModel));

                entries.Add(new OscilloscopeModelQueueEntry(
                    record.Id,
                    record.CreatedUtc,
                    record.Vendor,
                    record.SeriesOrModel,
                    OscilloscopeModelFlows.ValuesOf(record),
                    sender?.DisplayName ?? record.ContactName ?? string.Empty,
                    sender?.Email ?? record.ContactEmail ?? string.Empty,
                    record.Message ?? string.Empty,
                    record.FullTestSuitePassed,
                    record.ClientVersion ?? string.Empty,
                    same is null ? null : OscilloscopeModelValues.Of(same),
                    record.State == OscilloscopeModelStates.InBeta));
            }

            return new OscilloscopeModelQueueAnswer(entries);
        }

        // ###########################################################################################
        // POST /api/admin/oscilloscope-models/publish - see the header.
        // ###########################################################################################
        public static async Task<OscilloscopeModelDecisionOutcome> PublishAsync(
            ReviewAccess access,
            OscilloscopeModelDecisionRequest? request,
            IOscilloscopeModelStore store,
            IAccountStore accounts,
            IEmailSender mailer,
            ServerOptions options,
            Func<UnusedFileFlows.DataTree, int> writeManifest,
            PublishLock publishLock,
            DateTimeOffset nowUtc,
            ILogger? logger = null,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(access);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(writeManifest);
            ArgumentNullException.ThrowIfNull(publishLock);

            if (request is null)
                return OscilloscopeModelDecisionOutcome.NotFound(OscilloscopeModelFlows.NoSuchModel);

            IReadOnlyList<UnusedFileFlows.DataTree> trees = ManifestRebuildFlow.TreesToRebuild(options);
            UnusedFileFlows.DataTree? beta = trees.FirstOrDefault(tree => tree.Name == UnusedFileFlows.Beta);
            UnusedFileFlows.DataTree? stable = trees.FirstOrDefault(tree => tree.Name == UnusedFileFlows.Production);
            string? betaRoot = OscilloscopeModelFlows.BetaRoot(options);

            if (string.IsNullOrWhiteSpace(betaRoot))
                return OscilloscopeModelDecisionOutcome.Conflict("The server has no data tree configured.");

            using IDisposable gate = await publishLock.EnterAsync(cancellationToken).ConfigureAwait(false);

            OscilloscopeModelSubmissionRecord? record = await store.FindAsync(request.Id, cancellationToken);

            if (record is null)
                return OscilloscopeModelDecisionOutcome.NotFound(OscilloscopeModelFlows.NoSuchModel);

            if (!OscilloscopeModelStates.IsWaiting(record.State))
                return OscilloscopeModelDecisionOutcome.Conflict(OscilloscopeModelFlows.AlreadyDecided);

            if (MasterListing.NewestMasterPath(betaRoot) is not string betaMain)
                return OscilloscopeModelDecisionOutcome.Conflict(BoardListingRules.NoMasterMessage);

            MasterListingEdit betaEdit = MasterListing.PutOscilloscope(betaMain, OscilloscopeModelFlows.ModelOf(record), nowUtc);

            if (!betaEdit.IsDone)
                return OscilloscopeModelDecisionOutcome.Conflict(betaEdit.Failure ?? "BETA's main data file could not be written.");

            // *** FROM HERE ON, NOTHING IS CANCELLED *** - BETA's main file is written (BoardOrderFlow's
            // reason: a manifest left naming the old checksum makes every CRT throw the download away).
            var problems = new List<string>();

            if (betaEdit.Changed && beta is not null)
                await BoardOrderFlow.RebuildAsync(beta, writeManifest, problems, OscilloscopeModelFlows.WhatChanged).ConfigureAwait(false);

            bool? stableChanged = null;
            string? stableNotTaken = null;

            if (stable is not null)
            {
                stableChanged = false;

                if (MasterListing.NewestMasterPath(stable.Root) is not string stableMain)
                {
                    stableNotTaken = "The stable source has no main data file, so the model is not in it.";
                }
                else
                {
                    MasterListingEdit stableEdit = MasterListing.PutOscilloscope(stableMain, OscilloscopeModelFlows.ModelOf(record), nowUtc);

                    if (!stableEdit.IsDone)
                    {
                        stableNotTaken = $"The model could not be put into the stable source: {stableEdit.Failure}";
                    }
                    else if (stableEdit.Changed)
                    {
                        stableChanged = true;
                        await BoardOrderFlow.RebuildAsync(stable, writeManifest, problems, OscilloscopeModelFlows.WhatChanged).ConfigureAwait(false);
                    }
                }
            }

            string name = OscilloscopeModelFlows.NameOf(record.Vendor, record.SeriesOrModel);

            // *** ANOTHER COPY OF ITS NAME STILL IN BETA IS NO LONGER WHAT BETA HOLDS (code reviews,
            // 2026-10-29). *** Left listed with "press Publish again", that Publish wrote its older
            // commands back over this one in both trees. It is replaced - out of the list, its Publish
            // refused - once both trees are written, and a database that fails here fails nothing
            // written before it. A copy another contributor sent is told it went; one the same
            // contributor sent was replaced by their own newer copy.
            int settled = await OscilloscopeModelFlows.SettleOtherCopiesAsync(record, store, accounts, mailer, name, logger);

            // *** NOT IN STABLE: IT STAYS LISTED, IN BETA (code reviews, 2026-10-29). *** Marked
            // published, a second Publish was refused as "decided meanwhile" and stable never got the
            // model; left pending, a Reject or a newer copy took it out of the list while BETA kept it.
            // BETA keeps it (never rolled back), the BETA write is audited, its contributor is not told
            // yet; Publish again - once the stable source is put right - writes stable, and BETA,
            // holding the model, writes nothing.
            if (stableNotTaken is not null)
            {
                logger?.LogWarning("Publishing oscilloscope model {Id}: {Problem}", record.Id, stableNotTaken);

                if (record.State != OscilloscopeModelStates.InBeta && !await store.MarkInBetaAsync(record.Id, CancellationToken.None))
                    logger?.LogWarning("Oscilloscope model {Id} is in BETA, but its record could not be marked so.", record.Id);

                if (betaEdit.Changed)
                    await OscilloscopeModelFlows.AuditAsync(accounts, access, OscilloscopeModelFlows.InBetaAction, record, $"stable=not taken | {stableNotTaken}", nowUtc, logger);

                return OscilloscopeModelDecisionOutcome.Done(new OscilloscopeModelDecisionAnswer(
                    betaEdit.Changed,
                    stableChanged,
                    $"{name} is in BETA, not yet in the stable source - it stays in the list.{OscilloscopeModelFlows.SettledLine(settled)}",
                    string.Join(" ", problems.Prepend(stableNotTaken).Append(OscilloscopeModelFlows.PublishAgain))));
            }

            // What was written IS published - even when a newer copy from the same contributor marked
            // this one replaced while it was written (that copy waits on its own). Reject cannot land
            // meanwhile: it takes the same lock.
            if (!await store.DecideAsync(record.Id, OscilloscopeModelStates.Published, access.Account.Id, null, nowUtc, evenIfReplaced: true, CancellationToken.None))
            {
                logger?.LogWarning("Oscilloscope model {Id} was published, but its record could not be marked so.", record.Id);
            }

            await OscilloscopeModelFlows.AuditAsync(accounts, access, OscilloscopeModelFlows.PublishedAction, record, $"beta={(betaEdit.Changed ? "changed" : "unchanged")} | stable={(stableChanged switch { true => "changed", false => "unchanged", null => "none" })}", nowUtc, logger);
            // A manifest not rebuilt keeps CRT from taking the new main file yet: no promise of the next
            // launch then (code review, 2026-10-29).
            bool everybodyGetsItNow = problems.Count == 0;

            await OscilloscopeModelFlows.TellContributorAsync(
                record,
                accounts,
                mailer,
                (address, contributorName) => EmailTemplates.OscilloscopeModelPublished(address, name, contributorName, everybodyGetsItNow),
                logger);

            foreach (string line in problems)
                logger?.LogWarning("Publishing oscilloscope model {Id}: {Problem}", record.Id, line);

            // The administrator's sentence promises what the mail promises (code review, 2026-10-29).
            return OscilloscopeModelDecisionOutcome.Done(new OscilloscopeModelDecisionAnswer(
                betaEdit.Changed,
                stableChanged,
                (everybodyGetsItNow
                    ? $"{name} is published - CRT downloads it at its next launch."
                    : $"{name} is published - CRT downloads it once the checksum manifests are rebuilt.")
                + OscilloscopeModelFlows.SettledLine(settled),
                problems.Count == 0 ? null : string.Join(" ", problems)));
        }

        // ###########################################################################################
        // The other copies of `published`'s name still in BETA marked replaced (the store's
        // ReplaceInBetaOfNameAsync) - how many. Their contributors are mailed when another
        // contributor's model took their place: told "you get a mail either way", they were never
        // told otherwise (code review, 2026-10-29). Never fails the publish it follows - logged.
        // ###########################################################################################
        private static async Task<int> SettleOtherCopiesAsync(
            OscilloscopeModelSubmissionRecord published,
            IOscilloscopeModelStore store,
            IAccountStore accounts,
            IEmailSender mailer,
            string name,
            ILogger? logger)
        {
            IReadOnlyList<OscilloscopeModelSubmissionRecord> replaced;

            try
            {
                replaced = await store.ReplaceInBetaOfNameAsync(published.Vendor, published.SeriesOrModel, published.Id, CancellationToken.None);
            }
            catch (Exception ex)
            {
                logger?.LogWarning(ex, "The other copies of oscilloscope model {Id} in BETA could not be marked replaced.", published.Id);
                return 0;
            }

            foreach (OscilloscopeModelSubmissionRecord copy in replaced.Where(copy => !OscilloscopeModelFlows.SameContributor(copy, published)))
            {
                await OscilloscopeModelFlows.TellContributorAsync(
                    copy,
                    accounts,
                    mailer,
                    (address, contributorName) => EmailTemplates.OscilloscopeModelSuperseded(address, name, contributorName),
                    logger);
            }

            return replaced.Count;
        }

        // The same contributor sent both: the same account, or - neither signed in - the same address.
        private static bool SameContributor(OscilloscopeModelSubmissionRecord one, OscilloscopeModelSubmissionRecord other) =>
            one.AccountId is long account
                ? other.AccountId == account
                : other.AccountId is null && !string.IsNullOrWhiteSpace(one.ContactEmail) &&
                  string.Equals(one.ContactEmail?.Trim(), other.ContactEmail?.Trim(), StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // POST /api/admin/oscilloscope-models/reject - with the reason the contributor is mailed.
        // ###########################################################################################
        public static async Task<OscilloscopeModelDecisionOutcome> RejectAsync(
            ReviewAccess access,
            OscilloscopeModelDecisionRequest? request,
            IOscilloscopeModelStore store,
            IAccountStore accounts,
            IEmailSender mailer,
            PublishLock publishLock,
            DateTimeOffset nowUtc,
            ILogger? logger = null,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(access);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(publishLock);

            if (request is null)
                return OscilloscopeModelDecisionOutcome.NotFound(OscilloscopeModelFlows.NoSuchModel);

            string reason = request.Reason?.Trim() ?? string.Empty;

            if (reason.Length == 0)
                return OscilloscopeModelDecisionOutcome.BadRequest(OscilloscopeModelFlows.NoReason);

            if (reason.Length > OscilloscopeModelRules.MessageLimit)
                return OscilloscopeModelDecisionOutcome.BadRequest($"The reason is longer than {OscilloscopeModelRules.MessageLimit} characters.");

            OscilloscopeModelSubmissionRecord? record;

            // Under the publish lock (code review, 2026-10-29): a reject landing while the model's
            // publish ran left it published with the record saying rejected, and both mails sent.
            using (await publishLock.EnterAsync(cancellationToken).ConfigureAwait(false))
            {
                record = await store.FindAsync(request.Id, cancellationToken);

                if (record is null)
                    return OscilloscopeModelDecisionOutcome.NotFound(OscilloscopeModelFlows.NoSuchModel);

                // In BETA it is, and stays: rejected, the record said "rejected" and the contributor
                // was mailed "it will not be going in" while every BETA user kept downloading it (code
                // review, 2026-10-29). Only a publish marks a model so, under this same lock.
                if (record.State == OscilloscopeModelStates.InBeta)
                    return OscilloscopeModelDecisionOutcome.Conflict(
                        OscilloscopeModelFlows.InBetaNotRejected(OscilloscopeModelFlows.NameOf(record.Vendor, record.SeriesOrModel)));

                if (!await store.DecideAsync(record.Id, OscilloscopeModelStates.Rejected, access.Account.Id, reason, nowUtc, cancellationToken: cancellationToken))
                    return OscilloscopeModelDecisionOutcome.Conflict(OscilloscopeModelFlows.AlreadyDecided);
            }

            string name = OscilloscopeModelFlows.NameOf(record.Vendor, record.SeriesOrModel);

            await OscilloscopeModelFlows.AuditAsync(accounts, access, OscilloscopeModelFlows.RejectedAction, record, reason, nowUtc, logger);
            await OscilloscopeModelFlows.TellContributorAsync(
                record,
                accounts,
                mailer,
                (address, contributorName) => EmailTemplates.OscilloscopeModelRejected(address, name, reason, contributorName),
                logger);

            return OscilloscopeModelDecisionOutcome.Done(new OscilloscopeModelDecisionAnswer(
                false,
                null,
                $"{name} will not go in - the contributor is told why."));
        }

        // ###########################################################################################
        // How many models wait for `access` - the administrator decides them alone, so nobody else has
        // any. Sent with the review queue (ReviewQueueAnswer.OscilloscopeModelsWaiting), which the
        // Maintainer tab reads every minute: its badge counts them, and "hide the tab while no work
        // is waiting" no longer hides the models (code review, 2026-10-29).
        // ###########################################################################################
        public static async Task<int> WaitingForAsync(ReviewAccess access, IOscilloscopeModelStore store, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(access);
            ArgumentNullException.ThrowIfNull(store);

            return access.Account.IsAdministrator ? await store.CountWaitingAsync(cancellationToken) : 0;
        }

        // What a publish says of the earlier copies of its name it took the place of in BETA.
        private static string SettledLine(int settled) => settled switch
        {
            0 => string.Empty,
            1 => " An earlier copy of it that was in BETA only is replaced by this one, and leaves the list.",
            _ => $" {settled} earlier copies of it that were in BETA only are replaced by this one, and leave the list.",
        };

        // "Rigol DS1054Z" - how a model is named in the mails and answers.
        public static string NameOf(string vendor, string seriesOrModel) => $"{vendor.Trim()} {seriesOrModel.Trim()}";

        private static string? BetaRoot(ServerOptions options) =>
            ManifestRebuildFlow.TreesToRebuild(options).FirstOrDefault(tree => tree.Name == UnusedFileFlows.Beta)?.Root ?? options.DataTreeRoot;

        // The models BETA's main file publishes - none when it has none, or cannot be read (the list
        // shows the model as new then; the publish reads the file itself, and says what is wrong).
        private static IReadOnlyList<MainDocument.OscilloscopeRow> PublishedModels(string? root)
        {
            try
            {
                return MasterListing.NewestMasterPath(root) is string main && MainFileListing.IsMainFile(main)
                    ? DataTreeJson.ReadMain(File.ReadAllText(main)).Oscilloscopes
                    : [];
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
            {
                return [];
            }
        }

        private static IReadOnlyDictionary<string, string> ValuesOf(OscilloscopeModelSubmissionRecord record)
        {
            try
            {
                return JsonSerializer.Deserialize<Dictionary<string, string>>(record.ValuesJson) ?? new Dictionary<string, string>();
            }
            catch (JsonException)
            {
                return new Dictionary<string, string>();
            }
        }

        // A fresh model of the record's each time: the main file's write takes it over.
        private static MainDocument.OscilloscopeRow ModelOf(OscilloscopeModelSubmissionRecord record) =>
            OscilloscopeModelValues.ToModel(record.Vendor, record.SeriesOrModel, OscilloscopeModelFlows.ValuesOf(record));

        private static async Task AuditAsync(
            IAccountStore accounts,
            ReviewAccess access,
            string action,
            OscilloscopeModelSubmissionRecord record,
            string detail,
            DateTimeOffset nowUtc,
            ILogger? logger)
        {
            try
            {
                await accounts.WriteAuditAsync(
                    new AuditEntry(
                        access.Account.Id,
                        access.Account.Email,
                        action,
                        $"oscilloscope-model:{record.Id}",
                        $"{OscilloscopeModelFlows.NameOf(record.Vendor, record.SeriesOrModel)} | {detail}",
                        nowUtc),
                    CancellationToken.None);
            }
            catch (Exception ex)
            {
                // The decision stands - an audit that failed must not turn it into a 500.
                logger?.LogError(ex, "The decision on oscilloscope model {Id} could not be audited.", record.Id);
            }
        }

        // The contributor, at the account's address when it was sent signed in (the board rule,
        // ContributorAddresses), else at the address typed. Nothing escapes.
        private static async Task TellContributorAsync(
            OscilloscopeModelSubmissionRecord record,
            IAccountStore accounts,
            IEmailSender mailer,
            Func<string, string?, EmailMessage> build,
            ILogger? logger)
        {
            try
            {
                AccountRecord? account = record.AccountId is long accountId
                    ? await accounts.FindByIdAsync(accountId, CancellationToken.None)
                    : null;

                string address = account?.Email ?? record.ContactEmail ?? string.Empty;

                if (string.IsNullOrWhiteSpace(address))
                    return;

                await OscilloscopeModelFlows.SendAsync(mailer, () => build(address, account?.DisplayName ?? record.ContactName), logger);
            }
            catch (Exception ex)
            {
                logger?.LogWarning(ex, "The contributor of oscilloscope model {Id} could not be told.", record.Id);
            }
        }

        private static async Task SendAsync(IEmailSender mailer, Func<EmailMessage> build, ILogger? logger)
        {
            try
            {
                await mailer.SendAsync(build(), CancellationToken.None);
            }
            catch (Exception ex)
            {
                logger?.LogWarning(ex, "A mail about an oscilloscope model could not be sent.");
            }
        }
    }

    public sealed record OscilloscopeModelSubmitOutcome(OscilloscopeModelSubmissionAnswer? Answer, IReadOnlyList<string> Problems, bool IsRateLimited)
    {
        public static OscilloscopeModelSubmitOutcome Accepted(OscilloscopeModelSubmissionAnswer answer) => new(answer, [], false);

        public static OscilloscopeModelSubmitOutcome Refused(IReadOnlyList<string> problems) => new(null, problems, false);

        public static OscilloscopeModelSubmitOutcome RateLimited(string problem) => new(null, [problem], true);
    }

    public sealed record OscilloscopeModelDecisionOutcome(OscilloscopeModelDecisionAnswer? Answer, string? Error, int Status)
    {
        public static OscilloscopeModelDecisionOutcome Done(OscilloscopeModelDecisionAnswer answer) => new(answer, null, StatusCodes.Status200OK);

        public static OscilloscopeModelDecisionOutcome BadRequest(string error) => new(null, error, StatusCodes.Status400BadRequest);

        public static OscilloscopeModelDecisionOutcome NotFound(string error) => new(null, error, StatusCodes.Status404NotFound);

        public static OscilloscopeModelDecisionOutcome Conflict(string error) => new(null, error, StatusCodes.Status409Conflict);
    }
}
