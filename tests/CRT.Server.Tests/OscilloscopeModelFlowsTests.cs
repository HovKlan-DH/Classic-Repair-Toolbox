using System.Text.Json;
using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Http;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // OscilloscopeModelFlows - an oscilloscope model sent in from CRT's Oscilloscope tab, and the
    // administrator publishing it for everybody or turning it down (CRT 4.0 step 5). Two real trees,
    // each with a real CRT 4.0 main file; the manifest writer records which trees it was asked to
    // rebuild, the store and the mailer are fakes.
    // ###########################################################################################
    public sealed class OscilloscopeModelFlowsTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 29, 12, 0, 0, TimeSpan.Zero);

        private static readonly MasterListingRow C64 =
            new("Commodore 64", "250407 (long board)", "Commodore/C64/250407/board.json", string.Empty);

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-scope-models", Guid.NewGuid().ToString("N"));

        private readonly FakeOscilloscopeModelStore thisStore = new();

        private readonly FakeAccountStore thisAccounts = new();

        private readonly FakeEmailSender thisMailer = new();

        private readonly List<string> thisRebuilt = [];

        private string Beta => Path.Combine(this.thisRoot, "beta");

        private string Stable => Path.Combine(this.thisRoot, "stable");

        public OscilloscopeModelFlowsTests()
        {
            Directory.CreateDirectory(this.Beta);
            Directory.CreateDirectory(this.Stable);

            DataTreeBuilder.ListingMaster(this.Beta, C64);
            DataTreeBuilder.ListingMaster(this.Stable, C64);
            this.PublishByHand(this.Beta, Model("Rigol", "DS1000Z", ":DISP:DATA?"));
            this.PublishByHand(this.Stable, Model("Rigol", "DS1000Z", ":DISP:DATA?"));

            this.thisAccounts.Accounts[7] = new AccountRecord(
                7, "dennis@example.com", "dennis@example.com", "hash", "Dennis",
                IsVerified: true, IsAdministrator: true, IsLocked: false, Now, null);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        private static MainDocument.OscilloscopeRow Model(string vendor, string model, string dumpImage = ":DISP:DATA? BMP") => new()
        {
            Brand = vendor,
            SeriesOrModel = model,
            Port = "5025",
            Identify = "*IDN?",
            Stop = ":STOP",
            DumpImage = dumpImage,
            TimeDivList = "1ms,2ms",
        };

        private void PublishByHand(string root, MainDocument.OscilloscopeRow model)
        {
            string path = DataTreeBuilder.Full(root, DataTreeLayout.MainFileName);
            MainDocument main = DataTreeJson.ReadMain(File.ReadAllText(path));
            main.Oscilloscopes.Add(model);
            File.WriteAllText(path, DataTreeJson.Write(main));
        }

        private static IReadOnlyList<MainDocument.OscilloscopeRow> ModelsIn(string root) =>
            DataTreeJson.ReadMain(File.ReadAllText(DataTreeBuilder.Full(root, DataTreeLayout.MainFileName))).Oscilloscopes;

        private ServerOptions Options(bool stable = true)
        {
            var options = new ServerOptions
            {
                DataTreeRoot = this.Beta,
                ManifestPath = Path.Combine(this.thisRoot, "beta.json"),
                PublicDataBaseUrl = "https://example.invalid/app-data-BETA/",
            };

            if (stable)
            {
                options.ProductionDataTreeRoot = this.Stable;
                options.ProductionManifestPath = Path.Combine(this.thisRoot, "stable.json");
                options.ProductionPublicDataBaseUrl = "https://example.invalid/app-data/";
            }

            return options;
        }

        private ReviewAccess Administrator() => ReviewAccess.For(this.thisAccounts.Accounts[7]);

        private static OscilloscopeModelSubmissionRequest Request(
            MainDocument.OscilloscopeRow model,
            string? contactEmail = "anna@example.com",
            string? message = "My Voltcraft, firmware 1.2",
            bool passed = true) =>
            new(model.Brand, model.SeriesOrModel, OscilloscopeModelValues.Of(model), contactEmail, "Anna", message, passed);

        private Task<OscilloscopeModelSubmitOutcome> Submit(
            OscilloscopeModelSubmissionRequest request,
            AccountRecord? account = null,
            string? address = "203.0.113.5") =>
            OscilloscopeModelFlows.SubmitAsync(request, account, address, "CRT 4.0.0-alpha.3", this.thisStore, this.thisAccounts, this.thisMailer, Now);

        private Task<OscilloscopeModelDecisionOutcome> Publish(long id, ServerOptions? options = null, bool manifestsRebuilt = true) =>
            OscilloscopeModelFlows.PublishAsync(
                this.Administrator(),
                new OscilloscopeModelDecisionRequest(id, null),
                this.thisStore,
                this.thisAccounts,
                this.thisMailer,
                options ?? this.Options(),
                tree =>
                {
                    this.thisRebuilt.Add(tree.Name);
                    return manifestsRebuilt ? 1 : throw new IOException("The manifest is held open.");
                },
                new PublishLock(),
                Now);

        // The stable source with no main file - it cannot take a model.
        private void BreakStable() => File.Delete(DataTreeBuilder.Full(this.Stable, DataTreeLayout.MainFileName));

        private Task<OscilloscopeModelDecisionOutcome> Reject(long id, string? reason) =>
            OscilloscopeModelFlows.RejectAsync(
                this.Administrator(),
                new OscilloscopeModelDecisionRequest(id, reason),
                this.thisStore,
                this.thisAccounts,
                this.thisMailer,
                new PublishLock(),
                Now);

        // ------------------------------------------------------------------ sending one in

        [Fact]
        public async Task A_model_sent_in_waits_for_the_administrator_who_is_mailed()
        {
            OscilloscopeModelSubmitOutcome outcome = await this.Submit(Request(Model("Voltcraft", "DOV702")));

            OscilloscopeModelSubmissionAnswer answer = Assert.IsType<OscilloscopeModelSubmissionAnswer>(outcome.Answer);
            OscilloscopeModelSubmissionRecord kept = this.thisStore[answer.Id];

            Assert.Equal(OscilloscopeModelStates.Pending, kept.State);
            Assert.Equal("anna@example.com", kept.ContactEmail);
            Assert.True(kept.FullTestSuitePassed);
            Assert.Equal("CRT 4.0.0-alpha.3", kept.ClientVersion);
            Assert.Equal(":DISP:DATA? BMP", JsonSerializer.Deserialize<Dictionary<string, string>>(kept.ValuesJson)!["dumpImage"]);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Equal("dennis@example.com", mail.ToAddress);
            Assert.Contains("Voltcraft DOV702", mail.Body);
            Assert.Contains("My Voltcraft, firmware 1.2", mail.Body);
        }

        [Fact]
        public async Task A_model_with_a_problem_is_refused_and_nothing_is_kept()
        {
            MainDocument.OscilloscopeRow noIdentify = Model("Voltcraft", "DOV702");
            noIdentify.Identify = " ";
            MainDocument.OscilloscopeRow twoLines = Model("Voltcraft", "DOV702");
            twoLines.Stop = ":STOP\n:RUN";

            OscilloscopeModelSubmitOutcome missing = await this.Submit(Request(noIdentify));
            OscilloscopeModelSubmitOutcome broken = await this.Submit(Request(twoLines));
            OscilloscopeModelSubmitOutcome unknown = await this.Submit(Request(Model("Voltcraft", "DOV702")) with
            {
                Values = new Dictionary<string, string>(OscilloscopeModelValues.Of(Model("Voltcraft", "DOV702"))) { ["launchMissiles"] = "*RST" },
            });

            Assert.Contains(OscilloscopeModelRules.NoIdentify, missing.Problems);
            Assert.Contains(broken.Problems, problem => problem.Contains("\"stop\""));
            Assert.Contains(unknown.Problems, problem => problem.Contains("launchMissiles"));
            Assert.Empty(this.thisStore.Rows);
            Assert.Empty(this.thisMailer.Sent);
        }

        // As for a board: an address to be told at, or the account sent with.
        [Fact]
        public async Task A_model_needs_an_address_unless_it_is_sent_signed_in()
        {
            OscilloscopeModelSubmitOutcome none = await this.Submit(Request(Model("Voltcraft", "DOV702"), contactEmail: " "));
            OscilloscopeModelSubmitOutcome bad = await this.Submit(Request(Model("Voltcraft", "DOV702"), contactEmail: "anna"));
            OscilloscopeModelSubmitOutcome signedIn = await this.Submit(
                Request(Model("Voltcraft", "DOV702"), contactEmail: null),
                account: this.thisAccounts.Accounts[7]);

            Assert.Equal([OscilloscopeModelFlows.NoContact], none.Problems);
            Assert.Equal([OscilloscopeModelFlows.BadContact], bad.Problems);
            Assert.NotNull(signedIn.Answer);
            Assert.Equal(7, this.thisStore[signedIn.Answer!.Id].AccountId);
        }

        [Fact]
        public async Task An_address_may_send_ten_models_a_day()
        {
            for (int number = 1; number <= OscilloscopeModelFlows.MaxPerAddress; number++)
                Assert.NotNull((await this.Submit(Request(Model("Voltcraft", $"DOV70{number}")))).Answer);

            OscilloscopeModelSubmitOutcome eleventh = await this.Submit(Request(Model("Voltcraft", "DOV800")));
            OscilloscopeModelSubmitOutcome elsewhere = await this.Submit(Request(Model("Voltcraft", "DOV800")), address: "198.51.100.9");

            Assert.True(eleventh.IsRateLimited);
            Assert.NotNull(elsewhere.Answer);

            // Counted by the internet connection, not the email address typed - and said so (code
            // review, 2026-10-29: "This address" read as the email address, a fresh one refused).
            string said = Assert.Single(eleventh.Problems);
            Assert.Contains("internet connection", said);
            Assert.DoesNotContain("This address", said);
        }

        // ###########################################################################################
        // *** THE DAILY LIMIT HOLDS FOR A BURST (code review, 2026-10-29 - Security). *** It was a count
        // followed by an insert with nothing between them, so requests sent at once from one connection
        // all counted fewer than ten - each one kept, and each mailed to every administrator.
        // ###########################################################################################
        [Fact]
        public async Task A_burst_from_one_connection_keeps_no_more_than_the_daily_limit()
        {
            this.thisStore.WhileCounting = () => Task.Delay(10);

            OscilloscopeModelSubmitOutcome[] outcomes = await Task.WhenAll(
                Enumerable.Range(1, 25).Select(number => Task.Run(() => this.Submit(Request(Model("Voltcraft", $"DOV{number}"))))));

            Assert.Equal(OscilloscopeModelFlows.MaxPerAddress, outcomes.Count(outcome => outcome.Answer is not null));
            Assert.Equal(OscilloscopeModelFlows.MaxPerAddress, this.thisStore.Rows.Count);
        }

        // ###########################################################################################
        // *** AT MOST A HUNDRED WAIT AT ONCE, FROM EVERYBODY (code review, 2026-10-29 - Security). ***
        // The daily limit is per connection, and a whole IPv6 prefix holds thousands of /64s - every
        // model kept mailed to every administrator. The rows and the mails are bounded until decided.
        // ###########################################################################################
        [Fact]
        public async Task No_more_than_a_hundred_models_wait_at_once()
        {
            for (int number = 0; number < OscilloscopeModelFlows.MaxWaiting; number++)
            {
                Assert.NotNull((await this.Submit(
                    Request(Model("Voltcraft", $"DOV{number}")),
                    address: $"2001:db8:{number / OscilloscopeModelFlows.MaxPerAddress:x}::1")).Answer);
            }

            OscilloscopeModelSubmitOutcome refused = await this.Submit(Request(Model("Owon", "ADS800")), address: "2001:db8:ffff::1");

            Assert.True(refused.IsRateLimited);
            Assert.Equal([OscilloscopeModelFlows.TooManyWaiting], refused.Problems);

            // Decided, there is room again.
            await this.Reject(1, "Not this one.");
            Assert.NotNull((await this.Submit(Request(Model("Owon", "ADS800")), address: "2001:db8:ffff::1")).Answer);
        }

        // An IPv6 connection is its /64 network - a sender could otherwise take a new address of its
        // own network for every ten (code review, 2026-10-29 - Security). IPv4, and IPv4 inside IPv6,
        // count by the address.
        [Theory]
        [InlineData("2001:db8:1:2:aaaa:bbbb:cccc:dddd", "2001:db8:1:2::/64")]
        [InlineData("2001:db8:1:2::1", "2001:db8:1:2::/64")]
        [InlineData("203.0.113.5", "203.0.113.5")]
        [InlineData("::ffff:203.0.113.5", "203.0.113.5")]
        [InlineData(null, null)]
        [InlineData("  ", null)]
        public void An_internet_connection_is_an_IPv4_address_or_an_IPv6_network(string? address, string? connection)
        {
            Assert.Equal(connection, OscilloscopeModelFlows.ConnectionOf(address));
        }

        [Fact]
        public async Task Addresses_of_one_IPv6_network_share_the_daily_limit()
        {
            for (int number = 1; number <= OscilloscopeModelFlows.MaxPerAddress; number++)
                Assert.NotNull((await this.Submit(Request(Model("Voltcraft", $"DOV70{number}")), address: $"2001:db8:1:2::{number:x}")).Answer);

            Assert.True((await this.Submit(Request(Model("Voltcraft", "DOV800")), address: "2001:db8:1:2::ff")).IsRateLimited);
            Assert.NotNull((await this.Submit(Request(Model("Voltcraft", "DOV800")), address: "2001:db8:1:3::1")).Answer);
        }

        // ###########################################################################################
        // *** A TYPED ADDRESS ALONE NEVER TAKES ANOTHER'S MODEL OUT OF THE LIST (code review,
        // 2026-10-29 - Security). *** Anybody could type anna@example.com: a model of the same name sent
        // so marked Anna's replaced, and the administrator never saw it. Without an account, a newer
        // copy replaces one only from the same address AND the same internet connection.
        // ###########################################################################################
        [Fact]
        public async Task A_typed_address_from_another_connection_replaces_nothing()
        {
            long anna = (await this.Submit(Request(Model("Rigol", "DS1054Z+")))).Answer!.Id;

            long other = (await this.Submit(Request(Model("Rigol", "DS1054Z+", ":DISP:DATA? PNG")), address: "198.51.100.9")).Answer!.Id;

            Assert.Equal(OscilloscopeModelStates.Pending, this.thisStore[anna].State);
            Assert.Equal(OscilloscopeModelStates.Pending, this.thisStore[other].State);

            // From Anna's own connection, it is her newer copy.
            long again = (await this.Submit(Request(Model("Rigol", "DS1054Z+", ":DISP:DATA? JPG")))).Answer!.Id;

            Assert.Equal(OscilloscopeModelStates.Replaced, this.thisStore[anna].State);
            Assert.Equal(OscilloscopeModelStates.Pending, this.thisStore[other].State);
            Assert.Equal(OscilloscopeModelStates.Pending, this.thisStore[again].State);
        }

        // ###########################################################################################
        // *** NEVER HALF A CHARACTER IN A COLUMN (code review, 2026-10-29). *** The flow's own Clip cut
        // a name at 200 UTF-16 units, so an emoji across the cut left a lone high surrogate, which
        // MariaDB's utf8mb4 refuses - the model was refused with a 500. The shared helper steps back.
        // ###########################################################################################
        [Fact]
        public async Task A_long_name_is_kept_without_splitting_a_character()
        {
            OscilloscopeModelSubmissionRequest request = Request(Model("Voltcraft", "DOV702")) with
            {
                ContactName = new string('a', 199) + "\U0001F600 Anna",
            };

            long id = (await this.Submit(request)).Answer!.Id;

            Assert.Equal(new string('a', 199), this.thisStore[id].ContactName);
        }

        // The model is kept: a mailer that refuses changes nothing the contributor sees.
        [Fact]
        public async Task A_mail_that_does_not_go_never_fails_the_sending()
        {
            this.thisMailer.Delivers = false;

            Assert.NotNull((await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer);
        }

        // ------------------------------------------------------------------ the administrator's list

        [Fact]
        public async Task The_list_shows_each_waiting_model_with_the_published_one_of_its_name()
        {
            OscilloscopeModelSubmitOutcome update = await this.Submit(Request(Model("Rigol", "DS1000Z", ":DISP:DATA? PNG")));
            OscilloscopeModelSubmitOutcome fresh = await this.Submit(
                Request(Model("Voltcraft", "DOV702"), contactEmail: null),
                account: this.thisAccounts.Accounts[7]);

            OscilloscopeModelQueueAnswer list = await OscilloscopeModelFlows.ListAsync(this.thisStore, this.thisAccounts, this.Options());

            Assert.Equal([update.Answer!.Id, fresh.Answer!.Id], list.Entries.Select(entry => entry.Id));

            OscilloscopeModelQueueEntry first = list.Entries[0];
            Assert.Equal(":DISP:DATA? PNG", first.Values["dumpImage"]);
            Assert.Equal(":DISP:DATA?", first.PublishedValues!["dumpImage"]);
            Assert.Equal("Anna", first.ContactName);

            // Sent signed in: the account's name and address.
            Assert.Null(list.Entries[1].PublishedValues);
            Assert.Equal("Dennis", list.Entries[1].ContactName);
            Assert.Equal("dennis@example.com", list.Entries[1].ContactEmail);
        }

        // ------------------------------------------------------------------ publishing

        [Fact]
        public async Task Publish_puts_the_model_into_both_main_files_and_tells_the_contributor()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            this.thisMailer.Sent.Clear();

            OscilloscopeModelDecisionOutcome outcome = await this.Publish(id);

            OscilloscopeModelDecisionAnswer answer = Assert.IsType<OscilloscopeModelDecisionAnswer>(outcome.Answer);
            Assert.True(answer.BetaChanged);
            Assert.True(answer.StableChanged);
            Assert.Null(answer.Problem);

            foreach (string root in new[] { this.Beta, this.Stable })
            {
                Assert.Equal(["DS1000Z", "DOV702"], ModelsIn(root).Select(model => model.SeriesOrModel));
                Assert.Equal(":DISP:DATA? BMP", ModelsIn(root)[1].DumpImage);
            }

            // Each tree's manifest rebuilt, so CRT downloads the main file at its next launch.
            Assert.Equal([UnusedFileFlows.Beta, UnusedFileFlows.Production], this.thisRebuilt);
            Assert.Equal(OscilloscopeModelStates.Published, this.thisStore[id].State);
            Assert.Equal(OscilloscopeModelFlows.PublishedAction, Assert.Single(this.thisAccounts.Audit).Action);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.Equal("anna@example.com", mail.ToAddress);
            Assert.Contains("published", mail.Subject);
            Assert.Contains("next data check at launch", mail.Body);
        }

        // ###########################################################################################
        // *** A MANIFEST NOT REBUILT MAKES NO PROMISE OF THE NEXT LAUNCH (code review, 2026-10-29). ***
        // CRT throws a download away while the manifest names the old checksum, so the mail's "everybody
        // gets it with the next data check at launch" was not so until the manifests were rebuilt by hand.
        // ###########################################################################################
        [Fact]
        public async Task A_manifest_not_rebuilt_makes_the_mail_promise_no_next_launch()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            this.thisMailer.Sent.Clear();

            OscilloscopeModelDecisionAnswer answer = (await this.Publish(id, manifestsRebuilt: false)).Answer!;

            Assert.Contains("Rebuild checksum manifests", answer.Problem);
            Assert.Equal(OscilloscopeModelStates.Published, this.thisStore[id].State);

            // The administrator's sentence promises no more than the mail (code review, 2026-10-29).
            Assert.DoesNotContain("next launch", answer.Message);
            Assert.Contains("once the checksum manifests are rebuilt", answer.Message);

            EmailMessage mail = Assert.Single(this.thisMailer.Sent);
            Assert.DoesNotContain("next data check at launch", mail.Body);
            Assert.Contains("a little later", mail.Body);
        }

        // An update of a published model takes its place - one model of the name, where it was.
        [Fact]
        public async Task A_model_of_a_published_name_replaces_it_where_it_is()
        {
            this.PublishByHand(this.Beta, Model("Siglent", "SDS1000X-E"));
            long id = (await this.Submit(Request(Model("rigol", "DS1000Z", ":DISP:DATA? PNG")))).Answer!.Id;

            await this.Publish(id);

            IReadOnlyList<MainDocument.OscilloscopeRow> beta = ModelsIn(this.Beta);
            Assert.Equal(["DS1000Z", "SDS1000X-E"], beta.Select(model => model.SeriesOrModel));
            Assert.Equal(":DISP:DATA? PNG", beta[0].DumpImage);
        }

        [Fact]
        public async Task A_model_decided_already_or_unknown_is_not_published()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            await this.Reject(id, "Please send the DumpImage command too.");

            OscilloscopeModelDecisionOutcome decided = await this.Publish(id);
            OscilloscopeModelDecisionOutcome unknown = await this.Publish(999);

            Assert.Equal(StatusCodes.Status409Conflict, decided.Status);
            Assert.Equal(StatusCodes.Status404NotFound, unknown.Status);
            Assert.Single(ModelsIn(this.Beta));
            Assert.Empty(this.thisRebuilt);
        }

        [Fact]
        public async Task Without_a_stable_source_only_BETA_takes_it()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;

            OscilloscopeModelDecisionAnswer answer = (await this.Publish(id, this.Options(stable: false))).Answer!;

            Assert.True(answer.BetaChanged);
            Assert.Null(answer.StableChanged);
            Assert.Equal(2, ModelsIn(this.Beta).Count);
            Assert.Single(ModelsIn(this.Stable));
        }

        // ###########################################################################################
        // *** A STABLE SOURCE THAT COULD NOT TAKE IT LEAVES THE MODEL IN THE LIST, "IN BETA" (code
        // reviews, 2026-10-29). *** BETA is written first and never rolled back; the model was then
        // marked published all the same, so pressing Publish again - what the answer told the
        // administrator to do - got "decided meanwhile", and stable never got the model. Now it stays
        // listed as in BETA, its contributor not mailed yet, the BETA write audited, and Publish again
        // finishes it.
        // ###########################################################################################
        [Fact]
        public async Task A_stable_source_that_could_not_take_it_leaves_the_model_in_BETA_for_Publish_again()
        {
            string stableMain = DataTreeBuilder.Full(this.Stable, DataTreeLayout.MainFileName);
            byte[] kept = File.ReadAllBytes(stableMain);
            this.BreakStable();
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            this.thisMailer.Sent.Clear();

            OscilloscopeModelDecisionAnswer first = (await this.Publish(id)).Answer!;

            Assert.True(first.BetaChanged);
            Assert.False(first.StableChanged);
            Assert.Contains("stable source", first.Problem);
            Assert.Contains("Publish again", first.Problem);
            Assert.Equal(OscilloscopeModelStates.InBeta, this.thisStore[id].State);
            Assert.Empty(this.thisMailer.Sent);
            Assert.Equal(OscilloscopeModelFlows.InBetaAction, Assert.Single(this.thisAccounts.Audit).Action);

            OscilloscopeModelQueueEntry listed = Assert.Single((await OscilloscopeModelFlows.ListAsync(this.thisStore, this.thisAccounts, this.Options())).Entries);
            Assert.Equal(id, listed.Id);
            Assert.True(listed.InBeta);

            File.WriteAllBytes(stableMain, kept);
            OscilloscopeModelDecisionAnswer second = (await this.Publish(id)).Answer!;

            Assert.False(second.BetaChanged);
            Assert.True(second.StableChanged);
            Assert.Null(second.Problem);
            Assert.Equal(OscilloscopeModelStates.Published, this.thisStore[id].State);
            Assert.Single(this.thisMailer.Sent);
            Assert.Equal(2, ModelsIn(this.Stable).Count);
            Assert.Equal(OscilloscopeModelFlows.PublishedAction, this.thisAccounts.Audit[^1].Action);
        }

        // ###########################################################################################
        // *** A MODEL IN BETA CANNOT BE REJECTED (code review, 2026-10-29). *** Rejected, the record said
        // "rejected" and the contributor was mailed "it will not be going in" while every BETA user kept
        // downloading it. It cannot be taken back out of BETA, so the one way on is Publish again.
        // ###########################################################################################
        [Fact]
        public async Task A_model_in_BETA_cannot_be_rejected()
        {
            this.BreakStable();
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            await this.Publish(id);
            this.thisMailer.Sent.Clear();

            OscilloscopeModelDecisionOutcome rejected = await this.Reject(id, "The screen picture command is wrong.");

            Assert.Equal(StatusCodes.Status409Conflict, rejected.Status);
            Assert.Contains("Publish again", rejected.Error);
            Assert.Equal(OscilloscopeModelStates.InBeta, this.thisStore[id].State);
            Assert.Empty(this.thisMailer.Sent);
        }

        // A newer copy from the same contributor never takes the place of one already in BETA - it
        // would leave the list while BETA keeps it. Both wait: the one in BETA to be finished, the
        // newer one to be decided.
        [Fact]
        public async Task A_newer_copy_never_replaces_a_model_already_in_BETA()
        {
            this.BreakStable();
            long first = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            await this.Publish(first);

            long newer = (await this.Submit(Request(Model("Voltcraft", "DOV702", ":DISP:DATA? PNG")))).Answer!.Id;

            Assert.Equal(OscilloscopeModelStates.InBeta, this.thisStore[first].State);
            Assert.Equal(
                [first, newer],
                (await OscilloscopeModelFlows.ListAsync(this.thisStore, this.thisAccounts, this.Options())).Entries.Select(entry => entry.Id));
        }

        // ###########################################################################################
        // *** A MODEL OF ITS NAME WRITTEN INTO BETA SETTLES AN EARLIER ONE STILL IN BETA (code review,
        // 2026-10-29). *** The earlier copy stayed listed with "press Publish again to finish it" and no
        // Reject, and that Publish wrote its OLDER commands over the newer model in both trees. BETA no
        // longer holds it, so it is replaced - out of the list, and Publish on it is refused.
        // ###########################################################################################
        [Fact]
        public async Task Publishing_a_newer_copy_settles_the_older_one_still_in_BETA()
        {
            string stableMain = DataTreeBuilder.Full(this.Stable, DataTreeLayout.MainFileName);
            byte[] kept = File.ReadAllBytes(stableMain);
            this.BreakStable();
            long older = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            await this.Publish(older);
            long newer = (await this.Submit(Request(Model("Voltcraft", "DOV702", ":DISP:DATA? PNG")))).Answer!.Id;

            File.WriteAllBytes(stableMain, kept);
            Assert.Contains("earlier copy", (await this.Publish(newer)).Answer!.Message);

            Assert.Equal(OscilloscopeModelStates.Replaced, this.thisStore[older].State);
            Assert.Empty((await OscilloscopeModelFlows.ListAsync(this.thisStore, this.thisAccounts, this.Options())).Entries);
            Assert.Equal(StatusCodes.Status409Conflict, (await this.Publish(older)).Status);

            foreach (string root in new[] { this.Beta, this.Stable })
                Assert.Equal(":DISP:DATA? PNG", ModelsIn(root).Single(model => model.SeriesOrModel == "DOV702").DumpImage);

            // Its own newer copy took its place: nothing to tell the contributor beyond "published".
            Assert.DoesNotContain(this.thisMailer.Sent, mail => mail.Subject.Contains("Another model", StringComparison.Ordinal));
        }

        // ###########################################################################################
        // *** ANOTHER CONTRIBUTOR'S COPY THAT LOSES ITS PLACE IN BETA IS TOLD SO (code review,
        // 2026-10-29). *** Anna's model reached BETA only, so she was not mailed yet; Bob's of the same
        // name was published, hers marked replaced - and she, told "you get a mail either way", never
        // heard of it again.
        // ###########################################################################################
        [Fact]
        public async Task Another_contributor_s_copy_that_loses_its_place_in_BETA_is_told_so()
        {
            string stableMain = DataTreeBuilder.Full(this.Stable, DataTreeLayout.MainFileName);
            byte[] kept = File.ReadAllBytes(stableMain);
            this.BreakStable();
            long anna = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            await this.Publish(anna);
            long bob = (await this.Submit(Request(Model("Voltcraft", "DOV702", ":DISP:DATA? PNG"), contactEmail: "bob@example.com"), address: "198.51.100.9")).Answer!.Id;
            this.thisMailer.Sent.Clear();

            File.WriteAllBytes(stableMain, kept);
            await this.Publish(bob);

            Assert.Equal(OscilloscopeModelStates.Replaced, this.thisStore[anna].State);
            EmailMessage told = Assert.Single(this.thisMailer.Sent, mail => mail.ToAddress == "anna@example.com");
            Assert.Contains("Another model", told.Subject);
            Assert.Contains("published", Assert.Single(this.thisMailer.Sent, mail => mail.ToAddress == "bob@example.com").Subject);
        }

        // ###########################################################################################
        // *** NAMES ARE THE MAIN FILE'S RULE, NOT THE DATABASE'S (code review, 2026-10-29). *** The
        // table's collation ignores accents, so "Pico" settled "Picó" - two models the main file keeps
        // apart. The store and this fake compare by OscilloscopeModelRules.SameName.
        // ###########################################################################################
        [Fact]
        public async Task A_name_that_differs_by_an_accent_is_another_model()
        {
            this.BreakStable();
            long accented = (await this.Submit(Request(Model("Picó", "2204A")))).Answer!.Id;
            await this.Publish(accented);
            long plain = (await this.Submit(Request(Model("Pico", "2204A")))).Answer!.Id;

            await this.Publish(plain);

            Assert.Equal(OscilloscopeModelStates.InBeta, this.thisStore[accented].State);
            Assert.Equal(OscilloscopeModelStates.InBeta, this.thisStore[plain].State);
        }

        // A database failing as the other copies are settled fails nothing already written: both trees
        // hold the model and it is recorded published (code review, 2026-10-29: unguarded, it ended the
        // publish halfway, BETA written and the record pending).
        [Fact]
        public async Task A_database_failing_to_settle_the_other_copies_fails_nothing_written()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            this.thisStore.FailReplaceInBeta = new InvalidOperationException("The database went away.");

            OscilloscopeModelDecisionOutcome outcome = await this.Publish(id);

            Assert.NotNull(outcome.Answer);
            Assert.Equal(OscilloscopeModelStates.Published, this.thisStore[id].State);
            Assert.Equal(2, ModelsIn(this.Stable).Count);
        }

        // The same while stable still cannot take it: BETA holds the newer one, which waits in its place.
        [Fact]
        public async Task A_newer_copy_that_also_reaches_only_BETA_takes_the_older_one_s_place()
        {
            this.BreakStable();
            long older = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            await this.Publish(older);
            long newer = (await this.Submit(Request(Model("Voltcraft", "DOV702", ":DISP:DATA? PNG")))).Answer!.Id;

            await this.Publish(newer);

            Assert.Equal(OscilloscopeModelStates.Replaced, this.thisStore[older].State);
            Assert.Equal(OscilloscopeModelStates.InBeta, this.thisStore[newer].State);
        }

        // A newer copy arriving WHILE the publish runs, and stable then not taking it: what was written
        // into BETA is in BETA all the same, so it stays listed as that.
        [Fact]
        public async Task A_model_replaced_while_a_publish_reached_only_BETA_is_listed_as_in_BETA()
        {
            this.BreakStable();
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            this.thisStore.BeforeMarkInBeta = () => this.Submit(Request(Model("Voltcraft", "DOV702", ":DISP:DATA? PNG"))).GetAwaiter().GetResult();

            await this.Publish(id);

            Assert.Equal(OscilloscopeModelStates.InBeta, this.thisStore[id].State);
        }

        // ###########################################################################################
        // *** A DECISION NEVER LANDS IN THE MIDDLE OF A PUBLISH (code review, 2026-10-29). *** Reject
        // did not take the publish lock: a model rejected while its publish ran was published with the
        // record saying "rejected", and its contributor got both mails. Reject now waits for it.
        // ###########################################################################################
        [Fact]
        public async Task Reject_waits_for_a_publish_under_way()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            var publishLock = new PublishLock();

            Task<OscilloscopeModelDecisionOutcome> rejecting;

            using (await publishLock.EnterAsync())
            {
                rejecting = OscilloscopeModelFlows.RejectAsync(
                    this.Administrator(),
                    new OscilloscopeModelDecisionRequest(id, "No picture command."),
                    this.thisStore,
                    this.thisAccounts,
                    this.thisMailer,
                    publishLock,
                    Now);

                await Task.Delay(50);
                Assert.False(rejecting.IsCompleted);
                Assert.Equal(OscilloscopeModelStates.Pending, this.thisStore[id].State);
            }

            Assert.NotNull((await rejecting).Answer);
            Assert.Equal(OscilloscopeModelStates.Rejected, this.thisStore[id].State);
        }

        // A newer copy of the same name from the same contributor marks the waiting one replaced - also
        // while that one's publish runs. What was written IS published, so the record says so.
        [Fact]
        public async Task A_model_replaced_by_a_newer_copy_while_its_publish_ran_is_recorded_published()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            this.thisStore.BeforeDecide = () => this.Submit(Request(Model("Voltcraft", "DOV702", ":DISP:DATA? PNG"))).GetAwaiter().GetResult();

            OscilloscopeModelDecisionOutcome outcome = await this.Publish(id);

            Assert.NotNull(outcome.Answer);
            Assert.Equal(OscilloscopeModelStates.Published, this.thisStore[id].State);
            Assert.Single(this.thisStore.Rows, row => row.Record.State == OscilloscopeModelStates.Pending);
        }

        // ###########################################################################################
        // *** THE MODELS WAITING ARE THE ADMINISTRATOR'S WORK TOO (code review, 2026-10-29). *** With
        // "Hide the Maintainer tab while no work is waiting" ticked the tab stayed hidden while models
        // waited - the badge counted the two board queues only. The queue answer now carries how many
        // wait: pending and in BETA alike, for the administrator alone.
        // ###########################################################################################
        [Fact]
        public async Task The_models_waiting_are_counted_for_the_administrator_alone()
        {
            this.BreakStable();
            long inBeta = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            await this.Publish(inBeta);
            await this.Submit(Request(Model("Owon", "ADS800")));
            long rejected = (await this.Submit(Request(Model("Rigol", "DS1054Z+")))).Answer!.Id;
            await this.Reject(rejected, "No picture command.");

            var maintainer = new AccountRecord(
                8, "anna@example.com", "anna@example.com", "hash", "Anna",
                IsVerified: true, IsAdministrator: false, IsLocked: false, Now, null);

            Assert.Equal(2, await OscilloscopeModelFlows.WaitingForAsync(this.Administrator(), this.thisStore));
            Assert.Equal(0, await OscilloscopeModelFlows.WaitingForAsync(ReviewAccess.For(maintainer), this.thisStore));
        }

        // ------------------------------------------------------------------ rejecting

        [Fact]
        public async Task Reject_needs_a_reason_and_mails_it()
        {
            long id = (await this.Submit(Request(Model("Voltcraft", "DOV702")))).Answer!.Id;
            this.thisMailer.Sent.Clear();

            OscilloscopeModelDecisionOutcome none = await this.Reject(id, "  ");

            Assert.Equal(StatusCodes.Status400BadRequest, none.Status);
            Assert.Equal(OscilloscopeModelStates.Pending, this.thisStore[id].State);

            OscilloscopeModelDecisionOutcome rejected = await this.Reject(id, "The screen picture command is missing.");

            Assert.NotNull(rejected.Answer);
            Assert.Equal(OscilloscopeModelStates.Rejected, this.thisStore[id].State);
            Assert.Equal("The screen picture command is missing.", this.thisStore.NoteOf(id));
            Assert.Contains("The screen picture command is missing.", Assert.Single(this.thisMailer.Sent).Body);
            Assert.Single(ModelsIn(this.Beta));

            Assert.Equal(StatusCodes.Status409Conflict, (await this.Reject(id, "Again.")).Status);
        }
    }
}
