using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SubmissionReplacementRules and the step in SubmissionFlows.FinaliseAsync that applies
    // them (owner decision, 2026-09-26): a newer submission from the same contributor replaces the
    // older, untouched one of the same board - "the newest one always wins" - while one a
    // maintainer has worked on stays.
    //
    // The flow tests go through the real create and finalise, with the blob store on a temp
    // folder (as SubmissionFlowTests does) and the board's one file already published, so nothing
    // needs uploading.
    // ###########################################################################################
    public sealed class SubmissionReplacementTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 26, 12, 0, 0, TimeSpan.Zero);

        private const string BoardId = "Manu1/Hardware1/Board1";
        private const string ImagePath = "Manu1/Hardware1/Board1/main.png";

        private readonly string thisRoot =
            Path.Combine(Path.GetTempPath(), "crt-replacement-tests", Guid.NewGuid().ToString("N"));

        public SubmissionReplacementTests()
        {
            string image = Path.Combine(this.BoardFolder, ImagePath.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(image)!);
            File.WriteAllBytes(image, SubmissionReplacementTests.ImageBytes);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(this.thisRoot))
                    Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless; failing a test over cleanup is not.
            }
        }

        private string BoardFolder => Path.Combine(this.thisRoot, "board");

        // A real PNG as far as its opening bytes go - finalise checks bytes against names.
        private static readonly byte[] ImageBytes = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, .. Encoding.UTF8.GetBytes("BOARD")];

        private static SubmissionManifest Manifest(string summary) =>
            new()
            {
                BoardId = BoardId,
                Manufacturer = "Manu1",
                Hardware = "Hardware1",
                Board = "Board1",
                Summary = summary,
                CreatedUtc = Now,
                Files =
                {
                    new SubmissionFile
                    {
                        Path = ImagePath,
                        Sha256 = Convert.ToHexStringLower(SHA256.HashData(SubmissionReplacementTests.ImageBytes)),
                        SizeBytes = SubmissionReplacementTests.ImageBytes.Length
                    }
                },
                Rows = new SubmissionRows
                {
                    Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = ImagePath } },
                    Components = { new ComponentEntry { BoardLabel = "D1" } }
                }
            };

        // Creates a submission and, unless told not to, finalises it - so it is queued. Returns its id.
        private async Task<long> SendAsync(FakeSubmissionStore store, string email, string summary, bool finish = true)
        {
            BlobStore blobs = new(this.thisRoot, NullLogger<BlobStore>.Instance);

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionReplacementTests.Manifest(summary), Submitter.Anonymous(email, "192.0.2.1"), this.BoardFolder,
                store, blobs, Now, CancellationToken.None, PublishedTreeProbe.For(this.BoardFolder));

            Assert.True(created.IsAccepted, string.Join("; ", created.Findings.Select(finding => finding.Code)));

            if (finish)
            {
                SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                    created.Negotiation!.SubmissionId, created.Negotiation.UploadToken, store, blobs, Now);

                Assert.True(result.IsAccepted, string.Join("; ", result.Findings.Select(finding => finding.Code)));
            }

            return created.Negotiation!.SubmissionId;
        }

        // ------------------------------------------------------------------ The flow

        // ###########################################################################################
        // *** THE CASE ASKED FOR. *** The same contributor sends the board again before the first is
        // looked at: the first leaves the queue as withdrawn, with a comment the contributor reads
        // in CRT, and the newer one is the one waiting.
        // ###########################################################################################
        [Fact]
        public async Task A_newer_submission_from_the_same_contributor_replaces_the_older_one()
        {
            var store = new FakeSubmissionStore();

            long first = await this.SendAsync(store, "dennis@example.com", "First go.");
            long second = await this.SendAsync(store, "Dennis@Example.com ", "Second go.");

            Assert.Equal(SubmissionState.Withdrawn, store.Submissions[first].State);
            Assert.Equal(SubmissionReplacementRules.ReplacedComment, store.Submissions[first].DecisionComment);
            Assert.Equal(Now, store.Submissions[first].DecidedUtc);

            Assert.Equal(SubmissionState.Pending, store.Submissions[second].State);
            Assert.Equal([second], (await store.GetQueueAsync(100)).Select(record => record.Id));
        }

        // Three in a row: only the newest is left.
        [Fact]
        public async Task Only_the_newest_of_several_is_left_waiting()
        {
            var store = new FakeSubmissionStore();

            long first = await this.SendAsync(store, "dennis@example.com", "One.");
            long second = await this.SendAsync(store, "dennis@example.com", "Two.");
            long third = await this.SendAsync(store, "dennis@example.com", "Three.");

            Assert.Equal(SubmissionState.Withdrawn, store.Submissions[first].State);
            Assert.Equal(SubmissionState.Withdrawn, store.Submissions[second].State);
            Assert.Equal([third], (await store.GetQueueAsync(100)).Select(record => record.Id));
        }

        // ###########################################################################################
        // *** ONE A MAINTAINER HAS WORKED ON STAYS. *** Their table edits exist only in it - the
        // contributor's draft never had them - and a first approval would be lost too.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_a_maintainer_amended_stays()
        {
            var store = new FakeSubmissionStore();

            long first = await this.SendAsync(store, "dennis@example.com", "First go.");
            store.Amendments.Add((first, 1, SubmissionReplacementTests.Manifest("First go."), "A maintainer", Now));

            long second = await this.SendAsync(store, "dennis@example.com", "Second go.");

            Assert.Equal(SubmissionState.Pending, store.Submissions[first].State);
            Assert.Equal(SubmissionState.Pending, store.Submissions[second].State);
        }

        [Fact]
        public async Task A_submission_with_its_first_of_two_approvals_stays()
        {
            var store = new FakeSubmissionStore();

            long first = await this.SendAsync(store, "dennis@example.com", "First go.");
            store.Submissions[first] = store.Submissions[first] with { State = SubmissionState.Approved };

            await this.SendAsync(store, "dennis@example.com", "Second go.");

            Assert.Equal(SubmissionState.Approved, store.Submissions[first].State);
        }

        // Somebody else's submission of the same board is theirs to be reviewed.
        [Fact]
        public async Task Another_contributors_submission_of_the_same_board_stays()
        {
            var store = new FakeSubmissionStore();

            long theirs = await this.SendAsync(store, "someone@example.com", "Theirs.");
            await this.SendAsync(store, "dennis@example.com", "Mine.");

            Assert.Equal(SubmissionState.Pending, store.Submissions[theirs].State);
        }

        // ###########################################################################################
        // *** ONLY ONCE THE NEWER ONE IS QUEUED. *** A send that never finishes uploading must not
        // have taken the older one's place - the queue would then hold neither.
        // ###########################################################################################
        [Fact]
        public async Task A_newer_send_that_never_finishes_replaces_nothing()
        {
            var store = new FakeSubmissionStore();

            long first = await this.SendAsync(store, "dennis@example.com", "First go.");
            await this.SendAsync(store, "dennis@example.com", "Never finished.", finish: false);

            Assert.Equal(SubmissionState.Pending, store.Submissions[first].State);
        }

        // ###########################################################################################
        // *** BOTH SIDES OF THE WORD (One change, every side of it). *** The server leaves a replaced
        // submission 'withdrawn'; CRT, reading that state back, must say "replaced" - and must not
        // paint it as refused. Asked here, where the server's constant and CRT.Data's presenter
        // meet, so neither side can move alone.
        // ###########################################################################################
        [Fact]
        public void The_contributor_reads_the_replaced_state_as_replaced_not_refused()
        {
            Assert.Equal("Replaced by a newer submission", SubmissionReceiptPresenter.DescribeState(SubmissionState.Withdrawn));
            Assert.NotEqual(SubmissionOutcomeKind.Bad, SubmissionReceiptPresenter.ClassifyState(SubmissionState.Withdrawn));
            Assert.False(SubmissionReceiptPresenter.IsStillOpen(SubmissionState.Withdrawn));
        }

        // ------------------------------------------------------------------ The rules

        private static SubmissionRecord Record(
            long id,
            string? email = "dennis@example.com",
            long? account = null,
            string boardId = BoardId,
            string state = SubmissionState.Pending) =>
            new(id, boardId, account, email, "hash", "", state, null, 1, Now, null, null);

        [Theory]
        [InlineData("dennis@example.com", "dennis@example.com", true)]
        [InlineData("dennis@example.com", " DENNIS@example.com ", true)]
        [InlineData("dennis@example.com", "someone@example.com", false)]
        [InlineData("", "", false)]
        [InlineData(null, null, false)]
        public void Anonymous_contributors_are_the_same_by_their_email(string? first, string? second, bool same)
        {
            Assert.Equal(same, SubmissionReplacementRules.IsSameContributor(Record(1, first), Record(2, second)));
        }

        // A signed-in contributor is their account; a signed-in and an anonymous submission are
        // never the same person here - the signed-in one stores no email to compare.
        [Fact]
        public void Signed_in_contributors_are_the_same_by_their_account()
        {
            Assert.True(SubmissionReplacementRules.IsSameContributor(Record(1, null, account: 7), Record(2, null, account: 7)));
            Assert.False(SubmissionReplacementRules.IsSameContributor(Record(1, null, account: 7), Record(2, null, account: 8)));
            Assert.False(SubmissionReplacementRules.IsSameContributor(Record(1, null, account: 7), Record(2, "dennis@example.com")));
        }

        // Only OLDER, WAITING submissions of the SAME BOARD from the SAME CONTRIBUTOR - each
        // condition is the only thing that differs about one of the candidates here.
        [Fact]
        public void Only_older_waiting_submissions_of_the_same_board_and_contributor_are_replaced()
        {
            SubmissionRecord arrived = Record(10);

            IReadOnlyList<SubmissionRecord> waiting =
            [
                Record(3),
                Record(4, email: "someone@example.com"),
                Record(5, boardId: "Commodore/C64/250407"),
                Record(6, state: SubmissionState.Approved),
                Record(11),
                arrived
            ];

            Assert.Equal([3], SubmissionReplacementRules.ReplacedBy(arrived, waiting).Select(record => record.Id));
        }
    }
}
