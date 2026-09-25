using System.Text.Json;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The contributor's status answer, end to end (code review, 2026-09-25).
    //
    // GET /api/submissions/{id} is what CRT's "My submissions" reads - the state, the maintainer's
    // comment, whether a maintainer changed rows. It used to be an anonymous object with hand-typed
    // names, and the reviewer-to-maintainer rename changed those names at both ends with no test
    // that would have noticed only one moving: every contributor's feedback box would simply have
    // come back empty. The endpoint now answers with CRT.Data's SubmissionStatus, the type CRT reads
    // it into, built by SubmissionEndpoints.BuildStatus.
    //
    // This puts BuildStatus's answer through what each end really uses: the server's JSON settings
    // (ReviewApiContract.ApplyWireSettings on ASP.NET's web defaults, as Program.cs configures) and
    // CRT's reader - SubmissionClient.GetStatusAsync's ReadFromJsonAsync, which reads with
    // JsonSerializerDefaults.Web.
    // ###########################################################################################
    public sealed class SubmissionStatusWireTests
    {
        private static readonly DateTimeOffset Created = new(2026, 9, 25, 10, 0, 0, TimeSpan.Zero);
        private static readonly DateTimeOffset Decided = new(2026, 9, 25, 14, 30, 0, TimeSpan.Zero);

        private static SubmissionStatus RoundTrip(SubmissionStatus sent)
        {
            var server = new JsonSerializerOptions(JsonSerializerDefaults.Web);
            ReviewApiContract.ApplyWireSettings(server);

            string json = JsonSerializer.Serialize(sent, server);

            return JsonSerializer.Deserialize<SubmissionStatus>(json, new JsonSerializerOptions(JsonSerializerDefaults.Web))!;
        }

        private static SubmissionRecord Record(string? comment) =>
            new(
                Id: 42,
                SystemId: "Commodore/C64/250407",
                AccountId: null,
                ContactEmail: "someone@example.com",
                UploadTokenHash: "hash",
                BaseRevision: "2026-August-21",
                State: SubmissionState.Rejected,
                Summary: "Corrected U8.",
                FormatVersion: 1,
                CreatedUtc: SubmissionStatusWireTests.Created,
                ExpiresUtc: null,
                DecidedUtc: SubmissionStatusWireTests.Decided,
                DecisionComment: comment);

        [Fact]
        public void A_decided_submission_reaches_CRT_with_the_maintainers_comment_and_the_amendment()
        {
            var finding = new ValidationFinding
            {
                Severity = ValidationSeverity.Warning,
                Code = "image.orphan",
                Subject = "U9",
                Message = "An image references component [U9]."
            };

            SubmissionStatus read = SubmissionStatusWireTests.RoundTrip(SubmissionEndpoints.BuildStatus(
                SubmissionStatusWireTests.Record("Wrong pin on U8."),
                SubmissionState.Rejected,
                amendedByMaintainer: true,
                [finding]));

            Assert.Equal(42, read.Id);
            Assert.Equal("Commodore/C64/250407", read.SystemId);
            Assert.Equal(SubmissionState.Rejected, read.State);
            Assert.Equal("Corrected U8.", read.Summary);
            Assert.Equal(SubmissionStatusWireTests.Created, read.CreatedUtc);
            Assert.Equal(SubmissionStatusWireTests.Decided, read.DecidedUtc);
            Assert.Equal("Wrong pin on U8.", read.MaintainerComment);
            Assert.True(read.AmendedByMaintainer);
            Assert.Equal("image.orphan", Assert.Single(read.Findings).Code);
        }

        // No comment is an empty string, never missing: CRT compares the text to decide whether
        // there is something unread, and "nothing said" must read the same every time.
        [Fact]
        public void A_submission_with_no_comment_reaches_CRT_with_an_empty_one()
        {
            SubmissionStatus read = SubmissionStatusWireTests.RoundTrip(SubmissionEndpoints.BuildStatus(
                SubmissionStatusWireTests.Record(null),
                SubmissionState.Pending,
                amendedByMaintainer: false,
                []));

            Assert.Equal(string.Empty, read.MaintainerComment);
            Assert.False(read.AmendedByMaintainer);
            Assert.Empty(read.Findings);
        }
    }
}
