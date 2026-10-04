using CRT.Server.Handlers.Submissions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SystemDeletionRules - which submissions of a deleted system count as still in play
    // (their contributors are mailed), the fingerprint that holds a delete to what was confirmed,
    // and the refusal sentences that count rather than list.
    // ###########################################################################################
    public sealed class SystemDeletionRulesTests
    {
        // ###########################################################################################
        // Open: waiting for review (also again after a push-back), approved once, in BETA. NOT open:
        // finished one way or another - published to the stable source, rejected, sent back for
        // changes, replaced, or never arrived.
        // ###########################################################################################
        [Theory]
        [InlineData(SubmissionState.Pending, true)]
        [InlineData(ProductionPromotionRules.ReturnedState, true)]
        [InlineData(SubmissionState.Approved, true)]
        [InlineData(SubmissionState.Merged, true)]
        [InlineData(ProductionPromotionRules.PublishedState, false)]
        [InlineData(SubmissionState.Rejected, false)]
        [InlineData(SubmissionState.ChangesRequested, false)]
        [InlineData(SubmissionState.Withdrawn, false)]
        [InlineData(SubmissionState.Uploading, false)]
        [InlineData(SubmissionState.Abandoned, false)]
        [InlineData(null, false)]
        public void Only_a_submission_still_in_play_is_open(string? contributorFacingState, bool open)
        {
            Assert.Equal(open, SystemDeletionRules.IsOpen(contributorFacingState));
        }

        private static string Fingerprint(
            string[]? beta = null,
            string[]? production = null,
            bool listedInBeta = true,
            bool listedInProduction = true,
            bool hasRecord = true,
            (long, string)[]? submissions = null) =>
            SystemDeletionRules.Fingerprint(
                beta ?? ["Commodore/C64/999999/a.xlsx", "Commodore/C64/999999/b.png"],
                production ?? ["Commodore/C64/999999/a.xlsx"],
                listedInBeta,
                listedInProduction,
                hasRecord,
                submissions ?? [(4, "pending"), (2, "rejected")]);

        // The order things are found in does not matter - only what they are.
        [Fact]
        public void The_fingerprint_does_not_depend_on_the_order_things_are_listed_in()
        {
            Assert.Equal(
                SystemDeletionRulesTests.Fingerprint(),
                SystemDeletionRulesTests.Fingerprint(
                    beta: ["Commodore/C64/999999/b.png", "Commodore/C64/999999/a.xlsx"],
                    submissions: [(2, "rejected"), (4, "pending")]));
        }

        // ###########################################################################################
        // *** ANY CHANGE TO WHAT WOULD GO CHANGES IT *** - a file more, a file in the other tree, a row
        // added or taken out, the record, a submission more or one whose state moved. Each is
        // something the administrator would be deleting without having seen it.
        // ###########################################################################################
        [Fact]
        public void The_fingerprint_changes_with_everything_the_delete_would_remove()
        {
            string shown = SystemDeletionRulesTests.Fingerprint();

            string[] changed =
            [
                SystemDeletionRulesTests.Fingerprint(beta: ["Commodore/C64/999999/a.xlsx", "Commodore/C64/999999/b.png", "Commodore/C64/999999/c.png"]),
                SystemDeletionRulesTests.Fingerprint(production: ["Commodore/C64/999999/a.xlsx", "Commodore/C64/999999/b.png"]),
                SystemDeletionRulesTests.Fingerprint(listedInBeta: false),
                SystemDeletionRulesTests.Fingerprint(listedInProduction: false),
                SystemDeletionRulesTests.Fingerprint(hasRecord: false),
                SystemDeletionRulesTests.Fingerprint(submissions: [(4, "pending"), (2, "rejected"), (9, "pending")]),
                SystemDeletionRulesTests.Fingerprint(submissions: [(4, "merged"), (2, "rejected")]),

                // The same path in the OTHER tree is a different deletion.
                SystemDeletionRulesTests.Fingerprint(beta: ["Commodore/C64/999999/a.xlsx"], production: ["Commodore/C64/999999/a.xlsx", "Commodore/C64/999999/b.png"]),

                // The server's trees are case-sensitive: two spellings are two files.
                SystemDeletionRulesTests.Fingerprint(beta: ["Commodore/C64/999999/A.xlsx", "Commodore/C64/999999/b.png"]),
            ];

            Assert.All(changed, fingerprint => Assert.NotEqual(shown, fingerprint));
            Assert.Equal(changed.Length, changed.Distinct(StringComparer.Ordinal).Count());
        }

        // One file is named alone; many are counted after the first five.
        [Fact]
        public void Files_another_board_uses_are_named_and_the_rest_counted()
        {
            string one = SystemDeletionRules.UsedElsewhereMessage("BETA", ["Commodore/C64/999999/a.png"]);

            Assert.Contains("uses a file in this system's folder", one, StringComparison.Ordinal);
            Assert.Contains("[Commodore/C64/999999/a.png]", one, StringComparison.Ordinal);

            string many = SystemDeletionRules.UsedElsewhereMessage(
                "stable", [.. Enumerable.Range(1, 7).Select(number => $"Commodore/C64/999999/{number}.png")]);

            Assert.Contains("uses 7 files in this system's folder", many, StringComparison.Ordinal);
            Assert.Contains("[Commodore/C64/999999/5.png] and 2 more", many, StringComparison.Ordinal);
            Assert.DoesNotContain("6.png", many, StringComparison.Ordinal);
            Assert.Contains("in the stable data", many, StringComparison.Ordinal);
        }
    }
}
