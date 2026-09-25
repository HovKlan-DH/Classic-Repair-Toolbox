using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers SubmissionFileScopes.Classify - WHOSE file a submitted path is (security review,
    // 2026-09-25). A submission may change its own board's folder and the two shared folders;
    // anything else belongs to another board.
    // ###########################################################################################
    public sealed class SubmissionFileScopeTests
    {
        private static SubmissionFileScope Classify(string path) =>
            SubmissionFileScopes.Classify("Commodore", "C64", "250407", path);

        [Theory]
        [InlineData("Commodore/C64/250407/Sheet1.png", SubmissionFileScope.Own)]
        [InlineData("Commodore/C64/250407/Scope baseline/U8 pin 3.png", SubmissionFileScope.Own)]
        [InlineData("Commodore/Shared files/Component images/6526.png", SubmissionFileScope.ManufacturerShared)]
        [InlineData("Generic shared files/Datasheets/7805.pdf", SubmissionFileScope.GenericShared)]
        [InlineData("Commodore/C64/250425/Sheet1.png", SubmissionFileScope.Foreign)]
        [InlineData("Amstrad/CPC 664/MC0005A/board.png", SubmissionFileScope.Foreign)]
        [InlineData("Amstrad/Shared files/x.png", SubmissionFileScope.Foreign)]
        public void Each_path_is_classified_by_the_board_the_submission_names(string path, SubmissionFileScope expected)
        {
            Assert.Equal(expected, SubmissionFileScopeTests.Classify(path));
        }

        // CASE-SENSITIVE, like every path comparison against the Linux tree: a case-variant of this
        // board's folder is a DIFFERENT folder there, and must not pass as this board's own.
        [Theory]
        [InlineData("commodore/C64/250407/Sheet1.png")]
        [InlineData("Commodore/c64/250407/Sheet1.png")]
        [InlineData("Commodore/shared files/x.png")]
        [InlineData("generic shared files/x.png")]
        public void A_case_variant_of_an_allowed_folder_is_foreign(string path)
        {
            Assert.Equal(SubmissionFileScope.Foreign, SubmissionFileScopeTests.Classify(path));
        }

        // A board whose name merely STARTS the same is another board - the trailing slash is what
        // makes "250407" not match "2504070".
        [Fact]
        public void A_board_whose_name_starts_the_same_is_another_board()
        {
            Assert.Equal(SubmissionFileScope.Foreign, SubmissionFileScopeTests.Classify("Commodore/C64/2504070/a.png"));
        }

        // An unidentifiable submission owns nothing - a blank part must not turn "//" into a prefix
        // that matches.
        [Fact]
        public void A_blank_identity_owns_nothing()
        {
            Assert.Equal(SubmissionFileScope.Foreign, SubmissionFileScopes.Classify("", "", "", "a/b/c/x.png"));
            Assert.Equal(SubmissionFileScope.Foreign, SubmissionFileScopes.Classify("Commodore", "C64", "250407", ""));
        }
    }
}
