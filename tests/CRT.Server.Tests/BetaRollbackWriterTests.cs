using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the disk half of a BETA rollback on its own: BetaRollbackWriter's verified copy, and
    // BetaRollbackFiles' comparisons (code review, 2026-09-27).
    //
    // The flow tests prove the rollback end to end; these pin the two properties the review found
    // missing and that an end-to-end test cannot reach: a restore that is VERIFIED against the
    // bytes the plan compared (never a bare File.Copy onto the served path), and a comparison that
    // does not hash what it does not need to.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class BetaRollbackWriterTests : IDisposable
    {
        private const string Sheet = "Commodore/C64/250407/Images/sheet1.png";

        private readonly string thisRoot;
        private readonly string thisBeta;
        private readonly string thisProduction;

        public BetaRollbackWriterTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-rollback-writer", Guid.NewGuid().ToString("N"));
            this.thisBeta = Path.Combine(this.thisRoot, "beta");
            this.thisProduction = Path.Combine(this.thisRoot, "production");

            Directory.CreateDirectory(this.thisBeta);
            Directory.CreateDirectory(this.thisProduction);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private static SystemRecord System() =>
            new("Commodore/C64/250407", "Commodore", "C64", "250407", "r2", true, "beta", "r1", "production", DateTimeOffset.UtcNow);

        private static string Sha(string content) =>
            Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(content)));

        private string In(string root, string relative)
        {
            string full = Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            return full;
        }

        private static BetaRollbackFilePlan RestoreOf(string path, string expectedHash) =>
            new(
                new BetaRollbackPlanResult(BetaRollbackKind.RestoreFromProduction, [path], [], [], []),
                new Dictionary<string, string> { [path] = expectedHash });

        [Fact]
        public async Task A_restore_writes_productions_bytes_and_leaves_no_temporary_file()
        {
            File.WriteAllText(this.In(this.thisProduction, BetaRollbackWriterTests.Sheet), "PUBLISHED");
            File.WriteAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), "SUBMITTED");

            BetaRollbackWriteOutcome outcome = await BetaRollbackWriter.ApplyAsync(
                BetaRollbackWriterTests.RestoreOf(BetaRollbackWriterTests.Sheet, BetaRollbackWriterTests.Sha("PUBLISHED")),
                this.thisBeta, this.thisProduction, BetaRollbackWriterTests.System(), NullLogger.Instance);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal("PUBLISHED", File.ReadAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet)));

            // The hidden temporary file is renamed over the target, never left beside it.
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet))!));
        }

        // ###########################################################################################
        // *** A BETA FOLDER THE SERVICE MAY NOT WRITE STOPS THE PUSH-BACK BEFORE ANYTHING MOVES
        // (owner report, 2026-09-28). *** A push-back is the writer that restores FROM production,
        // so it is the first to meet a folder copied from there into BETA by hand. Restores and
        // removals alike are asked about first; the refusal names the folder and nothing changes.
        // (The probe is told the answer: such a folder cannot be made on every OS.)
        // ###########################################################################################
        [Fact]
        public async Task A_beta_folder_the_service_may_not_write_changes_nothing()
        {
            File.WriteAllText(this.In(this.thisProduction, BetaRollbackWriterTests.Sheet), "PUBLISHED");
            string betaSheet = this.In(this.thisBeta, BetaRollbackWriterTests.Sheet);
            File.WriteAllText(betaSheet, "SUBMITTED");

            string images = Path.GetDirectoryName(betaSheet)!;

            BetaRollbackWriteOutcome outcome = await BetaRollbackWriter.ApplyAsync(
                BetaRollbackWriterTests.RestoreOf(BetaRollbackWriterTests.Sheet, BetaRollbackWriterTests.Sha("PUBLISHED")),
                this.thisBeta, this.thisProduction, BetaRollbackWriterTests.System(), NullLogger.Instance,
                canWriteFolder: folder => folder != images);

            Assert.False(outcome.IsDone);
            Assert.Contains("[Commodore/C64/250407/Images] in the BETA data, so nothing was changed", outcome.Error, StringComparison.Ordinal);
            Assert.Equal("SUBMITTED", File.ReadAllText(betaSheet));
        }

        // A plan for a promoted system whose production folder cannot be read does nothing, even if
        // a caller hands it over without checking (the flow refuses it first).
        [Fact]
        public async Task A_plan_whose_production_cannot_be_read_changes_nothing()
        {
            File.WriteAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), "THE BOARD");

            BetaRollbackWriteOutcome outcome = await BetaRollbackWriter.ApplyAsync(
                new BetaRollbackFilePlan(
                    new BetaRollbackPlanResult(BetaRollbackKind.ProductionUnreadable, [], [BetaRollbackWriterTests.Sheet], [], []),
                    new Dictionary<string, string>()),
                this.thisBeta, this.thisProduction, BetaRollbackWriterTests.System(), NullLogger.Instance);

            Assert.False(outcome.IsDone);
            Assert.True(File.Exists(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet)));
        }

        // ###########################################################################################
        // *** PRODUCTION CHANGED AFTER IT WAS CHECKED - NOTHING IS WRITTEN (code review,
        // 2026-09-27). *** The first version's bare File.Copy copied whatever production held at
        // that moment, unseen. VerifiedFileCopy hashes what it copies and replaces the target only
        // on a match, so BETA keeps its bytes and the maintainer is told why. Fails against a
        // File.Copy writer, which would have overwritten the BETA file here.
        // ###########################################################################################
        [Fact]
        public async Task A_production_file_that_changed_since_the_plan_is_refused_and_beta_is_untouched()
        {
            File.WriteAllText(this.In(this.thisProduction, BetaRollbackWriterTests.Sheet), "EDITED BY HAND SINCE");
            File.WriteAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), "SUBMITTED");

            BetaRollbackWriteOutcome outcome = await BetaRollbackWriter.ApplyAsync(
                BetaRollbackWriterTests.RestoreOf(BetaRollbackWriterTests.Sheet, BetaRollbackWriterTests.Sha("WHAT THE PLAN SAW")),
                this.thisBeta, this.thisProduction, BetaRollbackWriterTests.System(), NullLogger.Instance);

            Assert.False(outcome.IsDone);
            Assert.Contains("changed in the stable source after it was checked", outcome.Error, StringComparison.Ordinal);
            Assert.Equal("SUBMITTED", File.ReadAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet)));
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet))!));
        }

        // -----------------------------------------------------------------------------------
        // The comparison (code review, 2026-09-27)
        // -----------------------------------------------------------------------------------

        // Different lengths are different files - decided without reading either.
        [Fact]
        public async Task Files_of_different_length_are_different()
        {
            File.WriteAllText(this.In(this.thisProduction, BetaRollbackWriterTests.Sheet), "SHORT");
            File.WriteAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), "MUCH LONGER");

            Assert.False(await BetaRollbackFiles.SameBytesAsync(this.thisBeta, this.thisProduction, BetaRollbackWriterTests.Sheet));
        }

        // Same length, different bytes: hashed, and still different.
        [Fact]
        public async Task Files_of_equal_length_are_compared_by_their_bytes()
        {
            File.WriteAllText(this.In(this.thisProduction, BetaRollbackWriterTests.Sheet), "AAAA");
            File.WriteAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), "BBBB");

            Assert.False(await BetaRollbackFiles.SameBytesAsync(this.thisBeta, this.thisProduction, BetaRollbackWriterTests.Sheet));

            File.WriteAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), "AAAA");
            File.SetLastWriteTimeUtc(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), DateTime.UtcNow.AddMinutes(1));

            Assert.True(await BetaRollbackFiles.SameBytesAsync(this.thisBeta, this.thisProduction, BetaRollbackWriterTests.Sheet));
        }

        // ###########################################################################################
        // *** THE CACHE IS PER FILE VERSION. *** Hashes are kept between the confirmation's pass and
        // the one under the lock, keyed by path, length and last-write time - so a file written in
        // between is read again rather than answered from the cache. Same length on purpose: only
        // the time tells the two versions apart.
        // ###########################################################################################
        [Fact]
        public async Task A_file_written_again_is_hashed_again()
        {
            string path = this.In(this.thisBeta, BetaRollbackWriterTests.Sheet);

            File.WriteAllText(path, "VERSION1");
            File.SetLastWriteTimeUtc(path, new DateTime(2026, 9, 1, 0, 0, 0, DateTimeKind.Utc));
            Assert.Equal(BetaRollbackWriterTests.Sha("VERSION1"), await BetaRollbackFiles.HashOfAsync(path));

            File.WriteAllText(path, "VERSION2");
            File.SetLastWriteTimeUtc(path, new DateTime(2026, 9, 2, 0, 0, 0, DateTimeKind.Utc));
            Assert.Equal(BetaRollbackWriterTests.Sha("VERSION2"), await BetaRollbackFiles.HashOfAsync(path));
        }

        // ###########################################################################################
        // *** A FILE THE ROLLBACK WROTE IS NEVER ANSWERED FROM THE CACHE. *** A version is length +
        // last-write time, and file times move in ticks, so a restore landing in the same tick as
        // the hash the plan read looked unchanged. The full suite caught it intermittently - a
        // second push-back re-restored a file already put back - and the same staleness the other
        // way would SKIP a restore that is needed. Deterministic here: the restored file is given
        // back the very time it had when it was hashed. Fails without BetaRollbackFiles.Forget.
        // ###########################################################################################
        [Fact]
        public async Task A_file_the_rollback_restored_is_not_answered_from_a_stale_hash()
        {
            string beta = this.In(this.thisBeta, BetaRollbackWriterTests.Sheet);
            string production = this.In(this.thisProduction, BetaRollbackWriterTests.Sheet);
            DateTime tick = new(2026, 9, 1, 0, 0, 0, DateTimeKind.Utc);

            File.WriteAllText(production, "PUBLISHED");
            File.WriteAllText(beta, "SUBMITTED");
            File.SetLastWriteTimeUtc(beta, tick);

            // The plan's pass: equal length, so both are hashed and BETA's is cached.
            Assert.False(await BetaRollbackFiles.SameBytesAsync(this.thisBeta, this.thisProduction, BetaRollbackWriterTests.Sheet));

            BetaRollbackWriteOutcome outcome = await BetaRollbackWriter.ApplyAsync(
                BetaRollbackWriterTests.RestoreOf(BetaRollbackWriterTests.Sheet, BetaRollbackWriterTests.Sha("PUBLISHED")),
                this.thisBeta, this.thisProduction, BetaRollbackWriterTests.System(), NullLogger.Instance);

            Assert.True(outcome.IsDone, outcome.Error);

            // Same length, and now the same time as when it was hashed - the same-tick case.
            File.SetLastWriteTimeUtc(beta, tick);

            Assert.True(await BetaRollbackFiles.SameBytesAsync(this.thisBeta, this.thisProduction, BetaRollbackWriterTests.Sheet));
        }

        // The hash is lower-case hex, the form VerifiedFileCopy compares - an upper-case one would
        // refuse every restore as "changed since it was checked".
        [Fact]
        public async Task A_hash_is_in_the_form_the_verified_copy_compares()
        {
            string path = this.In(this.thisBeta, BetaRollbackWriterTests.Sheet);
            File.WriteAllText(path, "ANYTHING");

            string? hash = await BetaRollbackFiles.HashOfAsync(path);

            Assert.Equal(hash, hash!.ToLowerInvariant());
        }

        // A dot-file beside the board is a copy in progress (VerifiedFileCopy's temporary name), not
        // part of it - never restored over, never removed.
        [Fact]
        public void A_hidden_temporary_file_is_not_part_of_the_board()
        {
            File.WriteAllText(this.In(this.thisBeta, "Commodore/C64/250407/Images/.sheet1.png.crt-publish-1.tmp"), "HALF");
            File.WriteAllText(this.In(this.thisBeta, BetaRollbackWriterTests.Sheet), "WHOLE");

            Assert.Equal([BetaRollbackWriterTests.Sheet], BetaRollbackFiles.FilesUnder(this.thisBeta, "Commodore/C64/250407"));
        }
    }
}
