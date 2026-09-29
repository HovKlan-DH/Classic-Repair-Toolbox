using System;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers what a contributor sees while a submission is in flight, and when a failed step is
    // worth retrying.
    //
    // Both are small pieces of arithmetic whose failure mode is a dialog that LOOKS broken: a bar
    // stuck at zero, a percentage that goes backwards, or a retry loop that hammers a server which
    // has already said no. None of those throw, so none of them would be caught by anything but a
    // test like these.
    // ###########################################################################################
    public class SubmissionProgressTests
    {
        // -----------------------------------------------------------------------------------
        // The fraction.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Progress_with_nothing_to_measure_reports_NULL_not_zero()
        {
            // A bar showing 0% for a phase with no measurable total looks stuck. Null is how the
            // UI knows to show an indeterminate state instead.
            var progress = new SubmissionProgress(SubmissionPhase.Preparing, 0, 0, 0, 0, null);

            Assert.Null(progress.Fraction);
        }

        [Fact]
        public void Bytes_are_preferred_over_file_counts_for_the_fraction()
        {
            // Twelve small files and one huge one would otherwise show 92% while most of the
            // actual transfer is still to come.
            var progress = new SubmissionProgress(SubmissionPhase.Uploading, 11, 12, 100, 1000, "big.png");

            Assert.Equal(0.1, progress.Fraction!.Value, 3);
        }

        [Fact]
        public void File_counts_are_used_when_there_are_no_bytes_to_count()
        {
            // The hashing phase reads files but sends nothing.
            var progress = new SubmissionProgress(SubmissionPhase.Hashing, 3, 12, 0, 0, null);

            Assert.Equal(0.25, progress.Fraction!.Value, 3);
        }

        // ###########################################################################################
        // *** THE BYTE COUNT CARRIES OVER FROM ONE FILE TO THE NEXT. *** The upload loop carried
        // the file count forward and never the bytes, so every file started again from "0 bytes" -
        // reported as "Uploading 470 of 1212 - U25_5_NTSC.png (0 bytes of 121.0 MB)" - and since
        // the fraction prefers bytes, the bar measured one small file against the whole total and
        // sat near empty for the entire upload.
        // ###########################################################################################
        [Fact]
        public void The_byte_count_carries_over_from_one_file_to_the_next()
        {
            var progress = new SubmissionProgress(SubmissionPhase.Uploading, 0, 3, 0, 300, "a.png");

            progress = progress.AfterFileSent(100).AfterFileSent(100);

            Assert.Equal(2, progress.FilesDone);
            Assert.Equal(200, progress.BytesDone);
            Assert.Equal(2.0 / 3.0, progress.Fraction!.Value, 3);
            Assert.Contains("(200 bytes of 300 bytes)", progress.Describe());
        }

        [Fact]
        public void The_fraction_never_exceeds_one_or_goes_negative()
        {
            // A miscounted total should not make a progress bar overflow or invert.
            var over = new SubmissionProgress(SubmissionPhase.Uploading, 0, 0, 2000, 1000, null);
            var under = new SubmissionProgress(SubmissionPhase.Uploading, 0, 0, -50, 1000, null);

            Assert.Equal(1.0, over.Fraction!.Value);
            Assert.Equal(0.0, under.Fraction!.Value);
        }

        // -----------------------------------------------------------------------------------
        // The description. Written for a person watching a slow upload.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_hashing_phase_explains_the_wait()
        {
            // Hashing a 76 MB system takes seconds during which nothing is sent. A bar sitting at
            // zero with no explanation looks broken.
            var progress = new SubmissionProgress(SubmissionPhase.Hashing, 3, 12, 0, 0, null);

            Assert.Contains("Checking which files need sending", progress.Describe(), StringComparison.Ordinal);
            Assert.Contains("3 of 12", progress.Describe(), StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** THE HEADING SAYS "SENDING" ONLY ONCE SOMETHING IS SENT (owner report, 2026-09-28). ***
        // It read "Sending contribution" through the whole local file check of a 2,242-file system,
        // while nothing had left the machine. Every phase before the upload is "Preparing".
        // ###########################################################################################
        [Theory]
        [InlineData(SubmissionPhase.Preparing, "Preparing contribution")]
        [InlineData(SubmissionPhase.Hashing, "Preparing contribution")]
        [InlineData(SubmissionPhase.Negotiating, "Preparing contribution")]
        [InlineData(SubmissionPhase.Uploading, "Sending contribution")]
        [InlineData(SubmissionPhase.Finalising, "Sending contribution")]
        [InlineData(SubmissionPhase.Done, "Contribution sent")]
        public void The_heading_says_sending_only_once_the_upload_has_started(SubmissionPhase phase, string expected)
        {
            var progress = new SubmissionProgress(phase, 1, 2, 10, 100, "file.png");

            Assert.Equal(expected, progress.Heading);
        }

        [Fact]
        public void A_submission_that_has_just_started_is_preparing_not_sending()
        {
            // What the dialog shows the moment Submit is pressed, before any progress is reported.
            Assert.Equal("Preparing contribution", SubmissionProgress.Starting().Heading);
        }

        [Fact]
        public void The_uploading_phase_names_the_file_and_the_sizes()
        {
            // "40%" tells a contributor nothing about whether it is stuck. The file name and the
            // byte counts do.
            var progress = new SubmissionProgress(
                SubmissionPhase.Uploading, 2, 12, 1536, 1048576, "U8-pin3.png");

            string text = progress.Describe();

            Assert.Contains("3 of 12", text, StringComparison.Ordinal);
            Assert.Contains("U8-pin3.png", text, StringComparison.Ordinal);
            Assert.Contains("1.5 KB", text, StringComparison.Ordinal);
            Assert.Contains("1.0 MB", text, StringComparison.Ordinal);
        }

        [Fact]
        public void An_upload_with_nothing_to_send_says_so_plainly()
        {
            // The typo-fix case. "Uploading 1 of 0" would be nonsense.
            var progress = new SubmissionProgress(SubmissionPhase.Uploading, 0, 0, 0, 0, null);

            Assert.Contains("already has every file", progress.Describe(), StringComparison.Ordinal);
        }

        [Fact]
        public void Every_phase_describes_itself()
        {
            // A phase added later without a description would show "Working..." - harmless but
            // uninformative, and easy to miss. This catches it.
            foreach (SubmissionPhase phase in Enum.GetValues<SubmissionPhase>())
            {
                var progress = new SubmissionProgress(phase, 1, 2, 10, 100, "file.png");

                Assert.False(string.IsNullOrWhiteSpace(progress.Describe()));
                Assert.NotEqual("Working...", progress.Describe());
            }
        }

        // -----------------------------------------------------------------------------------
        // Retry.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_transport_failure_with_no_response_is_retried()
        {
            // A dropped connection or a DNS blip - the ordinary transient case, and the one most
            // worth retrying.
            Assert.True(SubmissionRetryPolicy.ShouldRetry(null, 0));
        }

        [Theory]
        [InlineData(408)]
        [InlineData(429)]
        [InlineData(500)]
        [InlineData(502)]
        [InlineData(503)]
        [InlineData(504)]
        public void A_server_side_or_overload_failure_is_retried(int statusCode)
        {
            Assert.True(SubmissionRetryPolicy.ShouldRetry(statusCode, 0));
        }

        [Theory]
        [InlineData(400)]
        [InlineData(401)]
        [InlineData(403)]
        [InlineData(404)]
        [InlineData(409)]
        public void A_request_the_server_REFUSED_is_not_retried(int statusCode)
        {
            // It was wrong and will be wrong again. Retrying wastes the contributor's time and
            // hammers the server for nothing.
            Assert.False(SubmissionRetryPolicy.ShouldRetry(statusCode, 0));
        }

        [Fact]
        public void A_416_is_not_retried_because_it_is_a_RESUME_not_a_failure()
        {
            // The offset disagreed; the caller re-asks where to resume and continues. Treating it
            // as a retry would send the same wrong offset again.
            Assert.False(SubmissionRetryPolicy.ShouldRetry(416, 0));
        }

        [Fact]
        public void Retrying_stops_after_the_maximum_number_of_attempts()
        {
            // Otherwise a genuinely unreachable server leaves a dialog retrying forever.
            Assert.False(SubmissionRetryPolicy.ShouldRetry(null, SubmissionRetryPolicy.MaximumAttempts));
        }

        [Fact]
        public void The_delay_grows_but_is_capped()
        {
            // A fixed short delay against a briefly overloaded server is a small denial of
            // service; an unbounded exponential leaves someone staring at a frozen dialog.
            Assert.Equal(TimeSpan.Zero, SubmissionRetryPolicy.DelayBefore(0));
            Assert.Equal(TimeSpan.FromSeconds(2), SubmissionRetryPolicy.DelayBefore(1));
            Assert.Equal(TimeSpan.FromSeconds(4), SubmissionRetryPolicy.DelayBefore(2));
            Assert.Equal(TimeSpan.FromSeconds(8), SubmissionRetryPolicy.DelayBefore(3));

            Assert.True(SubmissionRetryPolicy.DelayBefore(20) <= SubmissionRetryPolicy.MaximumDelay);
        }
    }
}
