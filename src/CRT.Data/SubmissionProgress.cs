using System;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // How far a submission has got, and how that reads on screen. Pure - no timers, no UI, so the
    // arithmetic is a unit test.
    //
    // WHY THIS IS NOT JUST A PERCENTAGE. An upload that reports one number tells a contributor
    // nothing about whether it is stuck: 40% could be halfway through a 2 MB file or halfway
    // through a 300 MB one. Reporting the phase, the current file and the bytes separately is what
    // lets the UI say "uploading 3 of 12 - U8-pin3.png" rather than a bar that may or may not be
    // moving.
    //
    // THE PHASES MATTER TOO. Hashing a 76 MB board takes seconds during which nothing is sent,
    // and a progress bar that sits at zero looks broken. Naming the phase makes the wait
    // explicable.
    // ###########################################################################################
    public sealed record SubmissionProgress(
        SubmissionPhase Phase,
        int FilesDone,
        int FilesTotal,
        long BytesDone,
        long BytesTotal,
        string? CurrentFile)
    {
        public static SubmissionProgress Starting() =>
            new(SubmissionPhase.Preparing, 0, 0, 0, 0, null);

        // ###########################################################################################
        // A fraction from 0 to 1, or null when it cannot honestly be known.
        //
        // NULL RATHER THAN ZERO when there is nothing to measure against. A bar showing 0% for a
        // phase with no measurable total is a bar that appears stuck; the UI should show an
        // indeterminate state instead, and the only way it can know to is if this says so.
        // ###########################################################################################
        public double? Fraction
        {
            get
            {
                if (this.BytesTotal > 0)
                    return Math.Clamp((double)this.BytesDone / this.BytesTotal, 0, 1);

                if (this.FilesTotal > 0)
                    return Math.Clamp((double)this.FilesDone / this.FilesTotal, 0, 1);

                return null;
            }
        }

        // ###########################################################################################
        // The progress once one more file has been sent in full: the file counted AND its bytes.
        //
        // One step rather than two `with` fields at the call site, because the upload loop used to
        // carry the file count forward and not the bytes - every file then began again from
        // "0 bytes", and the bar measured one small file against the whole upload.
        // ###########################################################################################
        public SubmissionProgress AfterFileSent(long sizeBytes) =>
            this with
            {
                FilesDone = this.FilesDone + 1,
                BytesDone = this.BytesDone + Math.Max(0, sizeBytes)
            };

        // ###########################################################################################
        // One line for the UI. Written for a person watching a slow upload, so it says what is
        // happening rather than reporting internal state.
        // ###########################################################################################
        public string Describe()
        {
            return this.Phase switch
            {
                SubmissionPhase.Preparing =>
                    "Working out what to send...",

                // The files are read on THIS machine to fingerprint them; the server is asked which
                // fingerprints it lacks only in the next phase. So "which files need sending" is
                // what the two phases work out between them - nothing is sent or validated yet.
                SubmissionPhase.Hashing when this.FilesTotal > 0 =>
                    string.Create(CultureInfo.InvariantCulture,
                        $"Checking which files need sending ({this.FilesDone} of {this.FilesTotal})..."),

                SubmissionPhase.Hashing =>
                    "Checking which files need sending...",

                SubmissionPhase.Negotiating =>
                    "Asking the server what it already has...",

                SubmissionPhase.Uploading when this.FilesTotal == 0 =>
                    "Nothing to upload - the server already has every file.",

                SubmissionPhase.Uploading =>
                    string.Create(CultureInfo.InvariantCulture,
                        $"Uploading {this.FilesDone + 1} of {this.FilesTotal}" +
                        $"{(this.CurrentFile is null ? string.Empty : $" - {this.CurrentFile}")} " +
                        $"({SubmissionManifestBuilder.FormatSize(this.BytesDone)} of " +
                        $"{SubmissionManifestBuilder.FormatSize(this.BytesTotal)})"),

                SubmissionPhase.Finalising =>
                    "Finishing up...",

                SubmissionPhase.Done =>
                    "Submitted.",

                _ => "Working..."
            };
        }

        // ###########################################################################################
        // The dialog's heading for this phase (owner report, 2026-09-28).
        //
        // *** "SENDING" ONLY ONCE SOMETHING IS BEING SENT. *** The heading read "Sending
        // contribution" from the first moment, while the dialog spent most of a large board's wait
        // reading files on this machine with nothing leaving it - so the heading and the line under
        // it described two different things. Before the upload it is "Preparing".
        // ###########################################################################################
        public string Heading => this.Phase switch
        {
            SubmissionPhase.Preparing or SubmissionPhase.Hashing or SubmissionPhase.Negotiating =>
                "Preparing contribution",

            SubmissionPhase.Done =>
                "Contribution sent",

            _ => "Sending contribution"
        };
    }

    public enum SubmissionPhase
    {
        Preparing,
        Hashing,
        Negotiating,
        Uploading,
        Finalising,
        Done
    }

    // ###########################################################################################
    // Decides whether a failed upload step is worth retrying, and how long to wait.
    //
    // WHY RETRY AT ALL. An upload that dies at 90% and cannot resume will not be attempted again -
    // the strategy document says so outright. Resumption handles a dropped connection the
    // contributor notices; this handles the transient failures they should never have to.
    //
    // WHY NOT RETRY EVERYTHING. A 400 means the request was wrong and will be wrong again;
    // retrying it wastes the contributor's time and hammers the server. Only failures that could
    // plausibly succeed on a second attempt are retried, which is network-shaped ones.
    //
    // EXPONENTIAL WITH A CEILING. A fixed short delay against a server that is briefly overloaded
    // is a small denial of service; an unbounded exponential leaves someone staring at a frozen
    // dialog. Doubling to a one-minute ceiling covers a restart without ever looking hung.
    // ###########################################################################################
    public static class SubmissionRetryPolicy
    {
        public const int MaximumAttempts = 5;

        public static readonly TimeSpan FirstDelay = TimeSpan.FromSeconds(2);

        public static readonly TimeSpan MaximumDelay = TimeSpan.FromMinutes(1);

        // ###########################################################################################
        // Should this attempt be retried?
        //
        // statusCode is the HTTP status, or null when the request failed before getting one (a DNS
        // failure, a dropped connection) - which is exactly the case most worth retrying.
        // ###########################################################################################
        public static bool ShouldRetry(int? statusCode, int attemptsSoFar)
        {
            if (attemptsSoFar >= SubmissionRetryPolicy.MaximumAttempts)
                return false;

            // No response at all: a transport failure, which is the ordinary transient case.
            if (statusCode is null)
                return true;

            return statusCode switch
            {
                // The server is briefly unavailable or overloaded.
                408 or 429 or 500 or 502 or 503 or 504 => true,

                // 416 means the offset disagreed - the caller re-asks for the offset and continues,
                // which is a resume rather than a retry, so it is handled there and not here.
                _ => false
            };
        }

        // ###########################################################################################
        // How long to wait before attempt number `attemptsSoFar + 1`.
        // ###########################################################################################
        public static TimeSpan DelayBefore(int attemptsSoFar)
        {
            if (attemptsSoFar <= 0)
                return TimeSpan.Zero;

            double seconds = SubmissionRetryPolicy.FirstDelay.TotalSeconds * Math.Pow(2, attemptsSoFar - 1);

            return seconds >= SubmissionRetryPolicy.MaximumDelay.TotalSeconds
                ? SubmissionRetryPolicy.MaximumDelay
                : TimeSpan.FromSeconds(seconds);
        }
    }
}
