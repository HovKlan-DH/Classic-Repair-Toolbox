namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What the Feedback tab says when a feedback did not get through (2026-10-03, when the feedback
    // moved from the old PHP page to CRT.Server). The server answers each refusal with its own
    // status - see CRT.Server's FeedbackEndpoints - and the ones a user can do something about get
    // a sentence saying what; the rest keep the old "check the logfile" line.
    // ###########################################################################################
    public static class FeedbackWording
    {
        public const string Sent = "Feedback submitted successfully - thank you :-)";

        // Said before anything is sent, when the zip is over the limit the server keeps.
        public static string TooLarge =>
            $"The attached files are too large to send - at most {FeedbackContract.MaximumAttachmentBytes / (1024 * 1024)} MB " +
            "packed. Please attach fewer or smaller files.";

        public static string Failed(int statusCode) => statusCode switch
        {
            404 => "Failed to send feedback: Server endpoint not found (HTTP 404)",
            413 => FeedbackWording.TooLarge,
            429 => "Failed to send feedback: a lot of feedback was sent from here in a short time - please try again in an hour",
            502 => "Failed to send feedback: the server could not pass it on just now - please try again later",
            507 => "Failed to send feedback: the server has no room for the attachments just now - please try again later",
            _ => $"Failed to send feedback (HTTP {statusCode}) - please check the logfile for details"
        };
    }
}
