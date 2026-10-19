using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What the Feedback tab says when a feedback did not get through (2026-10-03, when the feedback
    // moved to CRT.Server). The server answers each refusal with its own
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

        // Said when there is nothing in the field to send (pictures alone are something - their
        // marker lines are text).
        public const string NothingWritten = "Please provide a description of your issue or suggestion before sending";

        // ###########################################################################################
        // Said, and nothing sent, when a picture in the text cannot be had as the feedback is sent
        // (code review, 2026-10-16). It is on the screen, so the user is told which one to put
        // right - a DROPPED picture whose file was moved or deleted since, and a PASTED one whose
        // picture could not be made ready (it has no file to blame), each in its own words.
        // ###########################################################################################
        public static string PicturesNotRead(IReadOnlyList<string> droppedFiles, IReadOnlyList<string> pasted)
        {
            ArgumentNullException.ThrowIfNull(droppedFiles);
            ArgumentNullException.ThrowIfNull(pasted);

            static string Named(IReadOnlyList<string> names) => string.Join(", ", names.Select(name => $"[{name}]"));

            var reasons = new List<string>();

            if (droppedFiles.Count == 1)
                reasons.Add($"the picture {Named(droppedFiles)} could not be read - its file may have been moved or deleted");
            else if (droppedFiles.Count > 1)
                reasons.Add($"the pictures {Named(droppedFiles)} could not be read - their files may have been moved or deleted");

            if (pasted.Count == 1)
                reasons.Add($"the pasted picture {Named(pasted)} could not be made ready to send");
            else if (pasted.Count > 1)
                reasons.Add($"the pasted pictures {Named(pasted)} could not be made ready to send");

            string what = droppedFiles.Count + pasted.Count == 1 ? "it" : "them";

            return $"Not sent: {string.Join("; ", reasons)}. Remove {what} from the text, or put {what} in again, and send again.";
        }

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
