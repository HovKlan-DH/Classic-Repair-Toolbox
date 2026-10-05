using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** FEEDBACK FROM CRT'S FEEDBACK TAB, AS IT TRAVELS (owner request, 2026-10-03: "have the
    // new backend server handle that"). *** One multipart form: the address typed (or the
    // signed-in account's), the text, the CRT version, and at most one zip of the attached files.
    // CRT builds it with BuildForm, the server (CRT.Server's FeedbackFormReader) reads it with
    // these names, and both ends judge the answer with IsSuccess - so a renamed field cannot
    // compile on one side and fail on the other.
    //
    // *** THE FORM IS THE ONE CRT HAS ALWAYS SENT, UNCHANGED, ON PURPOSE. *** Every CRT already
    // installed posts exactly this to the old feedback address,
    // https://classic-repair-toolbox.dk/app-feedback/, and reads "Success" at the start of the
    // answer as success. Apache forwards that address to the server's route (INSTALLING.md), so
    // older CRTs keep working - which only holds while these names, the zip's file name and the
    // "Success" answer stay as they are.
    // ###########################################################################################
    public static class FeedbackContract
    {
        // The route, under "/api" - the server maps "/api/" + this, CRT posts to CrtServerBaseUrl + "/" + this.
        public const string PathUnderApi = "feedback";

        public const string EmailField = "email";
        public const string FeedbackField = "feedback";
        public const string VersionField = "version";
        public const string AttachmentField = "attachmentFile";
        public const string AttachmentFileName = "FeedbackPayload.zip";

        // What the server answers, as the whole body, when the feedback was mailed.
        public const string SuccessAnswer = "Success";

        // ###########################################################################################
        // The zip may be this large; the whole request a little more (the text and the form's own
        // framing). CRT refuses a larger zip before sending anything, with a message that says
        // why, rather than uploading it all to be told "413".
        //
        // 250 MB (owner decision, 2026-10-03: "one could potentially zip the entire system, and then
        // send that, so please allow for more - e.g. 250MB ZIP'ed"). CRT streams the zip from a
        // temporary file and the server writes it straight to disk, so neither end holds it in
        // memory.
        // ###########################################################################################
        public const long MaximumAttachmentBytes = 250L * 1024 * 1024;

        public const long MaximumRequestBytes = 256L * 1024 * 1024;

        // The text box has no limit of its own; the server keeps this much of it.
        public const int MaximumFeedbackCharacters = 100_000;

        // ###########################################################################################
        // The files the mail SHOWS rather than saves - CRT's own text files, which the Feedback
        // tab's two check boxes attach under exactly these names (AppConfig's, held to these by
        // CRT.App.Tests' FeedbackAttachmentNamesTests). The mail always showed the first three; the
        // crash log joined them on the move to CRT.Server (2026-10-03), being the same kind of text.
        // ###########################################################################################
        public static readonly IReadOnlyList<FeedbackInlineFile> InlineFiles =
        [
            new("Classic-Repair-Toolbox.log", "Logfile"),
            new("Classic-Repair-Toolbox.crash.log", "Crash log"),
            new("Classic-Repair-Toolbox.settings.json", "Settings"),
            new("Classic-Repair-Toolbox.traces.json", "Traces")
        ];

        // ###########################################################################################
        // The form CRT sends. `attachmentZip` is read from where it stands as the form is sent, and
        // disposed with the form; null, or a seekable stream with nothing left in it, sends no
        // attachment part at all.
        // ###########################################################################################
        public static MultipartFormDataContent BuildForm(string email, string feedback, string version, Stream? attachmentZip)
        {
            var form = new MultipartFormDataContent
            {
                { new StringContent(email ?? string.Empty), FeedbackContract.EmailField },
                { new StringContent(feedback ?? string.Empty), FeedbackContract.FeedbackField },
                { new StringContent(version ?? string.Empty), FeedbackContract.VersionField }
            };

            bool empty = attachmentZip is null || (attachmentZip.CanSeek && attachmentZip.Length - attachmentZip.Position <= 0);

            if (!empty)
            {
                var zip = new StreamContent(attachmentZip!);
                zip.Headers.ContentType = new MediaTypeHeaderValue("application/zip");
                form.Add(zip, FeedbackContract.AttachmentField, FeedbackContract.AttachmentFileName);
            }

            return form;
        }

        // ###########################################################################################
        // Did the feedback arrive? A 2xx AND a body starting with "Success" - the old feedback
        // address could answer 200 with a warning in the body, which is why CRT always read the
        // body too.
        // ###########################################################################################
        public static bool IsSuccess(int statusCode, string? body) =>
            statusCode is >= 200 and < 300 &&
            (body ?? string.Empty).TrimStart().StartsWith(FeedbackContract.SuccessAnswer, StringComparison.OrdinalIgnoreCase);
    }

    // One of CRT's own files the feedback mail shows inline, under its heading.
    public sealed record FeedbackInlineFile(string FileName, string Title);
}
