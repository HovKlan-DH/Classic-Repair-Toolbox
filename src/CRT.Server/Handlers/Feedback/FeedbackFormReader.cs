using System.Text;
using Handlers.DataHandling;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Net.Http.Headers;

namespace CRT.Server.Handlers.Feedback
{
    // ###########################################################################################
    // Reads the Feedback tab's form (CRT.Data's FeedbackContract) into plain values - the rim's
    // half of the feedback route, kept out of the endpoint so a test can feed it the very form CRT
    // builds (FeedbackContract.BuildForm) and so the field names are proved on both ends at once.
    //
    // *** THE ZIP GOES STRAIGHT TO A FILE, NEVER INTO MEMORY OR /tmp. *** ASP.NET's own form reader
    // buffers a large file part to the temporary folder - under systemd's PrivateTmp, and sized only
    // by the request limit - so the multipart sections are walked here instead and the attachment
    // part is copied, capped, to the path the caller gives (inside the feedback folder, which the
    // service may write). The text fields are kept only up to their own limits.
    // ###########################################################################################
    public static class FeedbackFormReader
    {
        private const int MaximumEmailCharacters = 320;
        private const int MaximumVersionCharacters = 200;

        public static async Task<FeedbackFormRead> ReadAsync(
            string? contentType,
            Stream body,
            string attachmentPath,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(body);
            ArgumentException.ThrowIfNullOrWhiteSpace(attachmentPath);

            if (!MediaTypeHeaderValue.TryParse(contentType, out MediaTypeHeaderValue? mediaType) ||
                !mediaType.MediaType.Equals("multipart/form-data", StringComparison.OrdinalIgnoreCase))
            {
                return FeedbackFormRead.Refused(StatusCodes.Status415UnsupportedMediaType, "Feedback is sent as a multipart form.");
            }

            string boundary = HeaderUtilities.RemoveQuotes(mediaType.Boundary).Value ?? string.Empty;

            if (boundary.Length == 0)
                return FeedbackFormRead.Refused(StatusCodes.Status400BadRequest, "The form has no boundary.");

            var reader = new MultipartReader(boundary, body);

            string email = string.Empty;
            string feedback = string.Empty;
            string version = string.Empty;
            bool hasAttachment = false;

            try
            {
                MultipartSection? section;

                while ((section = await reader.ReadNextSectionAsync(cancellationToken)) is not null)
                {
                    if (!ContentDispositionHeaderValue.TryParse(section.ContentDisposition, out ContentDispositionHeaderValue? disposition))
                        continue;

                    string name = HeaderUtilities.RemoveQuotes(disposition.Name).Value ?? string.Empty;

                    switch (name)
                    {
                        case FeedbackContract.AttachmentField when !hasAttachment:
                            if (!await FeedbackFormReader.CopyCappedAsync(section.Body, attachmentPath, FeedbackContract.MaximumAttachmentBytes, cancellationToken))
                            {
                                File.Delete(attachmentPath);
                                return FeedbackFormRead.Refused(
                                    StatusCodes.Status413PayloadTooLarge,
                                    $"The attached files are larger than {FeedbackContract.MaximumAttachmentBytes / (1024 * 1024)} MB packed.");
                            }

                            hasAttachment = true;
                            break;

                        case FeedbackContract.EmailField:
                            email = await FeedbackFormReader.ReadTextAsync(section.Body, FeedbackFormReader.MaximumEmailCharacters, cancellationToken);
                            break;

                        case FeedbackContract.FeedbackField:
                            feedback = await FeedbackFormReader.ReadTextAsync(section.Body, FeedbackContract.MaximumFeedbackCharacters, cancellationToken);
                            break;

                        case FeedbackContract.VersionField:
                            version = await FeedbackFormReader.ReadTextAsync(section.Body, FeedbackFormReader.MaximumVersionCharacters, cancellationToken);
                            break;

                        // Anything else is read past (the next section skips it).
                    }
                }
            }
            catch (InvalidDataException)
            {
                // A form that breaks off or is not multipart after all.
                File.Delete(attachmentPath);
                return FeedbackFormRead.Refused(StatusCodes.Status400BadRequest, "The form could not be read.");
            }

            return new FeedbackFormRead(new FeedbackForm(email, feedback, version, hasAttachment ? attachmentPath : null), 0, string.Empty);
        }

        // Copies at most `limit` bytes; false when there was more.
        private static async Task<bool> CopyCappedAsync(Stream source, string path, long limit, CancellationToken cancellationToken)
        {
            await using var target = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true);

            var buffer = new byte[81920];
            long total = 0;
            int read;

            while ((read = await source.ReadAsync(buffer, cancellationToken)) > 0)
            {
                total += read;

                if (total > limit)
                    return false;

                await target.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
            }

            return true;
        }

        // The field's text, up to `limit` characters; the rest is read and dropped.
        private static async Task<string> ReadTextAsync(Stream source, int limit, CancellationToken cancellationToken)
        {
            using var reader = new StreamReader(source, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, bufferSize: 4096, leaveOpen: true);

            var text = new StringBuilder();
            var buffer = new char[4096];
            int read;

            while ((read = await reader.ReadAsync(buffer.AsMemory(), cancellationToken)) > 0)
            {
                int room = limit - text.Length;

                if (room > 0)
                    text.Append(buffer, 0, Math.Min(read, room));
            }

            return text.ToString();
        }
    }

    // The form as read: the four values, the zip's path when one came - or why it was refused.
    public sealed record FeedbackForm(string Email, string Feedback, string Version, string? AttachmentPath);

    public sealed record FeedbackFormRead(FeedbackForm? Form, int RefusalStatus, string RefusalReason)
    {
        public bool IsRefused => this.Form is null;

        public static FeedbackFormRead Refused(int status, string reason) => new(null, status, reason);
    }
}
