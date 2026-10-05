using Handlers.DataHandling;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // Reads a launch check-in (CRT.Data's CheckInContract) into plain values - the rim's half of the
    // check-in route, kept out of the endpoint so a test can feed it the very form CRT builds
    // (CheckInContract.BuildForm) and the one older CRTs send, and so the field names are proved on
    // both ends at once.
    //
    // ASP.NET's own form reading, unlike feedback's: the form is four short fields under the
    // default 64 KB body limit, so nothing large is ever buffered. A multipart form is read too,
    // as the old check-in address read both; CRT has always sent a URL-encoded one.
    // ###########################################################################################
    public static class CheckInFormReader
    {
        public static async Task<CheckInFormRead> ReadAsync(HttpRequest request, CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(request);

            if (!request.HasFormContentType)
                return CheckInFormRead.Refused(StatusCodes.Status415UnsupportedMediaType, "A check-in is sent as a form.");

            IFormCollection form;

            try
            {
                form = await request.ReadFormAsync(cancellationToken);
            }
            catch (InvalidDataException)
            {
                return CheckInFormRead.Refused(StatusCodes.Status400BadRequest, "The form could not be read.");
            }
            catch (BadHttpRequestException ex)
            {
                // Larger than the route's body limit, or broken off.
                return CheckInFormRead.Refused(ex.StatusCode, "The form could not be read.");
            }

            return new CheckInFormRead(
                new CheckInRequest(
                    request.Headers.UserAgent.ToString(),
                    form[CheckInContract.ControlField].ToString(),
                    form[CheckInContract.OsHighlevelField].ToString(),
                    form[CheckInContract.OsVersionField].ToString(),
                    form[CheckInContract.CpuField].ToString()),
                0,
                string.Empty);
        }
    }

    // A check-in as read, or why it could not be - with the status to answer.
    public sealed record CheckInFormRead(CheckInRequest? Request, int RefusalStatus, string RefusalReason)
    {
        public bool IsRefused => this.Request is null;

        public static CheckInFormRead Refused(int status, string reason) => new(null, status, reason);
    }
}
