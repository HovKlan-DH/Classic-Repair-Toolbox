using System.Net;
using System.Net.Http.Json;
using System.Text;
using CRT.Server.Handlers.Usage;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Http;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The launch check-in as the server reads it - fed the very form CRT builds
    // (CheckInContract.BuildForm), so a field renamed on one side fails here, and the form older
    // CRTs send, written out by hand with the names they have always used, so the contract cannot
    // drift away from what is already installed (Apache forwards their posts to the old check-in
    // address here).
    //
    // Read through ASP.NET's own form parsing on a DefaultHttpContext - the path the route takes -
    // with no server started (rule 6).
    // ###########################################################################################
    public sealed class CheckInFormReaderTests
    {
        private static async Task<CheckInFormRead> ReadAsync(HttpContent content, string? userAgent = "CRT 2026.10.0")
        {
            var context = new DefaultHttpContext();
            context.Request.Method = HttpMethods.Post;
            context.Request.ContentType = content.Headers.ContentType?.ToString();
            context.Request.Body = new MemoryStream(await content.ReadAsByteArrayAsync());

            if (userAgent is not null)
                context.Request.Headers.UserAgent = userAgent;

            return await CheckInFormReader.ReadAsync(context.Request, CancellationToken.None);
        }

        [Fact]
        public async Task The_form_CRT_builds_is_read_field_for_field_with_the_version_from_the_user_agent()
        {
            using FormUrlEncodedContent form = CheckInContract.BuildForm("macOS", "Darwin 23.6.0 Darwin Kernel Version 23.6.0", "ARM 64-bit");

            CheckInFormRead read = await CheckInFormReaderTests.ReadAsync(form);

            Assert.False(read.IsRefused);
            Assert.Equal(
                new CheckInRequest("CRT 2026.10.0", "CRT", "macOS", "Darwin 23.6.0 Darwin Kernel Version 23.6.0", "ARM 64-bit"),
                read.Request);
        }

        // ###########################################################################################
        // *** WHAT EVERY INSTALLED CRT SENDS, BY HAND. *** Its names are the ones CRT has always sent;
        // if this fails, the contract moved and older CRTs' check-ins would silently stop counting.
        // ###########################################################################################
        [Fact]
        public async Task The_form_older_CRTs_send_is_read_too()
        {
            using var form = new StringContent(
                "control=CRT&osHighlevel=Linux&osVersion=Linux+6.8.0-45-generic+%2345-Ubuntu&cpu=64-bit",
                Encoding.UTF8,
                "application/x-www-form-urlencoded");

            CheckInFormRead read = await CheckInFormReaderTests.ReadAsync(form, "CRT 2026.9.0");

            Assert.Equal(
                new CheckInRequest("CRT 2026.9.0", "CRT", "Linux", "Linux 6.8.0-45-generic #45-Ubuntu", "64-bit"),
                read.Request);
        }

        // The old check-in address read a multipart form as well.
        [Fact]
        public async Task A_multipart_form_is_read_too()
        {
            using var form = new MultipartFormDataContent
            {
                { new StringContent("CRT"), CheckInContract.ControlField },
                { new StringContent("Windows"), CheckInContract.OsHighlevelField }
            };

            CheckInFormRead read = await CheckInFormReaderTests.ReadAsync(form);

            Assert.Equal(("CRT", "Windows", string.Empty), (read.Request!.Control, read.Request.OsHighlevel, read.Request.Cpu));
        }

        [Fact]
        public async Task Without_a_user_agent_the_version_is_empty()
        {
            using FormUrlEncodedContent form = CheckInContract.BuildForm("Windows", "10", "64-bit");

            CheckInFormRead read = await CheckInFormReaderTests.ReadAsync(form, userAgent: null);

            Assert.Equal(string.Empty, read.Request!.UserAgent);
        }

        [Fact]
        public async Task A_body_that_is_not_a_form_is_refused_as_unsupported()
        {
            using JsonContent json = JsonContent.Create(new { control = "CRT" });

            CheckInFormRead read = await CheckInFormReaderTests.ReadAsync(json);

            Assert.True(read.IsRefused);
            Assert.Equal(StatusCodes.Status415UnsupportedMediaType, read.RefusalStatus);
        }

        // ###########################################################################################
        // *** BOTH ENDS AT ONCE. *** CRT's own form, read as the route reads it and put through the
        // flow, is the row crt_update gets - so CRT and the server cannot drift apart unnoticed.
        // ###########################################################################################
        [Fact]
        public async Task What_CRT_sends_is_what_crt_update_stores()
        {
            using FormUrlEncodedContent form = CheckInContract.BuildForm("Windows", "Microsoft Windows 10.0.19045", "64-bit");
            var store = new FakeCheckInStore();

            CheckInFormRead read = await CheckInFormReaderTests.ReadAsync(form);
            CheckInOutcome outcome = await CheckInFlow.RecordAsync(
                read.Request!,
                IPAddress.Parse("85.184.162.75"),
                new FakeCountryLookup(new CountryAnswer("DK", "Denmark")),
                store);

            Assert.True(outcome.IsStored);
            Assert.Equal(
                new CheckInRow(
                    "85.184.162.75", "CRT 2026.10.0", "Windows", "Microsoft Windows 10.0.19045", "64-bit", "DK", "Denmark",
                    """{"status":"success","countryCode":"DK","country":"Denmark"}"""),
                Assert.Single(store.Rows));
        }
    }
}
