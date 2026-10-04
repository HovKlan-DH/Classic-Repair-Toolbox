using CRT.Server.Handlers.Compat;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers LegacyContributionEndpoints - CRT 2.x's "Send contribution", answered "please update"
    // once the old PHP page is gone (owner decision, 2026-10-04).
    //
    // *** THE ANSWER IS READ BY BUILDS THAT CAN NEVER CHANGE. *** So it is checked against a copy of
    // the parser CRT 2.5.0 ships (ContributionPackaging.TryParseOutdatedVersionResponse at the 2.5.0
    // tag), kept here verbatim rather than referenced: the copy in CRT.App is no longer called by
    // anything in CRT 3 and may be deleted, while the installed 2.5.0 builds keep running this code.
    // ###########################################################################################
    public sealed class LegacyContributionEndpointsTests
    {
        [Fact]
        public void CRT_2_5_0_recognises_the_answer_and_reads_the_version_to_update_to()
        {
            string answer = LegacyContributionEndpoints.OutdatedAnswer("CRT 2.5.0");

            Assert.True(LegacyContributionEndpointsTests.Crt250TryParseOutdatedVersionResponse(answer, out string version));
            Assert.Equal("3.0.0", version);
        }

        // The PHP page's own words, so a 2.x log reads the same before and after the retirement -
        // and older 2.x builds, which do not parse the token, log exactly this text.
        [Fact]
        public void The_answer_is_the_PHP_pages_wording_with_the_senders_version()
        {
            Assert.Equal(
                "OUTDATED_VERSION 3.0.0 - this application version [2.4.0-beta.16] is too old to contribute data - please update to version [3.0.0] or newer.",
                LegacyContributionEndpoints.OutdatedAnswer("CRT 2.4.0-beta.16"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("Mozilla/5.0")]
        public void A_sender_naming_no_version_is_still_told_to_update(string? userAgent)
        {
            string answer = LegacyContributionEndpoints.OutdatedAnswer(userAgent);

            Assert.Contains("[unknown]", answer);
            Assert.True(LegacyContributionEndpointsTests.Crt250TryParseOutdatedVersionResponse(answer, out string version));
            Assert.Equal("3.0.0", version);
        }

        // ###########################################################################################
        // *** AN UPLOAD OVER THE ROUTE'S LIMIT STILL GETS THE WORDS (code review, 2026-10-04). ***
        // Kestrel stops such a body by throwing from the read; that escaped as a bare 413, so CRT
        // 2.5.0 never saw the "please update to 3.0.0" this route exists to send. Driven with a
        // DefaultHttpContext whose body throws as Kestrel's does (rule 6: nothing listens).
        // ###########################################################################################
        [Fact]
        public async Task An_upload_over_the_limit_is_still_answered_426_with_the_outdated_text()
        {
            var context = new DefaultHttpContext
            {
                RequestServices = new ServiceCollection().AddLogging().BuildServiceProvider()
            };

            context.Request.Method = "POST";
            context.Request.Headers.UserAgent = "CRT 2.5.0";
            context.Request.Body = new TooLargeBody();
            context.Response.Body = new MemoryStream();

            IResult result = await LegacyContributionEndpoints.ReceiveAsync(context, CancellationToken.None);
            await result.ExecuteAsync(context);

            Assert.Equal(426, context.Response.StatusCode);

            context.Response.Body.Position = 0;
            string answer = await new StreamReader(context.Response.Body).ReadToEndAsync();

            Assert.Equal(LegacyContributionEndpoints.OutdatedAnswer("CRT 2.5.0"), answer);
            Assert.True(LegacyContributionEndpointsTests.Crt250TryParseOutdatedVersionResponse(answer, out string version));
            Assert.Equal("3.0.0", version);
        }

        // A request body that, like Kestrel's past MaxRequestBodySize, throws on the first read.
        private sealed class TooLargeBody : Stream
        {
            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush()
            {
            }

            public override int Read(byte[] buffer, int offset, int count) =>
                throw new BadHttpRequestException("Request body too large.", StatusCodes.Status413PayloadTooLarge);

            public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) =>
                throw new BadHttpRequestException("Request body too large.", StatusCodes.Status413PayloadTooLarge);

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => throw new NotSupportedException();

            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }

        // The route Apache's forward of /app-contribution/api/ lands on.
        [Fact]
        public void The_route_is_mapped_where_Apache_forwards_the_old_address()
        {
            Assert.Single(
                ServerRouteTable.Routes,
                route => ServerRouteTable.NameOf(route) == "POST /api/legacy/contribution");
        }

        // ###########################################################################################
        // CRT 2.5.0's parser, VERBATIM from the 2.5.0 tag (Handlers/Data/ContributionPackaging.cs,
        // with its OutdatedVersionToken constant inlined). Do not "improve" it - it is what runs.
        // ###########################################################################################
        private static bool Crt250TryParseOutdatedVersionResponse(string? responseBody, out string newestVersion)
        {
            const string OutdatedVersionToken = "OUTDATED_VERSION";

            newestVersion = string.Empty;

            string body = responseBody?.Trim() ?? string.Empty;
            int tokenIndex = body.IndexOf(OutdatedVersionToken, StringComparison.OrdinalIgnoreCase);
            if (tokenIndex < 0)
            {
                return false;
            }

            string remainder = body.Substring(tokenIndex + OutdatedVersionToken.Length).Trim();
            string[] parts = remainder.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length > 0 && parts[0].Length > 0 && char.IsDigit(parts[0][0]))
            {
                newestVersion = parts[0];
            }

            return true;
        }
    }
}
