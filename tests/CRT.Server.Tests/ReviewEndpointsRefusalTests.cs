using CRT.Server.Handlers.Submissions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The 403 a signed-in account WITHOUT the maintainer role gets (code review, 2026-09-25).
    //
    // It used to be Results.Forbid(), which asks the authentication middleware's default scheme
    // to write the refusal. This service authenticates its own opaque bearer tokens and registers
    // no scheme, so executing it threw and the caller got a 500 - never the 403 that
    // ReviewEndpoints documents and that the maintainer app turns into "This account is not allowed
    // to review submissions".
    //
    // The result is EXECUTED here against a context set up the way this service runs - with no
    // authentication services at all - because constructing it succeeds either way; only
    // executing it tells the two apart. No HTTP pipeline is started (see the test csproj header).
    // ###########################################################################################
    public class ReviewEndpointsRefusalTests
    {
        [Fact]
        public async Task A_non_maintainer_gets_a_403_with_no_authentication_scheme_registered()
        {
            var context = new DefaultHttpContext
            {
                RequestServices = new ServiceCollection().AddLogging().BuildServiceProvider(),
            };

            context.Response.Body = new MemoryStream();

            await ReviewEndpoints.NotAMaintainer().ExecuteAsync(context);

            Assert.Equal(StatusCodes.Status403Forbidden, context.Response.StatusCode);

            context.Response.Body.Position = 0;
            string body = await new StreamReader(context.Response.Body).ReadToEndAsync();

            Assert.Contains("not allowed to review", body);
        }
    }
}
