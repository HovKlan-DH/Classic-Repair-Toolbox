using System.Net;
using CRT.Server.Handlers;
using Microsoft.AspNetCore.Http;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // PerAddressRateLimit - the one limiter the anonymous routes (feedback, the check-in, board
    // views, CRT 2.x's contribution upload) share (code review, 2026-10-04: it was four copies).
    // ###########################################################################################
    public sealed class PerAddressRateLimitTests
    {
        [Fact]
        public void Requests_are_counted_by_the_connections_address()
        {
            var context = new DefaultHttpContext();
            context.Connection.RemoteIpAddress = IPAddress.Parse("192.0.2.7");

            Assert.Equal("192.0.2.7", PerAddressRateLimit.PartitionOf(context));
        }

        // A request with no address shares one bucket rather than going unlimited.
        [Fact]
        public void A_request_with_no_address_shares_one_bucket()
        {
            Assert.Equal("unknown", PerAddressRateLimit.PartitionOf(new DefaultHttpContext()));
        }

        // ###########################################################################################
        // *** ONE COPY. *** A fixed-window limiter is built in PerAddressRateLimit and nowhere else in
        // the service, so a change to how an address is taken cannot reach three routes and miss the
        // fourth.
        // ###########################################################################################
        [Fact]
        public void No_other_source_file_of_the_service_builds_a_fixed_window_limiter()
        {
            string server = Path.Combine(ServerRouteTable.RepositoryRoot(), "src", "CRT.Server");

            List<string> builders = Directory
                .EnumerateFiles(server, "*.cs", SearchOption.AllDirectories)
                .Where(path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal) &&
                               !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
                .Where(path => File.ReadAllText(path).Contains("GetFixedWindowLimiter(", StringComparison.Ordinal))
                .Select(path => Path.GetFileName(path))
                .ToList();

            Assert.Equal(["PerAddressRateLimit.cs"], builders);
        }
    }
}
