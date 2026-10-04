using System.Threading.RateLimiting;
using Microsoft.AspNetCore.RateLimiting;

namespace CRT.Server.Handlers
{
    // ###########################################################################################
    // THE ANONYMOUS ROUTES' LIMIT: so many requests an hour from one address, counted in memory.
    // Feedback, the launch check-in, board views and CRT 2.x's contribution upload each name a policy
    // of this shape on their route.
    //
    // ONE COPY (code review, 2026-10-04). Each of the four registered its own identical limiter, so
    // a change to how a request's address is taken - a forwarded-address rule, say - had to be made
    // four times, and one missed would quietly limit that route by something else.
    // ###########################################################################################
    public static class PerAddressRateLimit
    {
        public static readonly TimeSpan Window = TimeSpan.FromHours(1);

        public static void AddPerAddressRateLimit(this IServiceCollection services, string policy, int permitsPerHour)
        {
            ArgumentNullException.ThrowIfNull(services);
            ArgumentException.ThrowIfNullOrWhiteSpace(policy);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(permitsPerHour);

            services.AddRateLimiter(limiter =>
            {
                limiter.RejectionStatusCode = StatusCodes.Status429TooManyRequests;

                limiter.AddPolicy(policy, context =>
                    RateLimitPartition.GetFixedWindowLimiter(
                        PerAddressRateLimit.PartitionOf(context),
                        _ => new FixedWindowRateLimiterOptions
                        {
                            PermitLimit = permitsPerHour,
                            Window = PerAddressRateLimit.Window,
                            QueueLimit = 0
                        }));
            });
        }

        // Whose requests count together: the connection's address, or one shared bucket without one.
        public static string PartitionOf(HttpContext context)
        {
            ArgumentNullException.ThrowIfNull(context);

            return context.Connection.RemoteIpAddress?.ToString() ?? "unknown";
        }
    }
}
