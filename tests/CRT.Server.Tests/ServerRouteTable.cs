using CRT.Server.Configuration;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http.Metadata;
using Microsoft.AspNetCore.Routing;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The server's REAL route table, built once for the tests that read it - through Program's own
    // AddServerServices and MapServerEndpoints, exactly as RequestBodyLimitsTests builds it.
    //
    // Rule 6 holds: the app is BUILT and never started, so nothing listens and no request is
    // served, and registering a service constructs no store.
    // ###########################################################################################
    internal static class ServerRouteTable
    {
        private static readonly Lazy<IReadOnlyList<RouteEndpoint>> Built = new(ServerRouteTable.Build);

        public static IReadOnlyList<RouteEndpoint> Routes => ServerRouteTable.Built.Value;

        // "POST /api/submissions/{submissionId:long}/finalise" - the method and the pattern without
        // a trailing slash, the form RequestBodyLimitsTests names routes in.
        public static string NameOf(RouteEndpoint endpoint) =>
            $"{ServerRouteTable.MethodOf(endpoint)} {ServerRouteTable.PatternOf(endpoint)}";

        public static string MethodOf(RouteEndpoint endpoint) =>
            endpoint.Metadata.GetMetadata<IHttpMethodMetadata>()?.HttpMethods.SingleOrDefault() ?? "*";

        public static string PatternOf(RouteEndpoint endpoint) =>
            (endpoint.RoutePattern.RawText ?? string.Empty).TrimEnd('/');

        // The repository's root folder - where Classic-Repair-Toolbox.slnx is.
        public static string RepositoryRoot()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);
            return folder!;
        }

        private static IReadOnlyList<RouteEndpoint> Build()
        {
            WebApplicationBuilder builder = WebApplication.CreateBuilder();
            Program.AddServerServices(builder.Services, new ServerOptions());

            using WebApplication app = builder.Build();
            Program.MapServerEndpoints(app);

            return ((IEndpointRouteBuilder)app).DataSources
                .SelectMany(source => source.Endpoints)
                .OfType<RouteEndpoint>()
                .ToList();
        }
    }
}
