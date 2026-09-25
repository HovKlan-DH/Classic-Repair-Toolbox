using System.Reflection;

namespace CRT.Server.Handlers.Health
{
    // ###########################################################################################
    // What GET /api/health answers with.
    //
    // This exists as a pure class rather than as a lambda inside Program.cs for the reason the
    // whole server is structured that way: anything that DECIDES is a pure function taking plain
    // values, so it can be unit tested with no server and no network. The endpoint itself is a
    // one-line rim that serialises what Build() returns.
    //
    // WHAT THIS DELIBERATELY DOES NOT REPORT. The health endpoint is unauthenticated and reachable
    // from the public internet through the Apache proxy, so it is a reconnaissance surface. It
    // reports only that the process is alive and which build is deployed:
    //   - no database state (a DB-checking health endpoint hands an attacker a free way to tell
    //     whether the database is down, and turns one outage into two);
    //   - no configuration values, no paths, no connection details;
    //   - no runtime/OS version, which names the patch level of the box.
    // "Is the deployment chain working" is the entire question this answers. Deeper checks belong
    // on an authenticated endpoint if they are ever wanted.
    // ###########################################################################################
    public static class HealthReport
    {
        // The value reported when the assembly carries no informational version. It is a visible
        // placeholder rather than an empty string so a misbuilt deployment is obvious in the
        // response instead of silently blank.
        public const string UnknownVersion = "unknown";

        // ###########################################################################################
        // Builds the health payload. "now" is passed in rather than read from the clock here so the
        // value is testable; every pure class in this project takes time as an argument for the
        // same reason.
        // ###########################################################################################
        public static HealthStatus Build(DateTimeOffset now, string? informationalVersion)
        {
            return new HealthStatus("ok", HealthReport.NormaliseVersion(informationalVersion), now.ToUniversalTime());
        }

        // ###########################################################################################
        // Trims the version and STRIPS ANY BUILD-METADATA SUFFIX (everything from the first "+").
        //
        // This is not cosmetic. The .NET SDK appends SourceRevisionId to InformationalVersion
        // automatically, so a plain read produces "0.1.0-phase3+7150bc943a81fe..." - the exact git
        // commit, published by an unauthenticated endpoint that anyone on the internet can call
        // through the Apache proxy. That tells an attacker precisely which source the deployed
        // binary was built from, in a repository that is public; it is free reconnaissance and it
        // contradicts this class's own rule about what health may reveal.
        //
        // "+build" is semver build metadata and carries no ordering meaning, so dropping it loses
        // nothing a human reading the health output needs. Caught by running the endpoint and
        // looking at the response, not by any test that existed first - which is why there is now
        // a test below pinning it.
        // ###########################################################################################
        private static string NormaliseVersion(string? informationalVersion)
        {
            if (string.IsNullOrWhiteSpace(informationalVersion))
                return HealthReport.UnknownVersion;

            string trimmed = informationalVersion.Trim();

            int metadataStart = trimmed.IndexOf('+');
            if (metadataStart >= 0)
                trimmed = trimmed[..metadataStart];

            trimmed = trimmed.Trim();

            return trimmed.Length == 0 ? HealthReport.UnknownVersion : trimmed;
        }

        // ###########################################################################################
        // Reads the informational version off an assembly, returning null when it carries none.
        // Split out from Build so the reflection stays at the rim and Build itself is pure.
        // ###########################################################################################
        public static string? ReadInformationalVersion(Assembly assembly)
        {
            ArgumentNullException.ThrowIfNull(assembly);

            return assembly
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                ?.InformationalVersion;
        }
    }

    // ###########################################################################################
    // The health payload as it appears on the wire. A record rather than an anonymous object so
    // the shape is named, testable and cannot drift between the endpoint and its test.
    //
    // Property names are lower-cased by the JSON options configured in Program.cs, giving
    // {"status":"ok","version":"...","utc":"..."}.
    // ###########################################################################################
    public sealed record HealthStatus(string Status, string Version, DateTimeOffset Utc);
}
