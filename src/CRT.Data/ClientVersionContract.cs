using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** EVERY CRT EVER RELEASED STAYS INSTALLED (owner request, 2026-10-04: "It is important that
    // all older versions will continue to work"). *** The server is deployed once and is always the
    // newest party; CRTs update when their users let them. So the server bends and the clients only
    // tolerate - and when the server truly cannot serve an old CRT any more, it SAYS so in words
    // instead of failing in a way nobody can read. This file is the part of that both ends speak:
    //
    //   - CrtVersion: the version a request names. Every CRT request carries "User-Agent: CRT
    //     <version>" (OnlineServices.VersionForServer); the server reads it with FromUserAgent.
    //   - ApiRevision: which shape of the API a CRT was built for - see below. The review and
    //     submission clients send it in ApiRevisionHeader; the server refuses an older one.
    //   - ClientOutdatedAnswer: the "update CRT" refusal - HTTP 426 with this body. CRT.Server's
    //     ClientVersionPolicy writes it; CRT reads it with ApiRefusal and shows its Message.
    //   - ApiRefusal: the server's own sentence out of ANY refusal body, so a refusal a future
    //     server invents still reaches the person in words, not as "the server answered 426".
    //
    // The policy itself - which parts of the API serve which CRTs, and what may never be refused -
    // is NewContributeStrategy.md's "Installed CRTs keep working" and CLAUDE.md's matching section.
    // ###########################################################################################
    public static class ClientVersionContract
    {
        // The product token CRT names itself with in its User-Agent ("CRT 3.0.0").
        public const string ProductToken = "CRT";

        // The refusal's code, and its status: 426 Upgrade Required, which is what the old PHP
        // contribution page answered an outdated CRT too.
        public const string OutdatedCode = "client.outdated";
        public const int OutdatedStatus = 426;

        // The User-Agent a CRT of this version sends.
        public static string UserAgentFor(string version) => $"{ClientVersionContract.ProductToken} {version}";

        // ###########################################################################################
        // *** THE API REVISION (owner request, 2026-10-04: "how do I know which version of app uses
        // which version of server? ... I do not have an overview on what kind of specific changes are
        // done to API and if those changes are compatible or not"). ***
        //
        // A CRT VERSION does not move when the API does: a build from source and the alpha published
        // last week both say "3.0.0-alpha.2" while speaking different APIs, so a minimum version
        // needs somebody to know which build has which shape - nobody can. This number is the shape.
        // Both ends are built from it, so a CRT always carries the revision of the source it was
        // built from, and the server serves exactly its own revision and newer (ClientVersionPolicy).
        //
        // *** IT IS RAISED BY ONE RULE, AND A TEST HOLDS IT. *** It goes up by one when the API
        // changes in a way a CRT built for it cannot use - a route, field or enum member gone or
        // retyped, a member added to an enum an answer carries. CRT.Server.Tests' ApiCompatibilityTests
        // keeps the API each revision stands for (api-revision-<N>.txt) and fails on such a change
        // until this is raised. What the test cannot see is raised BY HAND: a route given a new
        // meaning with the same shape, or a request field the server starts to require.
        //
        // Additions never raise it. And from CRT 3.0.0's release it should not move at all - a break
        // of a released CRT is refused by the same test (crt-<version>.txt) unless the project owner
        // decides it, and raising this then turns away every CRT built before.
        // ###########################################################################################
        public const int ApiRevision = 1;

        // The header the review and submission clients send ApiRevision in.
        public const string ApiRevisionHeader = "X-CRT-Api-Revision";

        // The revision a request names, or null when it names none or names it unreadably - which the
        // server lets through, as it does a request naming no CRT version.
        public static int? ApiRevisionFrom(string? headerValue) =>
            int.TryParse(headerValue?.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int revision) && revision > 0
                ? revision
                : null;

        // ###########################################################################################
        // The refusal for a CRT older than an area of the API now serves. Worded for a hobbyist at
        // the bench, with the two versions in brackets as CRT's other messages name them.
        // ###########################################################################################
        public static ClientOutdatedAnswer Outdated(CrtVersion client, CrtVersion minimum)
        {
            ArgumentNullException.ThrowIfNull(client);
            ArgumentNullException.ThrowIfNull(minimum);

            return new ClientOutdatedAnswer(
                ClientVersionContract.OutdatedCode,
                $"This version of CRT [{client}] is too old for this - please update CRT to version [{minimum}] or newer.",
                minimum.ToString());
        }

        // ###########################################################################################
        // The refusal for a CRT built for an older API revision than the server's. The server knows
        // no version that has its revision - only that the newest CRT does - so it names none, and
        // MinimumVersion is null. `client` is the version the request named, if it named one.
        // ###########################################################################################
        public static ClientOutdatedAnswer OutdatedApi(CrtVersion? client)
        {
            string which = client is null ? "This version of CRT" : $"This version of CRT [{client}]";

            return new ClientOutdatedAnswer(
                ClientVersionContract.OutdatedCode,
                $"{which} was made for an older version of the server - please update CRT to the newest version.",
                null);
        }
    }

    // The 426 body. Code is always ClientVersionContract.OutdatedCode. MinimumVersion is null when
    // the refusal is for the API revision (OutdatedApi), which names no version.
    public sealed record ClientOutdatedAnswer(string Code, string Message, string? MinimumVersion);

    // ###########################################################################################
    // What a refusal body says, read the same way whatever shape it came in.
    //
    // The server refuses in three shapes today - {"message"}, {"error"} and {"errors": [...]} - and
    // a client reading only the one its call expects falls back to a status code for the others.
    // Read never throws and never returns null: a body that is not JSON, or carries nothing, gives
    // an ApiRefusal with every field null, and the caller keeps its own fallback sentence.
    //
    // `errors` is deliberately NOT folded into Message: the submission path shows those as a list
    // of findings beside the message, and a message made of the same findings would say it twice.
    // ###########################################################################################
    public sealed record ApiRefusal(string? Code, string? Message, string? MinimumVersion)
    {
        public bool IsClientOutdated =>
            string.Equals(this.Code, ClientVersionContract.OutdatedCode, StringComparison.Ordinal);

        public static ApiRefusal Read(string? body)
        {
            var empty = new ApiRefusal(null, null, null);

            if (string.IsNullOrWhiteSpace(body))
                return empty;

            try
            {
                using JsonDocument document = JsonDocument.Parse(body);
                JsonElement root = document.RootElement;

                if (root.ValueKind != JsonValueKind.Object)
                    return empty;

                return new ApiRefusal(
                    ApiRefusal.Text(root, "code"),
                    ApiRefusal.Text(root, "message") ?? ApiRefusal.Text(root, "error"),
                    ApiRefusal.Text(root, "minimumVersion"));
            }
            catch (JsonException)
            {
                return empty;
            }
        }

        private static string? Text(JsonElement root, string name) =>
            root.TryGetProperty(name, out JsonElement value) &&
            value.ValueKind == JsonValueKind.String &&
            value.GetString() is { } text &&
            !string.IsNullOrWhiteSpace(text)
                ? text.Trim()
                : null;
    }

    // ###########################################################################################
    // A CRT version, ordered the way SemVer 2.0.0 orders them - which is how CRT's versions are
    // written ("2.4.0-beta.16", "3.0.0-alpha.2", "3.0.0"): a pre-release ranks BELOW its release,
    // numeric identifiers compare as numbers ("beta.2" < "beta.16"), and build metadata ("+abc")
    // is ignored. A leading "v" is accepted, as CRT's old version strings sometimes carried one.
    //
    // Only three-part versions parse: everything CRT has released is one, and refusing anything
    // else means a stray product token can never be mistaken for a version.
    // ###########################################################################################
    public sealed class CrtVersion : IComparable<CrtVersion>, IEquatable<CrtVersion>
    {
        private readonly string[] thisPreRelease;

        private CrtVersion(int major, int minor, int patch, string[] preRelease)
        {
            this.Major = major;
            this.Minor = minor;
            this.Patch = patch;
            this.thisPreRelease = preRelease;
        }

        public int Major { get; }

        public int Minor { get; }

        public int Patch { get; }

        public bool IsPreRelease => this.thisPreRelease.Length > 0;

        public static CrtVersion Parse(string text) =>
            CrtVersion.TryParse(text, out CrtVersion? version)
                ? version!
                : throw new FormatException($"[{text}] is not a CRT version.");

        public static bool TryParse(string? text, out CrtVersion? version)
        {
            version = null;

            string value = text?.Trim() ?? string.Empty;

            if (value.StartsWith('v') || value.StartsWith('V'))
                value = value.Substring(1);

            int plus = value.IndexOf('+');

            if (plus >= 0)
                value = value.Substring(0, plus);

            string core = value;
            string[] preRelease = [];
            int dash = value.IndexOf('-');

            if (dash >= 0)
            {
                core = value.Substring(0, dash);
                preRelease = value.Substring(dash + 1).Split('.');

                foreach (string identifier in preRelease)
                {
                    if (identifier.Length == 0 || !CrtVersion.IsIdentifier(identifier))
                        return false;
                }
            }

            string[] parts = core.Split('.');

            if (parts.Length != 3 ||
                !CrtVersion.TryNumber(parts[0], out int major) ||
                !CrtVersion.TryNumber(parts[1], out int minor) ||
                !CrtVersion.TryNumber(parts[2], out int patch))
            {
                return false;
            }

            version = new CrtVersion(major, minor, patch, preRelease);
            return true;
        }

        // ###########################################################################################
        // The version a User-Agent names, or null when it names none: "CRT 3.0.0-alpha.2" (CRT's
        // own) or "CRT/3.0.0", the product token in any case, anywhere among the agent's tokens. A
        // browser, curl, or a CRT whose version does not parse all give null - and null is let
        // through by the server's version policy (fail open, as the old PHP page did).
        // ###########################################################################################
        public static CrtVersion? FromUserAgent(string? userAgent)
        {
            string[] tokens = (userAgent ?? string.Empty).Split(
                [' ', '\t'], StringSplitOptions.RemoveEmptyEntries);

            for (int index = 0; index < tokens.Length; index++)
            {
                string token = tokens[index];
                string? candidate = null;

                if (string.Equals(token, ClientVersionContract.ProductToken, StringComparison.OrdinalIgnoreCase))
                    candidate = index + 1 < tokens.Length ? tokens[index + 1] : null;
                else if (token.StartsWith(ClientVersionContract.ProductToken + "/", StringComparison.OrdinalIgnoreCase))
                    candidate = token.Substring(ClientVersionContract.ProductToken.Length + 1);

                if (candidate is not null && CrtVersion.TryParse(candidate, out CrtVersion? version))
                    return version;
            }

            return null;
        }

        public int CompareTo(CrtVersion? other)
        {
            if (other is null)
                return 1;

            int result = this.Major.CompareTo(other.Major);

            if (result == 0)
                result = this.Minor.CompareTo(other.Minor);

            if (result == 0)
                result = this.Patch.CompareTo(other.Patch);

            if (result != 0)
                return result;

            // A release outranks every pre-release of it.
            if (!this.IsPreRelease || !other.IsPreRelease)
                return other.IsPreRelease.CompareTo(this.IsPreRelease);

            int shared = Math.Min(this.thisPreRelease.Length, other.thisPreRelease.Length);

            for (int index = 0; index < shared; index++)
            {
                result = CrtVersion.CompareIdentifiers(this.thisPreRelease[index], other.thisPreRelease[index]);

                if (result != 0)
                    return result;
            }

            // Equal so far: the one with more identifiers ranks higher ("beta" < "beta.1").
            return this.thisPreRelease.Length.CompareTo(other.thisPreRelease.Length);
        }

        public bool Equals(CrtVersion? other) => other is not null && this.CompareTo(other) == 0;

        public override bool Equals(object? obj) => obj is CrtVersion other && this.Equals(other);

        // ###########################################################################################
        // Agrees with Equals, which compares numeric identifiers as NUMBERS - so "beta.01" and "beta.1"
        // are one version and must hash alike (code review, 2026-10-04: the raw text was hashed, and
        // a set or dictionary keyed by version held them as two).
        // ###########################################################################################
        public override int GetHashCode()
        {
            var hash = new HashCode();

            hash.Add(this.Major);
            hash.Add(this.Minor);
            hash.Add(this.Patch);

            foreach (string identifier in this.thisPreRelease)
            {
                if (CrtVersion.TryNumber(identifier, out int number))
                    hash.Add(number);
                else
                    hash.Add(identifier, StringComparer.Ordinal);
            }

            return hash.ToHashCode();
        }

        public static bool operator <(CrtVersion left, CrtVersion right) => left.CompareTo(right) < 0;

        public static bool operator >(CrtVersion left, CrtVersion right) => left.CompareTo(right) > 0;

        public static bool operator <=(CrtVersion left, CrtVersion right) => left.CompareTo(right) <= 0;

        public static bool operator >=(CrtVersion left, CrtVersion right) => left.CompareTo(right) >= 0;

        public override string ToString() =>
            this.IsPreRelease
                ? FormattableString.Invariant($"{this.Major}.{this.Minor}.{this.Patch}-{string.Join(".", this.thisPreRelease)}")
                : FormattableString.Invariant($"{this.Major}.{this.Minor}.{this.Patch}");

        // Numeric identifiers compare as numbers and rank below alphanumeric ones; alphanumeric
        // ones compare as ordinal text.
        private static int CompareIdentifiers(string left, string right)
        {
            bool leftNumeric = CrtVersion.TryNumber(left, out int leftNumber);
            bool rightNumeric = CrtVersion.TryNumber(right, out int rightNumber);

            if (leftNumeric && rightNumeric)
                return leftNumber.CompareTo(rightNumber);

            if (leftNumeric != rightNumeric)
                return leftNumeric ? -1 : 1;

            return Math.Sign(string.CompareOrdinal(left, right));
        }

        private static bool TryNumber(string text, out int number) =>
            int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out number);

        private static bool IsIdentifier(string identifier)
        {
            foreach (char character in identifier)
            {
                if (!char.IsAsciiLetterOrDigit(character) && character != '-')
                    return false;
            }

            return true;
        }
    }
}
