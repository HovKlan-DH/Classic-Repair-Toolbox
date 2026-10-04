using System;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // The signed-in maintainer's own account (2026-10-03): GET /api/accounts/me, and what each
    // change in the "Your account" window answers. Forgiving like the rest of the parser: an
    // account without an id or an address is no account (null), since the session would be
    // rebuilt from it and a blank address would then be shared with the Feedback tab and the
    // Submit dialog.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        public static AccountAnswer? ParseAccount(string? json) =>
            ReviewApiParser.ReadAccount(ReviewApiParser.Root(json));

        // ###########################################################################################
        // GET /api/health (2026-10-04): the deployed server's version, and the API revision it serves
        // (server 4.6.0; null from an older one, or when it is not a whole number above 0). An answer
        // without a version is no answer - "Server version" followed by nothing would read as one.
        // ###########################################################################################
        public static HealthStatus? ParseHealth(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object ||
                ReviewApiParser.String(root, "version") is not string version ||
                string.IsNullOrWhiteSpace(version))
            {
                return null;
            }

            DateTimeOffset utc = root.TryGetProperty("utc", out JsonElement raw) &&
                                 raw.ValueKind == JsonValueKind.String &&
                                 raw.TryGetDateTimeOffset(out DateTimeOffset parsed)
                ? parsed
                : default;

            int? apiRevision = ReviewApiParser.Long(root, "apiRevision") is long revision and > 0 and <= int.MaxValue
                ? (int)revision
                : null;

            return new HealthStatus(ReviewApiParser.String(root, "status") ?? string.Empty, version.Trim(), utc, apiRevision);
        }

        // ###########################################################################################
        // A change's answer. The message is the server's sentence and is required - it is the
        // whole of what the window shows - while the account is absent on purpose while a new
        // address waits for its code (CodeSent). An account present but unreadable makes the
        // answer unreadable: the window must not say "done" and then keep the old details.
        // ###########################################################################################
        public static AccountChangeAnswer? ParseAccountChange(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object || ReviewApiParser.String(root, "message") is not string message)
                return null;

            AccountAnswer? account = null;

            if (root.TryGetProperty("account", out JsonElement raw) && raw.ValueKind != JsonValueKind.Null)
            {
                account = ReviewApiParser.ReadAccount(raw);

                if (account is null)
                    return null;
            }

            return new AccountChangeAnswer(message, account, ReviewApiParser.Bool(root, "codeSent") ?? false);
        }

        private static AccountAnswer? ReadAccount(JsonElement element)
        {
            if (element.ValueKind != JsonValueKind.Object ||
                ReviewApiParser.Long(element, "id") is not long id ||
                ReviewApiParser.String(element, "email") is not string email ||
                string.IsNullOrWhiteSpace(email))
            {
                return null;
            }

            return new AccountAnswer(
                id,
                email,
                ReviewApiParser.String(element, "displayName") ?? string.Empty,
                ReviewApiParser.Bool(element, "isVerified") ?? false,
                ReviewApiParser.Bool(element, "isAdministrator") ?? false,
                ReviewApiParser.Strings(element, "maintainerOf"),
                ReviewApiParser.Time(element, "createdUtc") ?? DateTimeOffset.MinValue,
                ReviewApiParser.Time(element, "lastLoginUtc"));
        }
    }
}
