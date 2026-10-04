using System.Net;
using System.Text;
using System.Text.Json;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // STORING A LAUNCH CHECK-IN (owner request, 2026-10-03) - POST /api/usage/check-in, and the old
    // /app-checkin/ address through Apache. The app-checkin PHP page's job, moved here so the page
    // can be retired. See CRT.Data's CheckInContract for the form.
    //
    // *** THE PHP PAGE'S RULES, KEPT. *** The table it wrote (crt_update) is read by the Fun facts
    // pages and a nightly statistics job, which nothing here may surprise:
    //   - the VERSION is the User-Agent with any markup stripped and trimmed, and only one naming
    //     "CRT " and made of letters, digits and  ,.#()*[]!:/-  counts (the PHP page's isValid);
    //   - the control field must say exactly "CRT";
    //   - anything else is not stored (the PHP page answered an empty 200; this answers 400, which
    //     no CRT ever gets and none reads);
    //   - a check-in from a LOCAL network is the project owner's own machines, and is answered but
    //     NEVER stored - the PHP page's "192.168." rule, and every Fun facts query filters them too.
    //     Unlike board views there is no setting to count them: the charts would drop them anyway.
    //     The outcome says so, so the endpoint can log it - the way the project owner sees their own
    //     test reach the server;
    //   - the ADDRESS is stored (every chart counts DISTINCT ipaddr) with the COUNTRY looked up at
    //     ip-api.com - the lookup board views use, which the check-in has always used. A failed
    //     lookup costs the country (""), never the row - the charts then leave the row out, as
    //     they did a PHP row with no country;
    //   - the operating system fields lose any markup, as the PHP page's strip_tags did (the Fun
    //     facts page prints them), and are cut to the lengths board views keep.
    //
    // *** WHAT IS DIFFERENT: apiJson. *** The PHP page stored ip-api.com's WHOLE answer - city,
    // region, postcode, coordinates, internet provider. Nothing reads that column (checked on both
    // sites' pages, 2026-10-03), and the shared lookup asks only for the country, so a row now
    // holds only that: {"status":"success","countryCode":"DK","country":"Denmark"}, or "{}".
    // ###########################################################################################
    public static class CheckInFlow
    {
        // Check-ins one address may send an hour - CRT sends one per launch. In memory only; see
        // CheckInEndpoints.
        public const int MaxPerAddressPerHour = 60;

        public static async Task<CheckInOutcome> RecordAsync(
            CheckInRequest request,
            IPAddress? sender,
            ICountryLookup countries,
            ICheckInStore store,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(countries);
            ArgumentNullException.ThrowIfNull(store);

            string? version = CheckInRules.VersionFrom(request.UserAgent);

            if (version is null)
                return CheckInOutcome.Refused("The check-in does not say which version of CRT sent it.");

            if (!string.Equals(request.Control, CheckInContract.ControlValue, StringComparison.Ordinal))
                return CheckInOutcome.Refused("The check-in is not from CRT.");

            IPAddress? address = SenderAddress.PublicOrNull(sender);

            if (address is null)
                return CheckInOutcome.NotStoredFromLocalNetwork;

            CountryAnswer? country = await countries.LookupAsync(address, cancellationToken);

            var row = new CheckInRow(
                address.ToString(),
                version,
                CheckInRules.Text(request.OsHighlevel, BoardViewRules.OsHighlevelLength),
                CheckInRules.Text(request.OsVersion, BoardViewRules.OsVersionLength),
                CheckInRules.Text(request.Cpu, BoardViewRules.CpuLength),
                country?.Code ?? string.Empty,
                country?.Name ?? string.Empty,
                CheckInRules.ApiJson(country));

            await store.RecordAsync(row, cancellationToken);

            return CheckInOutcome.Stored;
        }
    }

    // The check-in as the form and its headers carried it, before any rule is applied.
    public sealed record CheckInRequest(string? UserAgent, string? Control, string? OsHighlevel, string? OsVersion, string? Cpu);

    // ###########################################################################################
    // What became of a check-in. Refused carries why; Stored is a row written; NotStoredFromLocalNetwork
    // is one from the server's own network, accepted and not written (see CheckInFlow's header).
    // ###########################################################################################
    public sealed record CheckInOutcome(bool IsRefused, string? Reason, bool IsStored, bool IsFromLocalNetwork)
    {
        public static readonly CheckInOutcome Stored = new(false, null, true, false);

        public static readonly CheckInOutcome NotStoredFromLocalNetwork = new(false, null, false, true);

        public static CheckInOutcome Refused(string reason) => new(true, reason, false, false);
    }

    // ###########################################################################################
    // The PHP page's text rules, one method each - pure, so each is tested on its own.
    // ###########################################################################################
    public static class CheckInRules
    {
        // The characters a version may hold besides ASCII letters and digits - the PHP page's
        // isValid pattern, [a-zA-Z0-9 ,.#()*\[\]!:\/-].
        private const string VersionPunctuation = " ,.#()*[]!:/-";

        // ###########################################################################################
        // The CRT version a User-Agent names ("CRT 2026.10.0"), or null when it is not one the PHP
        // page would have stored: blank, without "CRT " (any case), with a character outside the
        // pattern, or longer than crt_board_views keeps a version.
        // ###########################################################################################
        public static string? VersionFrom(string? userAgent)
        {
            string version = CheckInRules.StripTags(userAgent ?? string.Empty).Trim();

            if (version.Length == 0 || version.Length > BoardViewRules.VersionLength)
                return null;

            if (!version.Contains("CRT ", StringComparison.OrdinalIgnoreCase))
                return null;

            foreach (char character in version)
            {
                if (!char.IsAsciiLetterOrDigit(character) && !CheckInRules.VersionPunctuation.Contains(character))
                    return null;
            }

            return version;
        }

        // A form field as stored: markup stripped, then cut to `maxLength` (BoardViewRules.Clip -
        // trimmed, no control characters, never half a character), and "" rather than null.
        public static string Text(string? value, int maxLength) =>
            BoardViewRules.Clip(CheckInRules.StripTags(value ?? string.Empty), maxLength) ?? string.Empty;

        // ###########################################################################################
        // PHP's strip_tags, as far as it matters here: a "<" that opens a tag removes everything up
        // to the next ">", or to the end when there is none. A "<" followed by a space, or last in
        // the text, opens nothing and is kept ("5 < 6"), as PHP keeps it.
        // ###########################################################################################
        public static string StripTags(string value)
        {
            ArgumentNullException.ThrowIfNull(value);

            var text = new StringBuilder(value.Length);
            int index = 0;

            while (index < value.Length)
            {
                if (value[index] == '<' && index + 1 < value.Length && !char.IsWhiteSpace(value[index + 1]))
                {
                    int end = value.IndexOf('>', index + 1);

                    if (end < 0)
                        break;

                    index = end + 1;
                    continue;
                }

                text.Append(value[index]);
                index++;
            }

            return text.ToString();
        }

        // What the row's apiJson holds - see CheckInFlow's header.
        public static string ApiJson(CountryAnswer? country) =>
            country is null
                ? "{}"
                : JsonSerializer.Serialize(new { status = "success", countryCode = country.Code, country = country.Name });
    }
}
