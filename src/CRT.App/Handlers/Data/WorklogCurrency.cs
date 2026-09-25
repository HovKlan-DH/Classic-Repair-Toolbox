using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Which currency the money in a worklog is denominated in - the user's own choice, made once in
    // the Configuration tab and shown everywhere a cost appears.
    //
    // Every cost in a workbook used to be printed as a bare number ("430"), because the app had
    // never asked which currency the user works in and printing a guessed symbol would have been
    // worse than printing none. It asks now, so the number can say what it is - which matters most
    // on the exported PDF, a document that leaves the repairer's machine and is read by a customer
    // who has no way to ask.
    //
    // The choice is stored as the three-letter ISO 4217 CODE, not as the country: a country name is
    // how the user picks (nobody knows offhand that Mauritius uses MUR), but the code is what gets
    // printed, and several of these countries share one. Storing the code means a country later
    // changing currency, or being dropped from the list, cannot silently repoint a workbook's money
    // at something else.
    //
    // Pure string work - no controls, no files - so it is unit tested like the rest of Handlers/.
    // ###########################################################################################
    public static class WorklogCurrency
    {
        // The country the drop-down opens on when the user has never chosen, and the code that goes
        // with it - what every existing installation gets on upgrade. USD rather than "none": the
        // setting is a drop-down with no empty row, so there is no state in which a workbook has no
        // currency, and a default that printed nothing would just be the old bare number wearing a
        // different name.
        //
        // The COUNTRY is named as well as the code, because USD belongs to two countries here and
        // Puerto Rico sorts first alphabetically - so resolving the default code by scanning the
        // sorted list would open the drop-down on "Puerto Rico (USD)" for every new user. Correct
        // currency, wrong answer to "where do you work", and not the default that was asked for.
        public const string DefaultCode = "USD";

        public const string DefaultCountry = "United States";

        // ###########################################################################################
        // One selectable row: the country the user recognises, and the ISO 4217 code that country
        // uses. DisplayName is what the drop-down shows and what the Configuration tab reads back.
        // ###########################################################################################
        public readonly record struct Option(string Country, string Code)
        {
            public string DisplayName => $"{this.Country} ({this.Code})";
        }

        // ###########################################################################################
        // The supported countries, each with its ISO 4217 currency code.
        //
        // Plain-ASCII English country names, ONE row per country. Two decisions are baked in here:
        //
        // The requested list carried "Netherlands"/"The Netherlands" and "Turkey"/"Turkiye" as
        // separate entries. Two rows resolving to the same code would let a user "change" the
        // currency without anything changing, so one name is kept for each.
        //
        // Names are ASCII: "Turkey", not the accented official form. A drop-down of English country
        // names is easier to scan when no row carries a character the rest do not, and an accented
        // name also sorts to the bottom of the list under an ordinal comparison - a trap this list
        // no longer has to avoid. Reunion was dropped for a related reason: it is an overseas
        // department of France rather than a country, and it uses the euro, so picking it was
        // indistinguishable from picking France.
        //
        // Sorted alphabetically by country in Options below rather than by hand, so adding a country
        // here needs no thought about where it goes - and so the declared order cannot drift from
        // what the drop-down actually shows.
        //
        // Several countries share a code (the euro across fifteen of them, USD across the United
        // States and Puerto Rico) - that is expected, and is why the stored value is the code while
        // the CHOICE is a country.
        // ###########################################################################################
        private static readonly (string Country, string Code)[] CountryCurrencies =
        {
            ("Algeria", "DZD"),
            ("Andorra", "EUR"),
            ("Argentina", "ARS"),
            ("Australia", "AUD"),
            ("Austria", "EUR"),
            ("Belgium", "EUR"),
            ("Brazil", "BRL"),
            ("Canada", "CAD"),
            ("Chile", "CLP"),
            ("China", "CNY"),
            ("Colombia", "COP"),
            ("Croatia", "EUR"),
            ("Czechia", "CZK"),
            ("Denmark", "DKK"),
            ("Finland", "EUR"),
            ("France", "EUR"),
            ("Germany", "EUR"),
            ("Greece", "EUR"),
            ("Guatemala", "GTQ"),
            ("Hong Kong", "HKD"),
            ("Hungary", "HUF"),
            ("India", "INR"),
            ("Indonesia", "IDR"),
            ("Iran", "IRR"),
            ("Iraq", "IQD"),
            ("Ireland", "EUR"),
            ("Italy", "EUR"),
            ("Japan", "JPY"),
            ("Kuwait", "KWD"),
            ("Latvia", "EUR"),
            ("Lithuania", "EUR"),
            ("Luxembourg", "EUR"),
            ("Macao", "MOP"),
            ("Mauritius", "MUR"),
            ("Mexico", "MXN"),
            ("Netherlands", "EUR"),
            ("New Zealand", "NZD"),
            ("Norway", "NOK"),
            ("Paraguay", "PYG"),
            ("Peru", "PEN"),
            ("Poland", "PLN"),
            ("Portugal", "EUR"),
            ("Puerto Rico", "USD"),
            ("Romania", "RON"),
            ("Russia", "RUB"),
            ("Rwanda", "RWF"),
            ("San Marino", "EUR"),
            ("Saudi Arabia", "SAR"),
            ("Senegal", "XOF"),
            ("Serbia", "RSD"),
            ("Singapore", "SGD"),
            ("Slovakia", "EUR"),
            ("Slovenia", "EUR"),
            ("South Africa", "ZAR"),
            ("South Korea", "KRW"),
            ("Spain", "EUR"),
            ("Sweden", "SEK"),
            ("Switzerland", "CHF"),
            ("Thailand", "THB"),
            ("Tunisia", "TND"),
            ("Turkey", "TRY"),
            ("Ukraine", "UAH"),
            ("United Kingdom", "GBP"),
            ("United States", "USD"),
            ("Uzbekistan", "UZS")
        };

        // ###########################################################################################
        // The rows the Configuration drop-down shows, sorted alphabetically by country.
        //
        // Sorted with InvariantCulture rather than Ordinal. Every name here is ASCII, so today the
        // two agree - but the comparer is the one that stays correct if an accented name is ever
        // added back, where an ordinal sort compares code points and files it past "Z", hiding the
        // row at the bottom of a list the user is scanning alphabetically.
        // ###########################################################################################
        public static IReadOnlyList<Option> Options { get; } = CountryCurrencies
            .Select(entry => new Option(entry.Country, entry.Code))
            .OrderBy(option => option.Country, StringComparer.InvariantCulture)
            .ToList();

        // ###########################################################################################
        // The country row a stored CODE should show as selected.
        //
        // A code can belong to several countries (EUR to fifteen of them), so the round trip
        // country -> code -> country is deliberately lossy: someone who picked "Austria" reopens the
        // tab on Andorra, the first euro country alphabetically. That is accepted rather than worked
        // around by storing the country alongside the code - what the app prints is the code, so the
        // two selections mean the same thing, and a second stored field could disagree with the
        // first after a hand edit.
        //
        // The DEFAULT code is the exception, and resolves by COUNTRY: two countries here use USD and
        // Puerto Rico sorts first, so scanning by code alone would open every new user's drop-down
        // on Puerto Rico. The default is a named country (DefaultCountry), not merely whichever row
        // happens to carry its code.
        //
        // Returns the FIRST row for a code that names no country here (a hand-edited settings file,
        // or a country dropped in a later release) rather than null - the drop-down has no empty row
        // to fall back to. NormalizeCode has already turned such a code into DefaultCode, so in
        // practice that arm is reached only if DefaultCountry ever stops naming a country.
        // ###########################################################################################
        public static Option ResolveOption(string? code)
        {
            string normalized = NormalizeCode(code);

            string? preferredCountry = string.Equals(normalized, DefaultCode, StringComparison.Ordinal)
                ? DefaultCountry
                : null;

            Option? firstWithCode = null;

            foreach (var option in Options)
            {
                if (!string.Equals(option.Code, normalized, StringComparison.Ordinal))
                    continue;

                if (preferredCountry != null && string.Equals(option.Country, preferredCountry, StringComparison.Ordinal))
                    return option;

                firstWithCode ??= option;
            }

            return firstWithCode ?? Options[0];
        }

        // ###########################################################################################
        // A stored code cleaned up to the form the table uses: trimmed and upper-cased, or the
        // default when it is blank or names no country here.
        //
        // Upper-casing with InvariantCulture, NOT the current culture: ToUpper() in Turkish maps
        // "i" to a dotted capital, so a user in Turkey would turn a stored "irr" into a string that
        // matches nothing - and Turkey is on this very list.
        // ###########################################################################################
        public static string NormalizeCode(string? code)
        {
            if (string.IsNullOrWhiteSpace(code))
                return DefaultCode;

            string trimmed = code.Trim().ToUpperInvariant();

            return CountryCurrencies.Any(entry => string.Equals(entry.Code, trimmed, StringComparison.Ordinal))
                ? trimmed
                : DefaultCode;
        }

        // ###########################################################################################
        // Which code a STORED work-done row's cost should be printed in: the one recorded on the row
        // itself, or - when the row predates that field - the user's current setting.
        //
        // WHY THE FALLBACK IS THE CURRENT SETTING and not the default code: a blank means the row was
        // written before the currency was recorded per row, so what it was actually entered in is
        // simply unknown. Printing DefaultCode for it would assert USD about money that was very
        // likely never USD; printing the current setting reproduces exactly what the app showed for
        // that row before this field existed, which is the one behaviour that is not a new claim.
        //
        // The one place that fallback is decided, so no caller can pick a different one - the whole
        // point being that two surfaces must never label the same figure differently.
        // ###########################################################################################
        public static string ResolveRecordedCode(string? recordedCode, string? currentSettingCode)
        {
            return string.IsNullOrWhiteSpace(recordedCode)
                ? NormalizeCode(currentSettingCode)
                : NormalizeCode(recordedCode);
        }

        // ###########################################################################################
        // A cost, formatted the way every surface in the app prints one: the same "0.##" invariant
        // number the bare figure always used, now followed by the currency code - "430 DKK".
        //
        // The CODE trails the number rather than a symbol leading it, for two reasons. These
        // currencies share symbols freely (kr is Danish, Norwegian AND Swedish; $ covers five of
        // them), so a symbol would identify the money less precisely than the code does. And the PDF
        // export renders through a font subset with no glyph for most currency symbols - an unmapped
        // one prints as a blank box, which reads as a defect rather than as money.
        //
        // InvariantCulture on purpose, matching hours and every other number the worklog prints: a
        // workbook's JSON is written invariant, and a figure that changed shape between the machine
        // that typed it and the machine that opened it would look like a different figure.
        // ###########################################################################################
        public static string FormatCost(double cost, string? code) =>
            $"{cost.ToString("0.##", CultureInfo.InvariantCulture)} {NormalizeCode(code)}";

        // ###########################################################################################
        // The same cost, but EMPTY when there is none to report - the cost half of
        // WorklogDurationFormatter's own zero rule, and for the same reason: a summary line reading
        // "1 worklog . 0 USD . 1 open" spends a column saying that a figure is absent, which the
        // absence itself says better. Asked for directly, alongside the time.
        //
        // A SEPARATE method rather than a change to FormatCost, which stays the "print a cost"
        // primitive. A field that must show something (an editor row being typed into, a column in
        // a table) needs the zero, and a shared formatter that silently returned nothing for it
        // would be a trap for the next caller rather than a convenience.
        //
        // Rounded before the test, not compared to zero: what decides this is whether the figure
        // would PRINT as zero, and a cost of 0.004 formats as "0" under "0.##". Comparing the raw
        // double would print a bare "0" for it and call that a real cost.
        // ###########################################################################################
        public static string FormatCostOrEmpty(double cost, string? code)
        {
            if (double.IsNaN(cost) || double.IsInfinity(cost) || Math.Round(cost, 2) == 0.0)
            {
                return string.Empty;
            }

            return FormatCost(cost, code);
        }
    }
}
