using System;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for WorklogCurrency - the country/currency list behind the Configuration tab's drop-down,
// and the one place a cost is turned into text.
//
// Three things here are worth more than they look. The drop-down's ORDER is asserted rather than
// eyeballed, because the table is hand-maintained and a row added at the bottom out of alphabetical
// order is invisible in a diff and hard to spot in a 65-row list. Every name is asserted to be
// ASCII, which is a display decision the table cannot enforce on its own. And NormalizeCode's
// fallback is what stands between a hand-edited settings file and a nonsense code printed beside
// every cost in a customer's PDF.
public class WorklogCurrencyTests
{
    // ---------------------------------------------------------------------- the list itself

    // The default the app ships with, and what every existing installation gets on upgrade.
    [Fact]
    public void The_default_currency_is_the_United_States_dollar()
    {
        Assert.Equal("USD", WorklogCurrency.DefaultCode);
        Assert.Equal("United States", WorklogCurrency.ResolveOption(null).Country);
    }

    // The drop-down is sorted by country - the order a user scanning for their own scans in. Worth
    // asserting rather than eyeballing: the table is hand-maintained, and a row appended out of
    // order is invisible in a diff and easy to miss in a list this long.
    [Fact]
    public void The_countries_are_sorted_alphabetically()
    {
        var countries = WorklogCurrency.Options.Select(o => o.Country).ToList();

        Assert.Equal(countries.OrderBy(c => c, StringComparer.InvariantCulture).ToList(), countries);
    }

    // ###########################################################################################
    // Every name is plain ASCII - no accented or otherwise non-English characters anywhere in the
    // list.
    //
    // Asked for directly: "Türkiye" was replaced by "Turkey" because a row carrying a character the
    // rest do not is harder to scan and reads as a rendering fault. The table cannot enforce this on
    // its own, so it is asserted - the failure mode is a well-meaning later edit "correcting" a name
    // back to its official accented form.
    //
    // Codes are covered separately by Every_code_is_a_three_letter_uppercase_code; this is the
    // country half.
    // ###########################################################################################
    [Fact]
    public void Every_country_name_is_plain_ascii()
    {
        foreach (var option in WorklogCurrency.Options)
        {
            Assert.All(option.Country, c => Assert.True(
                c <= 127,
                $"\"{option.Country}\" contains the non-ASCII character '{c}'"));
        }
    }

    // Turkey specifically, since it is the name that changed and the one most likely to be
    // "corrected" back to its official accented spelling by a later reader.
    [Fact]
    public void Turkey_is_listed_by_its_plain_english_name()
    {
        var countries = WorklogCurrency.Options.Select(o => o.Country).ToList();

        Assert.Contains("Turkey", countries);
        Assert.DoesNotContain("Türkiye", countries);
    }

    // ###########################################################################################
    // Reunion is NOT offered.
    //
    // It was on the requested list, but it is an overseas department of France rather than a
    // country, and it uses the euro - so the row was a choice that changed nothing a user could
    // see, France being already present with the same code. Pinned so it is not added back as a
    // "missing country" by someone comparing this table against the original request.
    // ###########################################################################################
    [Fact]
    public void Reunion_is_not_offered_because_it_is_a_part_of_France()
    {
        var countries = WorklogCurrency.Options.Select(o => o.Country).ToList();

        Assert.DoesNotContain("Reunion", countries);
        Assert.DoesNotContain("Réunion", countries);

        // France covers it, with the same code the dropped row carried.
        Assert.Equal("EUR", WorklogCurrency.Options.Single(o => o.Country == "France").Code);
    }

    // ###########################################################################################
    // ONE row per country, and one country per name.
    //
    // The requested list carried "Netherlands"/"The Netherlands" and "Turkey"/"Türkiye" as separate
    // entries. Both pairs resolve to the same code, so keeping both would give the user two rows
    // that look like a choice and are not - picking the other one changes nothing they can see.
    // One plain-English name is kept for each.
    // ###########################################################################################
    [Fact]
    public void Each_country_appears_exactly_once()
    {
        var countries = WorklogCurrency.Options.Select(o => o.Country).ToList();

        Assert.Equal(countries.Count, countries.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    [Fact]
    public void The_duplicate_country_names_from_the_request_are_not_both_listed()
    {
        var countries = WorklogCurrency.Options.Select(o => o.Country).ToList();

        Assert.Contains("Netherlands", countries);
        Assert.DoesNotContain("The Netherlands", countries);

        Assert.Contains("Turkey", countries);
        Assert.DoesNotContain("Türkiye", countries);
    }

    // ###########################################################################################
    // Countries sharing a currency is EXPECTED, not a duplicate to clean up - it is precisely why
    // the stored value is the code while the choice is a country. This asserts the euro group is
    // genuinely shared rather than each country having been given a made-up code of its own.
    // ###########################################################################################
    [Theory]
    [InlineData("Denmark", "DKK")]
    [InlineData("Norway", "NOK")]
    [InlineData("Sweden", "SEK")]
    [InlineData("Germany", "EUR")]
    [InlineData("Austria", "EUR")]
    [InlineData("Croatia", "EUR")]      // Euro since 2023, not the kuna
    [InlineData("Slovakia", "EUR")]
    [InlineData("Latvia", "EUR")]
    [InlineData("San Marino", "EUR")]
    [InlineData("United States", "USD")]
    [InlineData("Puerto Rico", "USD")]  // A US territory, so the dollar
    [InlineData("United Kingdom", "GBP")]
    [InlineData("Japan", "JPY")]
    [InlineData("South Korea", "KRW")]
    [InlineData("Switzerland", "CHF")]
    [InlineData("Turkey", "TRY")]
    [InlineData("Mauritius", "MUR")]
    [InlineData("Macao", "MOP")]
    [InlineData("Senegal", "XOF")]      // The West African CFA franc, shared across the union
    [InlineData("Czechia", "CZK")]
    [InlineData("South Africa", "ZAR")]
    public void A_country_maps_to_its_own_currency(string country, string expectedCode)
    {
        var option = WorklogCurrency.Options.Single(o =>
            string.Equals(o.Country, country, StringComparison.Ordinal));

        Assert.Equal(expectedCode, option.Code);
    }

    // Every code is a three-letter uppercase ISO 4217 code - the shape FormatCost prints and the
    // shape NormalizeCode matches against. A lowercase or four-letter entry would be stored fine
    // and then never match on reload, silently reverting the user's choice to USD.
    [Fact]
    public void Every_code_is_a_three_letter_uppercase_code()
    {
        foreach (var option in WorklogCurrency.Options)
        {
            Assert.Equal(3, option.Code.Length);
            Assert.Equal(option.Code.ToUpperInvariant(), option.Code);
            Assert.All(option.Code, c => Assert.True(c >= 'A' && c <= 'Z'));
        }
    }

    // What the drop-down actually shows: the country followed by its code in brackets, so the user
    // can pick by the country they know AND see what will be printed on their documents.
    [Fact]
    public void A_row_shows_the_country_and_its_currency()
    {
        Assert.Equal("Denmark (DKK)", WorklogCurrency.Options.Single(o => o.Country == "Denmark").DisplayName);
    }

    // The exact list the drop-down offers. Written out rather than counted, so a row deleted by
    // accident names itself in the failure instead of just moving a total, and asserted in BOTH
    // directions so an unrequested country cannot appear either.
    //
    // This is the requested list after three deliberate edits, each pinned by its own test above:
    // "The Netherlands" and "Turkey"/"Türkiye" collapsed to one row each, and Reunion dropped as a
    // part of France. Anything else missing here is a mistake, not a decision.
    [Fact]
    public void Every_requested_country_is_offered()
    {
        string[] expected =
        {
            "Algeria", "Andorra", "Argentina", "Australia", "Austria", "Belgium", "Brazil", "Canada",
            "Chile", "China", "Colombia", "Croatia", "Czechia", "Denmark", "Finland", "France",
            "Germany", "Greece", "Guatemala", "Hong Kong", "Hungary", "India", "Indonesia", "Iran",
            "Iraq", "Ireland", "Italy", "Japan", "Kuwait", "Latvia", "Lithuania", "Luxembourg",
            "Macao", "Mauritius", "Mexico", "Netherlands", "New Zealand", "Norway", "Paraguay",
            "Peru", "Poland", "Portugal", "Puerto Rico", "Romania", "Russia", "Rwanda",
            "San Marino", "Saudi Arabia", "Senegal", "Serbia", "Singapore", "Slovakia", "Slovenia",
            "South Africa", "South Korea", "Spain", "Sweden", "Switzerland", "Thailand", "Tunisia",
            "Turkey", "Ukraine", "United Kingdom", "United States", "Uzbekistan"
        };

        var offered = WorklogCurrency.Options.Select(o => o.Country).ToList();

        Assert.Empty(expected.Except(offered, StringComparer.Ordinal));
        Assert.Empty(offered.Except(expected, StringComparer.Ordinal));
    }

    // ---------------------------------------------------------------------- NormalizeCode

    // A code the user actually chose survives untouched.
    [Fact]
    public void A_known_code_is_kept()
    {
        Assert.Equal("DKK", WorklogCurrency.NormalizeCode("DKK"));
    }

    // settings.json is a plain file the user can edit, and the code is read on every cost the app
    // prints. Blank, whitespace, junk and a plausible-but-unlisted code all fall back to the
    // default rather than being printed as themselves beside a figure in an exported PDF.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not a currency")]
    [InlineData("XYZ")]
    [InlineData("ZWL")]     // A real ISO code, but no country here uses it
    public void An_unusable_code_falls_back_to_the_default(string? stored)
    {
        Assert.Equal(WorklogCurrency.DefaultCode, WorklogCurrency.NormalizeCode(stored));
    }

    // Case and surrounding whitespace are tolerated - a hand-edited file is the likely source of
    // both, and rejecting "dkk" would quietly move the user's money to dollars.
    [Theory]
    [InlineData("dkk")]
    [InlineData("Dkk")]
    [InlineData("  DKK  ")]
    [InlineData("\tdkk\n")]
    public void A_known_code_is_recognised_whatever_its_case_or_padding(string stored)
    {
        Assert.Equal("DKK", WorklogCurrency.NormalizeCode(stored));
    }

    // ###########################################################################################
    // Upper-casing is InvariantCulture, not the current culture.
    //
    // ToUpper() under a Turkish culture maps "i" to the dotted capital İ, so a stored "irr" would
    // become "İRR" and match nothing - the user's currency would silently revert to dollars. That
    // is not a theoretical locale either: Turkey is on this very list, so a user in exactly that
    // culture is expected. The test forces the culture rather than trusting the machine's own.
    // ###########################################################################################
    [Fact]
    public void A_lowercase_code_is_recognised_under_a_Turkish_culture()
    {
        var previous = System.Globalization.CultureInfo.CurrentCulture;

        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("tr-TR");

            Assert.Equal("IRR", WorklogCurrency.NormalizeCode("irr"));
            Assert.Equal("TRY", WorklogCurrency.NormalizeCode("try"));
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = previous;
        }
    }

    // ---------------------------------------------------------------------- ResolveOption

    // Which row the drop-down opens on for a stored code.
    [Fact]
    public void A_stored_code_resolves_to_a_country_that_uses_it()
    {
        Assert.Equal("Denmark", WorklogCurrency.ResolveOption("DKK").Country);
        Assert.Equal("Japan", WorklogCurrency.ResolveOption("JPY").Country);
    }

    // ###########################################################################################
    // The country -> code -> country round trip is deliberately LOSSY for a shared currency.
    //
    // Sixteen countries here use the euro, and only the code is stored, so someone who picked
    // "Austria" reopens the Configuration tab on Andorra - the first euro country alphabetically.
    // That is the accepted cost of storing the code rather than the country: the two selections
    // print the same thing, which is what the setting is for. Pinned so the behaviour is a decision
    // on record rather than something a later reader mistakes for a bug and "fixes" by storing the
    // country too - a second field that can disagree with the first after a hand edit.
    // ###########################################################################################
    [Fact]
    public void A_shared_currency_resolves_to_the_first_country_using_it()
    {
        Assert.Equal("Andorra", WorklogCurrency.ResolveOption("EUR").Country);
        Assert.Equal("EUR", WorklogCurrency.ResolveOption("EUR").Code);
    }

    // Never null and never an exception: the drop-down has no empty row to fall back to, and a
    // settings-file typo must not take the Configuration tab down as it is being built.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("XYZ")]
    public void An_unusable_code_resolves_to_the_default_country(string? stored)
    {
        Assert.Equal("United States", WorklogCurrency.ResolveOption(stored).Country);
    }

    // ---------------------------------------------------------------------- FormatCost

    // ###########################################################################################
    // The one format every surface prints a cost with: the app's existing "0.##" number, then the
    // code. The CODE trails the number rather than a symbol leading it - these currencies share
    // symbols freely (kr is Danish, Norwegian and Swedish at once), and the PDF's font subset has
    // no glyph for most of them, which prints as a blank box.
    // ###########################################################################################
    [Theory]
    [InlineData(430.0, "DKK", "430 DKK")]
    [InlineData(0.0, "DKK", "0 DKK")]
    [InlineData(12.5, "EUR", "12.5 EUR")]
    [InlineData(12.567, "USD", "12.57 USD")]   // "0.##" rounds to two places
    [InlineData(1234567.0, "JPY", "1234567 JPY")]
    public void A_cost_is_printed_as_the_number_then_the_code(double cost, string code, string expected)
    {
        Assert.Equal(expected, WorklogCurrency.FormatCost(cost, code));
    }

    // A trailing ".00" is dropped, matching how hours and every other worklog number print - the
    // "0.##" format the entry card, the editor and the summary strip already shared.
    [Fact]
    public void A_whole_cost_prints_without_decimals()
    {
        Assert.Equal("100 SEK", WorklogCurrency.FormatCost(100.0, "SEK"));
    }

    // InvariantCulture, like every other number the worklog writes: entries.json is written
    // invariant, so a figure that grew a comma on a Danish machine and a full stop on an English
    // one would look like two different numbers in one shared workbook.
    [Fact]
    public void A_decimal_cost_prints_invariantly_whatever_the_machines_culture()
    {
        var previous = System.Globalization.CultureInfo.CurrentCulture;

        try
        {
            // da-DK writes decimals with a comma.
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("da-DK");

            Assert.Equal("12.5 DKK", WorklogCurrency.FormatCost(12.5, "DKK"));
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = previous;
        }
    }

    // A cost printed with an unusable code still says something true - the default - rather than
    // printing the junk or dropping the currency and going back to a bare number.
    [Fact]
    public void A_cost_with_an_unusable_code_prints_the_default_currency()
    {
        Assert.Equal("430 USD", WorklogCurrency.FormatCost(430.0, "nonsense"));
        Assert.Equal("430 USD", WorklogCurrency.FormatCost(430.0, null));
    }

    // ---------------------------------------------------------------------------------------------
    // FormatCostOrEmpty - the cost half of the "say nothing rather than say zero" rule, the twin of
    // WorklogDurationFormatter's own. The summary strip, the worklog cards, the editor's work-done
    // rows and the exported PDF all drop the cost when there is none: "1 worklog . 0 USD . 1 open"
    // spends a column reporting the absence of a figure.
    // ---------------------------------------------------------------------------------------------

    [Fact]
    public void A_real_cost_is_formatted_exactly_as_FormatCost_does()
    {
        Assert.Equal(
            WorklogCurrency.FormatCost(430.0, "DKK"),
            WorklogCurrency.FormatCostOrEmpty(430.0, "DKK"));
    }

    // Zero, and anything that would PRINT as zero. The test is on the rounded figure rather than on
    // the raw double: 0.004 formats as "0" under "0.##", and treating that as a real cost would put
    // a bare "0 DKK" back on the line this rule exists to keep clean.
    [Theory]
    [InlineData(0.0)]
    [InlineData(0.004)]
    [InlineData(-0.004)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_cost_that_would_print_as_zero_prints_nothing_at_all(double cost)
    {
        Assert.Equal(string.Empty, WorklogCurrency.FormatCostOrEmpty(cost, "DKK"));
    }

    // The threshold is where "0.##" itself rounds, not somewhere of its own: a cost that prints as
    // a real figure must still print, however small.
    [Fact]
    public void A_small_but_printable_cost_still_prints()
    {
        Assert.Equal("0.01 DKK", WorklogCurrency.FormatCostOrEmpty(0.01, "DKK"));
    }

    // FormatCost itself is UNCHANGED and still prints the zero - it is the "print a cost" primitive,
    // and a field being typed into needs to show one. Only the callers that render a summary line
    // opt into dropping it.
    [Fact]
    public void FormatCost_itself_still_prints_a_zero_cost()
    {
        Assert.Equal("0 DKK", WorklogCurrency.FormatCost(0.0, "DKK"));
    }

    // ###########################################################################################
    // WHICH CODE A STORED COST IS PRINTED IN. The currency is one app-wide preference and a cost is
    // a bare number, so reading the setting wherever a cost is shown silently relabels every
    // historical figure the moment that preference changes - a repair costed in DKK starts reading
    // "430 GBP", most damagingly in a re-exported PDF that goes to someone who cannot ask.
    //
    // So the code is recorded ON the work-done row when the figure is typed, and this is the one
    // place the record-or-fallback rule is decided.
    // ###########################################################################################
    [Fact]
    public void A_recorded_currency_wins_over_the_current_setting()
    {
        Assert.Equal("DKK", WorklogCurrency.ResolveRecordedCode("DKK", "GBP"));
    }

    // Blank means the row predates the field, so what it was entered in is simply unknown. Falling
    // back to the CURRENT setting reproduces exactly what the app showed for that row before this
    // existed - the one answer that is not a fresh claim about old money. DefaultCode would assert
    // USD about a figure that was very likely never USD.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_row_with_no_recorded_currency_falls_back_to_the_current_setting(string? recorded)
    {
        Assert.Equal("GBP", WorklogCurrency.ResolveRecordedCode(recorded, "GBP"));
    }

    // Both halves go through NormalizeCode, so a hand-edited entries.json carrying a code this app
    // does not know cannot put an unrecognised string beside a figure in an exported document.
    [Fact]
    public void An_unknown_recorded_currency_normalises_rather_than_printing_as_typed()
    {
        Assert.Equal(WorklogCurrency.DefaultCode, WorklogCurrency.ResolveRecordedCode("ZZZ", "GBP"));
        Assert.Equal("DKK", WorklogCurrency.ResolveRecordedCode("  dkk  ", "GBP"));
    }
}
