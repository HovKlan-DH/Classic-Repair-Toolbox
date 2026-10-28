using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SimilarNames - "Did you mean Commodore?" (CRT 4.0 step 3, slice 3e - the plan: "Did you mean
// Commodore? when a typed name is close to an existing one"). A new board's names are free text, and
// one written another way than CRT has it makes a second manufacturer or machine beside the real one.
// The names here are the live data's: its manufacturers, its machines and their folders.
// ###########################################################################################
public sealed class SimilarNamesTests
{
    private static readonly KnownName[] Manufacturers =
        [new("Amstrad", "Amstrad"), new("Commodore", "Commodore"), new("ZX Spectrum", "ZX Spectrum")];

    // Each machine by its name and by its folder, as the new-board form hands them over.
    private static readonly KnownName[] Hardware =
    [
        new("Commodore 64", "Commodore 64"), new("C64", "Commodore 64"),
        new("Commodore 128", "Commodore 128"), new("C128", "Commodore 128"),
        new("Commodore 16", "Commodore 16"), new("C16", "Commodore 16"),
        new("Commodore Plus/4", "Commodore Plus/4"), new("Plus4", "Commodore Plus/4"),
        new("Commodore VIC-20", "Commodore VIC-20"), new("VIC-20", "Commodore VIC-20"),
        new("Amstrad CPC 664", "Amstrad CPC 664"), new("CPC 664", "Amstrad CPC 664"),
    ];

    [Theory]
    [InlineData("Comodore", "Commodore")]
    [InlineData("Commodre", "Commodore")]
    [InlineData("Commodroe", "Commodore")]
    [InlineData("commodore", "Commodore")]
    [InlineData(" COMMODORE ", "Commodore")]
    [InlineData("Amstard", "Amstrad")]
    [InlineData("ZX-Spectrum", "ZX Spectrum")]
    [InlineData("zxspectrum", "ZX Spectrum")]
    public void A_manufacturer_a_slip_away_is_offered(string typed, string offered)
    {
        Assert.Equal(offered, SimilarNames.DidYouMean(typed, Manufacturers));
    }

    // Written exactly as CRT has it, or nothing like any of them: nothing to ask.
    [Theory]
    [InlineData("Commodore")]
    [InlineData("Sinclair")]
    [InlineData("Atari")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("-")]
    [InlineData(null)]
    public void A_manufacturer_as_CRT_has_it_or_a_new_one_asks_nothing(string? typed)
    {
        Assert.Null(SimilarNames.DidYouMean(typed, Manufacturers));
    }

    // A machine typed as its FOLDER - the new-board form's old example "C64" - is the machine the
    // drop-down list names otherwise: offered by that name, so the board is listed with its siblings.
    [Theory]
    [InlineData("C64", "Commodore 64")]
    [InlineData("c-64", "Commodore 64")]
    [InlineData("C128", "Commodore 128")]
    [InlineData("VIC 20", "Commodore VIC-20")]
    [InlineData("Plus/4", "Commodore Plus/4")]
    [InlineData("Commodore64", "Commodore 64")]
    [InlineData("Comodore 64", "Commodore 64")]
    [InlineData("Commodore VIC20", "Commodore VIC-20")]
    [InlineData("Amstrad CPC664", "Amstrad CPC 664")]
    public void A_machine_written_another_way_is_offered_by_its_listed_name(string typed, string offered)
    {
        Assert.Equal(offered, SimilarNames.DidYouMean(typed, Hardware));
    }

    // A slip never changes a number - those are other machines - and three letters are never a slip.
    [Theory]
    [InlineData("Commodore 65")]
    [InlineData("Commodore 116")]
    [InlineData("C65")]
    [InlineData("C61")]
    [InlineData("Amstrad CPC 464")]
    [InlineData("Commodore 64")]
    [InlineData("Amiga 500")]
    public void Another_number_is_another_machine(string typed)
    {
        Assert.Null(SimilarNames.DidYouMean(typed, Hardware));
    }

    // ###########################################################################################
    // A "+" and a model letter after the number are the machine too (code review, 2026-10-28): with
    // only letters and digits compared, "Amiga 500+" was "Amiga 500", and "C128D" a slip from "C128".
    // ###########################################################################################
    [Theory]
    [InlineData("Amiga 500+")]
    [InlineData("C128D")]
    [InlineData("Commodore 128D")]
    [InlineData("Spectrum +2A")]
    [InlineData("Commodore 64C")]
    public void A_plus_or_a_model_letter_after_the_number_is_another_machine(string typed)
    {
        KnownName[] known =
        [
            new("Amiga 500", "Amiga 500"), new("Commodore 128", "Commodore 128"), new("C128", "Commodore 128"),
            new("Spectrum +2", "Spectrum +2"), new("Commodore 64", "Commodore 64"),
        ];

        Assert.Null(SimilarNames.DidYouMean(typed, known));
    }

    // The same machine written another way still is - its "+" and its letters kept.
    [Theory]
    [InlineData("amiga 500+", "Amiga 500+")]
    [InlineData("Spectum +2A", "Spectrum +2A")]
    [InlineData("c128d", "Commodore 128D")]
    public void The_same_machine_with_its_plus_or_model_letter_is_still_offered(string typed, string offered)
    {
        KnownName[] known = [new("Amiga 500+", "Amiga 500+"), new("Spectrum +2A", "Spectrum +2A"), new("C128D", "Commodore 128D")];

        Assert.Equal(offered, SimilarNames.DidYouMean(typed, known));
    }

    // Two names equally close: a guess would be wrong half the time - nothing is offered.
    [Fact]
    public void Two_names_equally_close_offer_neither()
    {
        KnownName[] known = [new("Commodore", "Commodore"), new("Commodores", "Commodores")];

        Assert.Null(SimilarNames.DidYouMean("Commodorse", known));

        // One name reached two ways is one name.
        Assert.Equal("Commodore 64", SimilarNames.DidYouMean("c 64", [new("C64", "Commodore 64"), new("Commodore 64", "Commodore 64")]));
    }

    // A name CRT has, typed exactly, is never offered something else - though another way of
    // writing another name says it.
    [Fact]
    public void A_name_CRT_has_typed_exactly_asks_nothing_whatever_else_it_resembles()
    {
        KnownName[] known = [new("C64", "Commodore 64"), new("C64", "C64")];

        Assert.Null(SimilarNames.DidYouMean("C64", known));
    }

    [Fact]
    public void The_question_and_its_button_are_plain_words()
    {
        Assert.Equal("Did you mean Commodore?", SimilarNames.Question("Commodore"));
        Assert.Equal("Use Commodore", SimilarNames.UseButton("Commodore"));
    }
}
