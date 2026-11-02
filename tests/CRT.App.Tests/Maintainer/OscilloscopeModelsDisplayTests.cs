using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests.Maintainer;

// The words of Account > "Oscilloscope models" (CRT 4.0 step 5): each value in the Oscilloscope
// tab's own words and order, an update's changes marked against the published model, the sender
// named in bold as everywhere on the Maintainer tab.
public sealed class OscilloscopeModelsDisplayTests
{
    private static OscilloscopeModelQueueEntry Entry(
        IReadOnlyDictionary<string, string>? published = null,
        string name = "Anna",
        string address = "anna@example.com",
        bool passed = true) => new(
        41,
        new DateTimeOffset(2026, 10, 29, 12, 0, 0, TimeSpan.Zero),
        "Rigol",
        "DS1000Z",
        new Dictionary<string, string> { ["identify"] = "*IDN?", ["dumpImage"] = ":DISP:DATA? PNG" },
        name,
        address,
        "  ",
        passed,
        "CRT 4.0.0",
        published);

    // What decides whether the cards are made again: anything a card says, nothing else.
    [Fact]
    public void Two_reads_are_the_same_only_when_every_entry_says_the_same()
    {
        OscilloscopeModelQueueEntry entry = Entry();

        Assert.True(OscilloscopeModelsDisplay.SameModels([entry], [Entry()]));
        Assert.True(OscilloscopeModelsDisplay.SameModels([], []));

        Assert.False(OscilloscopeModelsDisplay.SameModels(null, []));
        Assert.False(OscilloscopeModelsDisplay.SameModels([entry], []));
        Assert.False(OscilloscopeModelsDisplay.SameModels([entry], [entry with { Id = 42 }]));
        Assert.False(OscilloscopeModelsDisplay.SameModels([entry], [entry with { InBeta = true }]));
        Assert.False(OscilloscopeModelsDisplay.SameModels([entry], [entry with { Values = new Dictionary<string, string> { ["identify"] = "*IDN? " } }]));
        Assert.False(OscilloscopeModelsDisplay.SameModels([entry], [Entry(new Dictionary<string, string> { ["identify"] = "*IDN?" })]));
    }

    [Fact]
    public void Every_value_is_listed_in_the_editor_s_words_and_order()
    {
        IReadOnlyList<ScopeModelValueRow> rows = OscilloscopeModelsDisplay.Rows(Entry());

        Assert.Equal(ScopeModelFields.All.Select(field => field.Label), rows.Select(row => row.Label));
        Assert.Equal("*IDN?", rows.Single(row => row.Label == "Identify").Sent);

        // A value not sent is shown empty - an empty command is something to see.
        Assert.Equal(string.Empty, rows.Single(row => row.Label == "Stop").Sent);
        Assert.All(rows, row => Assert.Null(row.Published));
        Assert.DoesNotContain(rows, row => row.Differs);
        Assert.Equal(OscilloscopeModelsDisplay.NewModel, OscilloscopeModelsDisplay.Kind(Entry()));
    }

    [Fact]
    public void An_update_marks_what_it_changes_against_the_published_model()
    {
        OscilloscopeModelQueueEntry update = Entry(published: new Dictionary<string, string>
        {
            ["identify"] = " *IDN? ",
            ["dumpImage"] = ":DISP:DATA?",
            ["stop"] = ":STOP",
        });

        IReadOnlyList<ScopeModelValueRow> rows = OscilloscopeModelsDisplay.Rows(update);

        Assert.False(rows.Single(row => row.Label == "Identify").Differs);
        Assert.True(rows.Single(row => row.Label == "Screen picture").Differs);
        Assert.True(rows.Single(row => row.Label == "Stop").Differs);
        Assert.Equal(":STOP", rows.Single(row => row.Label == "Stop").Published);
        Assert.Equal(OscilloscopeModelsDisplay.Update, OscilloscopeModelsDisplay.Kind(update));
    }

    // A model a publish put into BETA only is said to be so, new or an update alike - what to do
    // about it, not what it is, is what the administrator needs then.
    [Fact]
    public void A_model_in_BETA_is_said_to_be_there_and_how_to_finish_it()
    {
        Assert.Equal(OscilloscopeModelsDisplay.InBeta, OscilloscopeModelsDisplay.Kind(Entry() with { InBeta = true }));
        Assert.Equal(OscilloscopeModelsDisplay.InBeta, OscilloscopeModelsDisplay.Kind(Entry(published: new Dictionary<string, string>()) with { InBeta = true }));
        Assert.Contains("Publish again", OscilloscopeModelsDisplay.InBeta);
    }

    [Fact]
    public void The_sender_is_named_in_bold_with_the_address_the_day_and_the_CRT()
    {
        IReadOnlyList<ReviewNoteRun> runs = OscilloscopeModelsDisplay.SentBy(Entry());

        Assert.Contains(runs, run => run is { Text: "Anna", IsCount: true });
        Assert.StartsWith("Sent by Anna (anna@example.com) on 2026-10-2", PersonRuns.Text(runs));
        Assert.EndsWith(", with CRT 4.0.0", PersonRuns.Text(runs));

        Assert.Contains("somebody who left no address", PersonRuns.Text(OscilloscopeModelsDisplay.SentBy(Entry(name: "", address: ""))));
    }

    [Fact]
    public void The_suite_line_and_an_empty_message_are_said()
    {
        Assert.Equal(OscilloscopeModelsDisplay.SuitePassed, OscilloscopeModelsDisplay.SuiteLine(Entry()));
        Assert.Equal(OscilloscopeModelsDisplay.SuiteNotPassed, OscilloscopeModelsDisplay.SuiteLine(Entry(passed: false)));
        Assert.Equal(OscilloscopeModelsDisplay.NoMessage, OscilloscopeModelsDisplay.Message(Entry()));
    }

    // ###########################################################################################
    // *** ANYBODY CAN TYPE ANY NAME (code review, 2026-11-02 - Security). *** Sending a model needs no
    // account, so "Sent by Dennis" could be anybody, and Publish pushes the sender's commands to every
    // CRT. Each card says which it is, in the Contributor view's words for a board's submission: a
    // verified account, or a name and address nothing checked. The answer of a server that does not
    // say reads as typed in - the careful side.
    // ###########################################################################################
    [Fact]
    public void A_card_says_whether_the_sender_is_a_verified_account_or_typed_in()
    {
        Assert.Equal(
            "Sent without a verified account - the name and address were typed in when sending, and nothing checked either.",
            OscilloscopeModelsDisplay.SenderLine(Entry()));
        Assert.Equal(
            "Sent with an account - the address is verified.",
            OscilloscopeModelsDisplay.SenderLine(Entry() with { SignedIn = true }));
    }

    // The suite line is the sender's CRT's word for it, which nothing here checked - never CRT's own
    // statement that it went through.
    [Fact]
    public void The_suite_line_says_it_is_what_the_senders_CRT_reported()
    {
        Assert.Equal("The sender's CRT reported that its full test suite went through on it - nothing here checked that.", OscilloscopeModelsDisplay.SuiteLine(Entry()));
        Assert.Equal("The sender's CRT did not report that its full test suite went through on it as it was sent.", OscilloscopeModelsDisplay.SuiteLine(Entry(passed: false)));
    }
}
