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
}
