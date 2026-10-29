using System.Text;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// "Try it" on the Oscilloscope tab's model editor (ScopeCommandTrial): one command, as the box
// holds it, sent to a scope - here a fake one that answers from a table and records what it was
// sent - and its answer said in plain words. What goes down the wire to someone's scope is the
// point: a Set command must set the value the scope already HAS, so trying it changes nothing.
public sealed class ScopeCommandTrialTests
{
    private sealed class FakeScope : IScopeClient
    {
        private readonly Dictionary<string, string> thisAnswers;

        public FakeScope(Dictionary<string, string>? answers = null) => this.thisAnswers = answers ?? new();

        public List<string> Sent { get; } = new();

        public byte[] Picture { get; set; } = [];

        public async Task SendAsync(string commandText, CancellationToken cancellationToken)
        {
            this.Sent.Add(commandText);
            await Task.Yield();
        }

        public async Task<string> QueryLineAsync(string commandText, CancellationToken cancellationToken)
        {
            this.Sent.Add(commandText);
            await Task.Yield();
            return this.thisAnswers.TryGetValue(commandText, out string? answer) ? answer : "0,\"No error\"";
        }

        public async Task<byte[]> QueryBinaryBlockAsync(string commandText, CancellationToken cancellationToken)
        {
            this.Sent.Add(commandText);
            await Task.Yield();
            return this.Picture;
        }
    }

    private static MainDocument.OscilloscopeRow Model() => new()
    {
        Brand = "Voltcraft",
        SeriesOrModel = "DOV702",
        Identify = "*IDN?",
        DrainErrorQueue = ":SYST:ERR?",
        OperationComplete = "*OPC?",
        Stop = ":STOP",
        QueryTriggerLevel = ":TRIG:EDGE:LEV?",
        SetTriggerLevel = ":TRIG:EDGE:LEV {0}",
        QueryTimeDiv = ":TIM:SCAL?",
        SetTimeDiv = ":TIM:SCAL",
        DumpImage = ":DISP:DATA?",
    };

    private static Task<ScopeTrialOutcome> Try(FakeScope scope, MainDocument.OscilloscopeRow model, ScopeCommand command, List<string>? log = null) =>
        ScopeCommandTrial.TryAsync(scope, model, command, line => log?.Add(line), CancellationToken.None);

    [Fact]
    public async Task A_query_says_what_the_scope_answered_with_the_serial_masked()
    {
        var scope = new FakeScope(new() { ["*IDN?"] = "OWON,ADS822A,SN1234567,V1.2" });
        var log = new List<string>();

        ScopeTrialOutcome outcome = await Try(scope, Model(), ScopeCommand.Identify, log);

        Assert.True(outcome.Worked);
        Assert.Equal(["*IDN?"], scope.Sent);
        Assert.Contains("ADS822A", outcome.Reply);
        Assert.DoesNotContain("SN1234567", outcome.Reply);
        Assert.DoesNotContain(log, line => line.Contains("SN1234567"));
        Assert.Contains("SCPI >> *IDN?", log);
    }

    [Fact]
    public async Task An_empty_command_is_not_sent()
    {
        var scope = new FakeScope();
        MainDocument.OscilloscopeRow model = Model();
        model.Run = "  ";

        ScopeTrialOutcome outcome = await Try(scope, model, ScopeCommand.Run);

        Assert.False(outcome.Worked);
        Assert.Equal(ScopeCommandTrial.NoCommand, outcome.Reply);
        Assert.Empty(scope.Sent);
    }

    // The keyboard steps need a number: an answer CRT cannot read as one is a command that fails.
    [Fact]
    public async Task A_query_for_a_number_says_the_number_or_that_it_is_none()
    {
        ScopeTrialOutcome read = await Try(new FakeScope(new() { [":TIM:SCAL?"] = "1.0e-3" }), Model(), ScopeCommand.QueryTimeDiv);
        ScopeTrialOutcome garbled = await Try(new FakeScope(new() { [":TIM:SCAL?"] = "1 ms" }), Model(), ScopeCommand.QueryTimeDiv);

        Assert.True(read.Worked);
        Assert.Contains("1ms", read.Reply, StringComparison.OrdinalIgnoreCase);
        Assert.False(garbled.Worked);
        Assert.Contains("no number", garbled.Reply);
    }

    // A command that answers nothing: sent, then the model's own Finished and Last error asked,
    // as the palettes do - an error the scope kept means it did not take it.
    [Fact]
    public async Task A_command_that_answers_nothing_is_followed_by_finished_and_last_error()
    {
        var scope = new FakeScope(new() { ["*OPC?"] = "1" });

        ScopeTrialOutcome outcome = await Try(scope, Model(), ScopeCommand.Stop);

        Assert.True(outcome.Worked);
        Assert.Equal([":STOP", "*OPC?", ":SYST:ERR?"], scope.Sent);
        Assert.Contains("Finished: 1", outcome.Reply);
    }

    [Fact]
    public async Task An_error_the_scope_keeps_after_a_command_means_it_did_not_take_it()
    {
        var scope = new FakeScope(new() { [":SYST:ERR?"] = "-113,\"Undefined header\"" });

        ScopeTrialOutcome outcome = await Try(scope, Model(), ScopeCommand.Stop);

        Assert.False(outcome.Worked);
        Assert.Contains("Undefined header", outcome.Reply);
    }

    [Theory]
    [InlineData("0,\"No error\"", false)]
    [InlineData("+0,\"No error\"", false)]
    [InlineData("-113,\"Undefined header\"", true)]
    [InlineData("+222,\"Data out of range\"", true)]
    [InlineData("No error", false)]
    [InlineData("", false)]
    public void An_answer_of_last_error_names_an_error_by_its_number(string reply, bool isAnError)
    {
        Assert.Equal(isAnError, ScopeCommandTrial.IsAnError(reply));
    }

    // Trying a Set command changes nothing on the scope: the value set is the one read first.
    [Fact]
    public async Task A_set_command_sets_the_value_the_scope_already_has()
    {
        var scope = new FakeScope(new() { [":TRIG:EDGE:LEV?"] = "1.5" });

        ScopeTrialOutcome outcome = await Try(scope, Model(), ScopeCommand.SetTriggerLevel);

        Assert.True(outcome.Worked);
        Assert.Equal([":TRIG:EDGE:LEV?", ":TRIG:EDGE:LEV 1.5", "*OPC?", ":SYST:ERR?"], scope.Sent);
        Assert.Contains("1.5V", outcome.Reply);
    }

    // Without "{0}", the value goes at the end - the palettes' own rule.
    [Fact]
    public async Task A_set_command_without_a_place_for_the_value_gets_it_at_the_end()
    {
        var scope = new FakeScope(new() { [":TIM:SCAL?"] = "0.001" });

        await Try(scope, Model(), ScopeCommand.SetTimeDiv);

        Assert.Contains(":TIM:SCAL 0.001", scope.Sent);
    }

    [Fact]
    public async Task A_set_command_whose_read_command_is_missing_or_unreadable_sends_nothing_set()
    {
        MainDocument.OscilloscopeRow noRead = Model();
        noRead.QueryTriggerLevel = string.Empty;
        var silent = new FakeScope();

        ScopeTrialOutcome missing = await Try(silent, noRead, ScopeCommand.SetTriggerLevel);

        Assert.False(missing.Worked);
        Assert.Contains("Read trigger level", missing.Reply);
        Assert.Empty(silent.Sent);

        var garbled = new FakeScope(new() { [":TRIG:EDGE:LEV?"] = "LEVEL?" });
        ScopeTrialOutcome unreadable = await Try(garbled, Model(), ScopeCommand.SetTriggerLevel);

        Assert.False(unreadable.Worked);
        Assert.Equal([":TRIG:EDGE:LEV?"], garbled.Sent);
    }

    // ------------------------------------------------------------------ the screen picture

    private static byte[] Block(byte[] payload)
    {
        string length = payload.Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return [.. Encoding.ASCII.GetBytes($"#{length.Length}{length}"), .. payload];
    }

    private static byte[] Bmp(int width, int height)
    {
        byte[] bmp = new byte[54];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        BitConverter.GetBytes(54).CopyTo(bmp, 2);
        BitConverter.GetBytes(40).CopyTo(bmp, 14);
        BitConverter.GetBytes(width).CopyTo(bmp, 18);
        BitConverter.GetBytes(height).CopyTo(bmp, 22);
        BitConverter.GetBytes((short)1).CopyTo(bmp, 26);
        BitConverter.GetBytes((short)24).CopyTo(bmp, 28);
        return bmp;
    }

    [Fact]
    public async Task The_screen_picture_says_what_came_back()
    {
        var scope = new FakeScope { Picture = Block(Bmp(800, 480)) };

        ScopeTrialOutcome outcome = await Try(scope, Model(), ScopeCommand.DumpImage);

        Assert.True(outcome.Worked);
        Assert.Contains("BMP", outcome.Reply);
        Assert.Contains("800x480", outcome.Reply);
    }

    [Fact]
    public void A_png_or_jpeg_picture_is_one_too_and_anything_else_is_not()
    {
        byte[] png = [0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R', 0, 0, 3, 32, 0, 0, 1, 224];

        Assert.Contains("PNG picture of 800x480", ScopeCommandTrial.DescribePicture(Block(png)).Reply);
        Assert.True(ScopeCommandTrial.DescribePicture(Block([0xFF, 0xD8, 0xFF, 0xE0])).Worked);

        ScopeTrialOutcome text = ScopeCommandTrial.DescribePicture(Block(Encoding.ASCII.GetBytes("not a picture")));
        Assert.False(text.Worked);
        Assert.Contains("not as a picture", text.Reply);
    }

    // A JPEG's size is said too - read by CRT's one reader of picture headers, past the segments
    // before its frame header (code review, 2026-10-29: read here by hand, it was never said).
    [Fact]
    public void A_jpeg_picture_says_its_size_from_its_frame_header()
    {
        byte[] jpeg =
        [
            0xFF, 0xD8,
            0xFF, 0xE0, 0x00, 0x06, (byte)'J', (byte)'F', (byte)'I', (byte)'F',
            0xFF, 0xC0, 0x00, 0x11, 0x08, 0x01, 0xE0, 0x03, 0x20, 0x03, 0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
            0xFF, 0xD9,
        ];

        Assert.Contains("JPEG picture of 800x480", ScopeCommandTrial.DescribePicture(Block(jpeg)).Reply);
    }

    // A PNG whose signature was damaged in transit is no PNG - its size would be whatever bytes sit
    // where the header should be (WorkbookPdfExporter's own reason to check all eight bytes).
    [Fact]
    public void A_png_with_a_damaged_signature_is_no_picture()
    {
        byte[] damaged = [0x89, (byte)'P', (byte)'N', (byte)'G', 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R', 0, 0, 3, 32, 0, 0, 1, 224, 0];

        Assert.False(ScopeCommandTrial.DescribePicture(Block(damaged)).Worked);
    }
}
