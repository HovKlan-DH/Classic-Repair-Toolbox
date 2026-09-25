using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Test seams for the oscilloscope command sequencing.
    //
    // Same convention as the two TabSchematics seam files. What is worth testing in this tab is
    // the SEQUENCING - which SCPI commands a palette sends, in what order, which of them expect a
    // response, and what the tab does with the values that come back. That was untestable because
    // every path took a concrete ScopeScpiClient, which opens a TCP socket.
    //
    // ScopeScpiClient now implements IScopeClient and the tab's methods take the interface, so a
    // fake can answer with canned responses and no scope has to be on the network. The real client
    // stays deliberately uncovered - it is an I/O boundary, and .claude/CLAUDE.md lists it as such.
    //
    // AppendOutputLine, which ExecutePaletteAsync calls throughout, only appends to a buffered list
    // and defers the actual TextBox write to the dispatcher, so it is safe to run headlessly and
    // the log it produces is itself assertable - via StartRecordingOutputLinesForTests and
    // RecordedOutputLinesForTests, NOT by reading the flush queue, which drains on a timer.
    //
    // Part of the TabOscilloscope partial class.
    // ###########################################################################################
    public partial class TabOscilloscope
    {
        // ###########################################################################################
        // Runs one command palette against the supplied client - the real ExecutePaletteAsync, with
        // only the client swapped for a fake.
        // ###########################################################################################
        internal Task ExecutePaletteForTestsAsync(
            IScopeClient scopeClient,
            OscilloscopeEntry oscilloscope,
            ScopeCommandPalette palette,
            CancellationToken cancellationToken = default) =>
            this.ExecutePaletteAsync(scopeClient, oscilloscope, palette, cancellationToken);

        // Every line AppendOutputLine produced since recording was switched on, in order, and never
        // drained.
        //
        // This exists rather than reading thisPendingOutputLines directly because that buffer is the
        // UI's FLUSH QUEUE: FlushPendingOutputLinesAsync clears it 40ms after the first line lands.
        // A test asserting on it is therefore racing a timer it does not control - the palette run
        // it awaits pumps the dispatcher, so on a loaded runner the flush can fire mid-run and the
        // test reads an empty list. Recording into a list nothing clears removes the race entirely.
        //
        // Null unless StartRecordingOutputLinesForTests has been called, so the shipping app
        // allocates nothing and keeps no unbounded log.
        private List<string>? thisRecordedOutputLinesForTests;

        // ###########################################################################################
        // Starts capturing output lines for assertion, discarding anything recorded before now.
        // ###########################################################################################
        internal void StartRecordingOutputLinesForTests()
        {
            lock (this.thisPendingOutputLinesLock)
            {
                this.thisRecordedOutputLinesForTests = new List<string>();
            }
        }

        // The debug/info log the palette run produced, in order. Read under the same lock the
        // producer uses, since ExecutePaletteAsync can append from a background continuation.
        internal IReadOnlyList<string> RecordedOutputLinesForTests
        {
            get
            {
                lock (this.thisPendingOutputLinesLock)
                {
                    return this.thisRecordedOutputLinesForTests is null
                        ? new List<string>()
                        : new List<string>(this.thisRecordedOutputLinesForTests);
                }
            }
        }

        // The three values the tab caches from query responses, and which its Set* commands then
        // send back. Null until a query has populated them.
        internal double? LastTriggerLevelVoltsForTests => this.thisLastTriggerLevelVolts;

        internal double? LastTimeDivSecondsForTests => this.thisLastTimeDivSeconds;

        internal double? LastVoltsDivVoltsForTests => this.thisLastVoltsDivVolts;

        internal void SetLastTriggerLevelVoltsForTests(double? volts) =>
            this.thisLastTriggerLevelVolts = volts;

        internal void SetLastTimeDivSecondsForTests(double? seconds) =>
            this.thisLastTimeDivSeconds = seconds;

        internal void SetLastVoltsDivVoltsForTests(double? volts) =>
            this.thisLastVoltsDivVolts = volts;
    }
}
