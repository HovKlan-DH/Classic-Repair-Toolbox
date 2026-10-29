using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace Handlers.MaintainerHandling
{
    // One value of a model sent in, beside the published model's of its name (null for a new model).
    public sealed record ScopeModelValueRow(string Label, string Sent, string? Published, bool Differs);

    // ###########################################################################################
    // THE WORDS OF ACCOUNT > "OSCILLOSCOPE MODELS" (CRT 4.0 step 5): the models sent in from CRT's
    // Oscilloscope tab, each with who sent it, whether CRT's full test suite went through on it, the
    // contributor's message and every value - by the labels the Oscilloscope tab's editor shows
    // (ScopeModelFields), beside the published model's when it replaces one, the changes marked.
    // ###########################################################################################
    public static class OscilloscopeModelsDisplay
    {
        public const string Heading = "Oscilloscope models";

        public const string Explanation =
            "Models people made on CRT's Oscilloscope tab and sent in. Publish puts a model into the BETA and the stable " +
            "data at once - in place of the published model of its name - and everybody gets it with CRT's next data check. " +
            "Reject needs a reason, which the contributor is mailed.";

        public const string Empty = "No oscilloscope model is waiting.";

        public const string NewModel = "A new model.";

        public const string Update = "Replaces the published model of this name - what it changes is marked.";

        // A publish wrote it into BETA and the stable source could not take it (the server's
        // OscilloscopeModelStates.InBeta) - it cannot be rejected any more.
        public const string InBeta =
            "In BETA already - the stable source could not take it when it was published. Put that right, then press " +
            "Publish again to finish it. It cannot be rejected now.";

        public const string SuitePassed = "CRT's full test suite went through on it.";

        public const string SuiteNotPassed = "CRT's full test suite was not run on it as it was sent, or did not go through.";

        public const string NoMessage = "(no message)";

        public const string ReasonPlaceholder = "Why it will not go in - the contributor is mailed this";

        public const string NoReason = "Write why it will not go in first - the contributor is mailed it.";

        public const string ValueColumn = "Sent";

        public const string PublishedColumn = "Published now";

        public static string Title(OscilloscopeModelQueueEntry entry) => $"{entry.Vendor.Trim()} {entry.SeriesOrModel.Trim()}";

        public static string Kind(OscilloscopeModelQueueEntry entry) =>
            entry.InBeta ? OscilloscopeModelsDisplay.InBeta
            : entry.PublishedValues is null ? OscilloscopeModelsDisplay.NewModel
            : OscilloscopeModelsDisplay.Update;

        // "Sent by **Anna** (anna@example.com) on 2026-10-29, with CRT 4.0.0-alpha.3".
        public static IReadOnlyList<ReviewNoteRun> SentBy(OscilloscopeModelQueueEntry entry)
        {
            var runs = new List<ReviewNoteRun> { PersonRuns.Plain("Sent by ") };

            runs.AddRange(PersonRuns.NameAndAddress(entry.ContactName, entry.ContactEmail, "somebody who left no address"));
            runs.Add(PersonRuns.Plain($" on {entry.SubmittedUtc.ToLocalTime().ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}"));

            if (!string.IsNullOrWhiteSpace(entry.ClientVersion))
                runs.Add(PersonRuns.Plain($", with {entry.ClientVersion.Trim()}"));

            return runs;
        }

        public static string SuiteLine(OscilloscopeModelQueueEntry entry) =>
            entry.FullTestSuitePassed ? OscilloscopeModelsDisplay.SuitePassed : OscilloscopeModelsDisplay.SuiteNotPassed;

        public static string Message(OscilloscopeModelQueueEntry entry) =>
            string.IsNullOrWhiteSpace(entry.Message) ? OscilloscopeModelsDisplay.NoMessage : entry.Message.Trim();

        // ###########################################################################################
        // Every value, in the editor's order and words - an empty one too, since an empty command is
        // something the administrator should see. Differs: an update changes it (spaces around aside).
        // ###########################################################################################
        public static IReadOnlyList<ScopeModelValueRow> Rows(OscilloscopeModelQueueEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            return ScopeModelFields.All
                .Select(field =>
                {
                    string sent = entry.Values.TryGetValue(field.Key, out string? value) ? value ?? string.Empty : string.Empty;
                    string? published = entry.PublishedValues is null
                        ? null
                        : entry.PublishedValues.TryGetValue(field.Key, out string? had) ? had ?? string.Empty : string.Empty;

                    return new ScopeModelValueRow(
                        field.Label,
                        sent,
                        published,
                        published is not null && !string.Equals(sent.Trim(), published.Trim(), StringComparison.Ordinal));
                })
                .ToList();
        }
    }
}
