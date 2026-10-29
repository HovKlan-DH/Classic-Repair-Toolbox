using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // AN OSCILLOSCOPE MODEL OF THE USER'S OWN - the rules the Oscilloscope tab's model editor keeps
    // (CRT 4.0 step 5 - the plan: "On the Oscilloscope tab, start from a similar model and edit each
    // command with Try it, then run the full test suite. Your own models are kept outside the synced
    // data"). The models live in the user's own file (OwnOscilloscopeFile), listed after the
    // published ones.
    //
    // *** A MODEL IS KNOWN BY ITS VENDOR AND SERIES OR MODEL, capitals and spaces around ignored ***
    // - the drop-downs list each name once and pick a model by it. So no two models in the list,
    // published or own, may share a name (NameProblem): the second would never be reachable.
    //
    // *** A MODEL PUBLISHED LATER UNDER AN OWN MODEL'S NAME TAKES IT OVER, OR NAMES IT APART ***
    // (Reconciled, at every launch): exactly as published - the user's model sent in and published,
    // say - the own copy is no longer needed and goes; with any value different, the user's copy
    // is kept beside it as "<model> (my version)", the rule OwnOscilloscopeModels keeps for what 3.x
    // left. Nothing a user made is lost, and nothing is listed twice.
    // ###########################################################################################
    public static class OscilloscopeModelRules
    {
        public const string NoVendor = "Type the vendor.";

        public const string NoModel = "Type the series or model.";

        public static string Taken(string vendor, string model) =>
            $"There is already a {vendor.Trim()} {model.Trim()} in the list - give yours a name of its own.";

        // The longest vendor or series or model, value, and message a model sent in may carry.
        public const int NameLimit = 100;

        public const int ValueLimit = 2000;

        public const int MessageLimit = 4000;

        public const string NoIdentify = "The model needs its Identify command - CRT connects to the scope with it.";

        // ###########################################################################################
        // WHAT KEEPS A MODEL FROM BEING SENT IN FOR EVERYBODY - CRT asks before it sends, the server
        // when it takes it (OscilloscopeModelSubmissionRequest). Empty when nothing does. A command is
        // one line: CRT sends it with a line ending of its own, so a line break inside it would send
        // two commands. A value whose name the main file does not know is refused - the server is
        // always the newest party, so it is a mistake, not a later CRT's value.
        // ###########################################################################################
        public static IReadOnlyList<string> SubmissionProblems(
            string? vendor,
            string? seriesOrModel,
            IReadOnlyDictionary<string, string>? values,
            string? message)
        {
            var problems = new List<string>();

            if (string.IsNullOrWhiteSpace(vendor))
                problems.Add(OscilloscopeModelRules.NoVendor);
            else if (vendor.Trim().Length > OscilloscopeModelRules.NameLimit)
                problems.Add($"The vendor is longer than {OscilloscopeModelRules.NameLimit} characters.");

            if (string.IsNullOrWhiteSpace(seriesOrModel))
                problems.Add(OscilloscopeModelRules.NoModel);
            else if (seriesOrModel.Trim().Length > OscilloscopeModelRules.NameLimit)
                problems.Add($"The series or model is longer than {OscilloscopeModelRules.NameLimit} characters.");

            if (OscilloscopeModelRules.HasLineBreakOrControl(vendor) || OscilloscopeModelRules.HasLineBreakOrControl(seriesOrModel))
                problems.Add("The name holds a line break or another character that is not text.");

            if (values is null)
            {
                problems.Add("The model came without its commands.");
            }
            else
            {
                foreach ((string name, string value) in values)
                {
                    if (!OscilloscopeModelValues.Names.Contains(name))
                        problems.Add($"\"{name}\" is no value of an oscilloscope model this server knows.");
                    else if ((value ?? string.Empty).Length > OscilloscopeModelRules.ValueLimit)
                        problems.Add($"\"{name}\" is longer than {OscilloscopeModelRules.ValueLimit} characters.");
                    else if (OscilloscopeModelRules.HasLineBreakOrControl(value))
                        problems.Add($"\"{name}\" holds a line break or another character that is not text - a command is one line.");
                }

                if (!values.TryGetValue("identify", out string? identify) || string.IsNullOrWhiteSpace(identify))
                    problems.Add(OscilloscopeModelRules.NoIdentify);
            }

            if ((message ?? string.Empty).Length > OscilloscopeModelRules.MessageLimit)
                problems.Add($"The message is longer than {OscilloscopeModelRules.MessageLimit} characters.");

            return problems;
        }

        private static bool HasLineBreakOrControl(string? text) => (text ?? string.Empty).Any(char.IsControl);

        // Whether `model` is the one named `vendor` + `seriesOrModel`.
        public static bool IsNamed(MainDocument.OscilloscopeRow model, string vendor, string seriesOrModel)
        {
            ArgumentNullException.ThrowIfNull(model);

            return OscilloscopeModelRules.SameName(model.Brand, model.SeriesOrModel, vendor, seriesOrModel);
        }

        // Two names one model: trimmed, capitals aside - nothing more (accents count, "ss" is not "ß").
        // The server's store asks this too rather than its database's collation, which ignores accents
        // and would take two models the main file keeps apart for one (code review, 2026-10-29).
        public static bool SameName(string? vendor, string? seriesOrModel, string? otherVendor, string? otherSeriesOrModel) =>
            string.Equals(
                OscilloscopeModelRules.Key(vendor, seriesOrModel),
                OscilloscopeModelRules.Key(otherVendor, otherSeriesOrModel),
                StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // What keeps `vendor` + `seriesOrModel` from naming a model - null when nothing does. `others`
        // is every OTHER model in the list (published and own, the one being renamed left out).
        // ###########################################################################################
        public static string? NameProblem(string? vendor, string? seriesOrModel, IEnumerable<MainDocument.OscilloscopeRow> others)
        {
            ArgumentNullException.ThrowIfNull(others);

            if (string.IsNullOrWhiteSpace(vendor))
                return OscilloscopeModelRules.NoVendor;

            if (string.IsNullOrWhiteSpace(seriesOrModel))
                return OscilloscopeModelRules.NoModel;

            return others.Any(other => OscilloscopeModelRules.IsNamed(other, vendor, seriesOrModel))
                ? OscilloscopeModelRules.Taken(vendor, seriesOrModel)
                : null;
        }

        // ###########################################################################################
        // A new model named `vendor` + `seriesOrModel` with every command, setting and list of
        // `source` - the "similar model" it starts from. Fields this CRT does not know come along.
        // ###########################################################################################
        public static MainDocument.OscilloscopeRow CopyOf(MainDocument.OscilloscopeRow source, string vendor, string seriesOrModel)
        {
            MainDocument.OscilloscopeRow copy = OscilloscopeModelRules.Copy(source);

            copy.Brand = vendor.Trim();
            copy.SeriesOrModel = seriesOrModel.Trim();

            return copy;
        }

        // A copy of `model` with `change` made to it - the list held is never changed in place.
        public static MainDocument.OscilloscopeRow With(MainDocument.OscilloscopeRow model, Action<MainDocument.OscilloscopeRow> change)
        {
            ArgumentNullException.ThrowIfNull(change);

            MainDocument.OscilloscopeRow copy = OscilloscopeModelRules.Copy(model);
            change(copy);

            return copy;
        }

        // ###########################################################################################
        // The user's own models as they are to be listed beside `published` (see the header), and
        // whether that differs from `own` - the file is written again then. An own model whose name
        // is taken by another own model above it is named apart too: hand-edited, the file can hold
        // a name twice, and the second was unreachable.
        // ###########################################################################################
        public static (IReadOnlyList<MainDocument.OscilloscopeRow> Own, bool Changed) Reconciled(
            IReadOnlyList<MainDocument.OscilloscopeRow> own,
            IReadOnlyList<MainDocument.OscilloscopeRow> published) =>
            OscilloscopeModelRules.Reconciled(own, published, out _);

        // The same - and `renamed`, each own model named apart: its vendor, the name it had, the name
        // it has now. What the Oscilloscope tab follows a model by (code review, 2026-10-29: guessed by
        // its name's start, a "(my version)" the user had named so already was taken for it).
        public static (IReadOnlyList<MainDocument.OscilloscopeRow> Own, bool Changed) Reconciled(
            IReadOnlyList<MainDocument.OscilloscopeRow> own,
            IReadOnlyList<MainDocument.OscilloscopeRow> published,
            out IReadOnlyList<(string Vendor, string From, string To)> renamed)
        {
            ArgumentNullException.ThrowIfNull(own);
            ArgumentNullException.ThrowIfNull(published);

            var publishedByName = new Dictionary<string, MainDocument.OscilloscopeRow>(StringComparer.OrdinalIgnoreCase);

            foreach (MainDocument.OscilloscopeRow model in published)
                publishedByName.TryAdd(OscilloscopeModelRules.Key(model.Brand, model.SeriesOrModel), model);

            var taken = new HashSet<string>(publishedByName.Keys, StringComparer.OrdinalIgnoreCase);
            var kept = new List<MainDocument.OscilloscopeRow>(own.Count);
            var renames = new List<(string Vendor, string From, string To)>();
            bool changed = false;

            foreach (MainDocument.OscilloscopeRow model in own)
            {
                string key = OscilloscopeModelRules.Key(model.Brand, model.SeriesOrModel);
                MainDocument.OscilloscopeRow listed = model;

                if (publishedByName.TryGetValue(key, out MainDocument.OscilloscopeRow? same) && OwnOscilloscopeModels.SameValues(model, same))
                {
                    changed = true;
                    continue;
                }

                if (taken.Contains(key))
                {
                    string free = OscilloscopeModelRules.FreeName(model.Brand, model.SeriesOrModel.Trim() + OwnOscilloscopeModels.MyVersionSuffix, taken);
                    listed = OscilloscopeModelRules.With(model, copy => copy.SeriesOrModel = free);
                    renames.Add((model.Brand.Trim(), model.SeriesOrModel.Trim(), free));
                    changed = true;
                }

                taken.Add(OscilloscopeModelRules.Key(listed.Brand, listed.SeriesOrModel));
                kept.Add(listed);
            }

            renamed = renames;
            return (kept, changed);
        }

        // `wanted` when no model of `vendor` has it, else "<wanted> 2", "<wanted> 3", ...
        private static string FreeName(string vendor, string wanted, HashSet<string> taken)
        {
            if (!taken.Contains(OscilloscopeModelRules.Key(vendor, wanted)))
                return wanted;

            for (int number = 2; ; number++)
            {
                string candidate = $"{wanted} {number}";

                if (!taken.Contains(OscilloscopeModelRules.Key(vendor, candidate)))
                    return candidate;
            }
        }

        private static string Key(string? vendor, string? seriesOrModel) => $"{vendor?.Trim()}\n{seriesOrModel?.Trim()}";

        private static MainDocument.OscilloscopeRow Copy(MainDocument.OscilloscopeRow source)
        {
            ArgumentNullException.ThrowIfNull(source);

            MainDocument.OscilloscopeRow copy = OwnOscilloscopeModels.Copy(source);

            copy.Unknown = source.Unknown is null ? null : new Dictionary<string, JsonElement>(source.Unknown);

            return copy;
        }
    }
}
