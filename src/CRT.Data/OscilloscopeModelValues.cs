using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // AN OSCILLOSCOPE MODEL'S VALUES BY THE MAIN FILE'S OWN NAMES FOR THEM - how a model sent in
    // travels (OscilloscopeModelSubmissionRequest) and how the server holds it until it is published:
    // every value of MainDocument.OscilloscopeRow but the vendor and the series or model, keyed by its
    // JSON name ("port", "identify", "dumpImage").
    //
    // Only the names the row knows are taken. The server is always the newest party, so a name it
    // does not know is a mistake, and OscilloscopeModelRules.SubmissionProblems refuses it.
    // ###########################################################################################
    public static class OscilloscopeModelValues
    {
        private static readonly IReadOnlyList<(string Name, PropertyInfo Property)> Known = typeof(MainDocument.OscilloscopeRow)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => property.PropertyType == typeof(string))
            .Where(property => property.Name is not (nameof(MainDocument.OscilloscopeRow.Brand) or nameof(MainDocument.OscilloscopeRow.SeriesOrModel)))
            .Select(property => (property.GetCustomAttribute<JsonPropertyNameAttribute>()!.Name, property))
            .ToList();

        // The names a model's values go by, in the main file's order.
        public static IReadOnlyList<string> Names { get; } = OscilloscopeModelValues.Known.Select(known => known.Name).ToList();

        private static readonly IReadOnlyDictionary<string, PropertyInfo> ByName =
            OscilloscopeModelValues.Known.ToDictionary(known => known.Name, known => known.Property, StringComparer.Ordinal);

        // ###########################################################################################
        // ONE VALUE BY ITS NAME - THE ONE MAP OF A MODEL'S VALUES (code review, 2026-10-29: they were
        // mapped four times over - OscilloscopeEntry.From's list, ScopeCommandResolver's switch, the
        // editor's setter switch, and this - and a value added to the row was sent and published while
        // the tab, Try it and the editor never saw it). The editor's fields (ScopeModelFields) read and
        // write through these, OscilloscopeEntry copies the row and answers ValueOf by them.
        // ###########################################################################################
        public static string Get(MainDocument.OscilloscopeRow model, string name)
        {
            ArgumentNullException.ThrowIfNull(model);

            return (string?)OscilloscopeModelValues.PropertyOf(name).GetValue(model) ?? string.Empty;
        }

        public static void Set(MainDocument.OscilloscopeRow model, string name, string? value)
        {
            ArgumentNullException.ThrowIfNull(model);

            OscilloscopeModelValues.PropertyOf(name).SetValue(model, value ?? string.Empty);
        }

        // The row's property a value of `name` lives in - an unknown name is a mistake in the code.
        internal static PropertyInfo PropertyOf(string name) =>
            OscilloscopeModelValues.ByName.TryGetValue(name, out PropertyInfo? property)
                ? property
                : throw new ArgumentOutOfRangeException(nameof(name), name, "No value of an oscilloscope model goes by this name.");

        // Every value of `model` but its name - an empty one included, so the set says the whole model.
        public static IReadOnlyDictionary<string, string> Of(MainDocument.OscilloscopeRow model)
        {
            ArgumentNullException.ThrowIfNull(model);

            var values = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach ((string name, PropertyInfo property) in OscilloscopeModelValues.Known)
                values[name] = (string?)property.GetValue(model) ?? string.Empty;

            return values;
        }

        // ###########################################################################################
        // A model as one string - what "Run full test suite went through on this model as it is now"
        // compares (the Oscilloscope tab's "Submit this model"): the name and EVERY value, as the
        // model is sent - a list of the tab's own would miss a value added to the model, and the
        // administrator be told of a suite never run on it (code reviews, 2026-10-29).
        // ###########################################################################################
        public static string Fingerprint(MainDocument.OscilloscopeRow model)
        {
            ArgumentNullException.ThrowIfNull(model);

            return string.Join(
                "\n",
                OscilloscopeModelValues.Of(model)
                    .OrderBy(value => value.Key, StringComparer.Ordinal)
                    .Select(value => $"{value.Key}={value.Value}")
                    .Prepend(model.SeriesOrModel.Trim())
                    .Prepend(model.Brand.Trim()));
        }

        // The model `vendor` + `seriesOrModel` with `values` - names trimmed, each known value where
        // the row keeps it; a name the row does not know is left out.
        public static MainDocument.OscilloscopeRow ToModel(string? vendor, string? seriesOrModel, IReadOnlyDictionary<string, string>? values)
        {
            var model = new MainDocument.OscilloscopeRow
            {
                Brand = vendor?.Trim() ?? string.Empty,
                SeriesOrModel = seriesOrModel?.Trim() ?? string.Empty,
            };

            foreach ((string name, PropertyInfo property) in OscilloscopeModelValues.Known)
            {
                if (values is not null && values.TryGetValue(name, out string? value))
                    property.SetValue(model, value ?? string.Empty);
            }

            return model;
        }
    }
}
