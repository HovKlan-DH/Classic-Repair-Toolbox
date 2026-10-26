using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE OSCILLOSCOPE MODELS A USER MADE THEIRSELF, found when CRT 4.0 first starts.
    //
    // Before 4.0 the way to add a scope CRT did not know was to add a row to the Oscilloscope sheet
    // of one's own copy of the main Excel file, with the launch data check switched off so the sync
    // did not put the file back (the Wiki's "My scope is not listed"). In 4.0 the published models
    // come from the 4.0 main file, so a row a user added that way would silently vanish. This finds
    // those rows, for 4.0 to keep as the user's own models:
    //
    //   - a model the published list does not have (by brand and model, ignoring capitals) is kept
    //     as it is;
    //   - a model the published list has, with any value different, is kept TOO - as a copy named
    //     "<model> (my version)" beside the published one - because the user may have changed a
    //     command to suit their firmware, and nothing a user made is dropped. An old copy of the
    //     workbook differs the same way and is kept the same way: one surplus entry, which the user
    //     can delete, is better than a lost edit.
    //
    // A model the user's copy has exactly as published is the published one, and is not repeated.
    // ###########################################################################################
    public static class OwnOscilloscopeModels
    {
        public const string MyVersionSuffix = " (my version)";

        private static readonly PropertyInfo[] Values = typeof(MainDocument.OscilloscopeRow)
            .GetProperties()
            .Where(property => property.PropertyType == typeof(string))
            .ToArray();

        public static IReadOnlyList<MainDocument.OscilloscopeRow> Find(
            IReadOnlyList<MainDocument.OscilloscopeRow> local,
            IReadOnlyList<MainDocument.OscilloscopeRow> published)
        {
            ArgumentNullException.ThrowIfNull(local);
            ArgumentNullException.ThrowIfNull(published);

            var publishedByModel = new Dictionary<string, MainDocument.OscilloscopeRow>(StringComparer.OrdinalIgnoreCase);

            foreach (MainDocument.OscilloscopeRow model in published)
                publishedByModel.TryAdd(OwnOscilloscopeModels.Key(model), model);

            var own = new List<MainDocument.OscilloscopeRow>();

            foreach (MainDocument.OscilloscopeRow model in local)
            {
                if (!publishedByModel.TryGetValue(OwnOscilloscopeModels.Key(model), out MainDocument.OscilloscopeRow? same))
                {
                    own.Add(model);
                }
                else if (!OwnOscilloscopeModels.SameValues(model, same))
                {
                    MainDocument.OscilloscopeRow copy = OwnOscilloscopeModels.Copy(model);
                    copy.SeriesOrModel = model.SeriesOrModel + OwnOscilloscopeModels.MyVersionSuffix;
                    own.Add(copy);
                }
            }

            return own;
        }

        private static string Key(MainDocument.OscilloscopeRow model) => $"{model.Brand.Trim()}\n{model.SeriesOrModel.Trim()}";

        // The commands and lists; the brand and model are what identify a model, and were matched
        // already (ignoring capitals).
        private static bool SameValues(MainDocument.OscilloscopeRow a, MainDocument.OscilloscopeRow b) =>
            OwnOscilloscopeModels.Values
                .Where(property => property.Name is not (nameof(MainDocument.OscilloscopeRow.Brand) or nameof(MainDocument.OscilloscopeRow.SeriesOrModel)))
                .All(property => string.Equals(
                    ((string?)property.GetValue(a))?.Trim(),
                    ((string?)property.GetValue(b))?.Trim(),
                    StringComparison.Ordinal));

        private static MainDocument.OscilloscopeRow Copy(MainDocument.OscilloscopeRow model)
        {
            var copy = new MainDocument.OscilloscopeRow();

            foreach (PropertyInfo property in OwnOscilloscopeModels.Values)
                property.SetValue(copy, property.GetValue(model));

            return copy;
        }
    }
}
