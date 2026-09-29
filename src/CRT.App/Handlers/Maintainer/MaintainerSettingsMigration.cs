using System;
using System.IO;
using System.Text.Json;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // CARRIES THE ONE SETTING WORTH KEEPING out of the separate maintainer application's settings
    // file, once (2026-09-29: CRT Maintainer became CRT's Maintainer tab).
    //
    // That application kept "CRT-Maintainer-Settings.json" beside CRT's own settings: its window
    // placement and "Show changes only". The placement is NOT carried over - the tab lives in CRT's
    // window, which remembers its own. "Show changes only" is handed to `setShowChangesOnly` (CRT
    // passes UserSettings.MaintainerShowChangesOnly's setter) and the file is then DELETED, so
    // nothing is left behind that no program reads.
    //
    // *** SOFT ON EVERY FAILURE, like the file it reads. *** It was a convenience; losing it costs a
    // tick in a check box, never a start-up. A file that is not JSON is deleted too (there is
    // nothing in it to carry, and leaving it would retry the failure on every launch); a file that
    // cannot be read or deleted right now - held open, no permission - is left for the next launch.
    //
    // Takes the folder rather than resolving AppData, so a test points it at a temp folder and
    // never at the user's real one.
    // ###########################################################################################
    public static class MaintainerSettingsMigration
    {
        // The separate application's file name (its MaintainerSettingsStore.FileName).
        public const string LegacyFileName = "CRT-Maintainer-Settings.json";

        // True when the file was dealt with (carried and deleted, or deleted as unreadable) - false
        // when there was none, or it has to wait for another launch.
        public static bool Apply(string settingsFolder, Action<bool> setShowChangesOnly)
        {
            ArgumentNullException.ThrowIfNull(setShowChangesOnly);

            if (string.IsNullOrWhiteSpace(settingsFolder))
                return false;

            string path = Path.Combine(settingsFolder, MaintainerSettingsMigration.LegacyFileName);

            try
            {
                if (!File.Exists(path))
                    return false;

                bool? showChangesOnly = MaintainerSettingsMigration.ReadShowChangesOnly(File.ReadAllText(path));

                if (showChangesOnly is bool value)
                    setShowChangesOnly(value);

                File.Delete(path);
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return false;
            }
        }

        // ###########################################################################################
        // The "ShowChangesOnly" value in the old file's JSON, or null when the JSON is not an
        // object, is not JSON at all, or does not carry a true/false there. The old file was
        // written by System.Text.Json with its default (PascalCase) names.
        // ###########################################################################################
        internal static bool? ReadShowChangesOnly(string json)
        {
            try
            {
                using JsonDocument document = JsonDocument.Parse(json);

                if (document.RootElement.ValueKind == JsonValueKind.Object &&
                    document.RootElement.TryGetProperty("ShowChangesOnly", out JsonElement value) &&
                    value.ValueKind is JsonValueKind.True or JsonValueKind.False)
                {
                    return value.GetBoolean();
                }

                return null;
            }
            catch (JsonException)
            {
                return null;
            }
        }
    }
}
