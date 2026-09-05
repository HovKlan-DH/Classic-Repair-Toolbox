using System;
using System.IO;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Writes a JSON file atomically: serialize to a sibling ".tmp" file, then swap it over the real
    // one, so a crash, a full disk or a killed process mid-write cannot leave a truncated,
    // unparseable file where working data used to be.
    //
    // This exists because the same dozen lines had been copied into three places (UserSettings.Save,
    // WorklogManager.SaveEntries and SaveWorkbook) - and only the first copy carried the note
    // explaining why File.Replace is used rather than File.Move, so a future fix to the swap
    // semantics would have been applied to one copy and silently missed in the others.
    // ###########################################################################################
    public static class AtomicJsonFile
    {
        // ###########################################################################################
        // Serializes value to path via a temp-file swap. Returns false (and logs) when nothing
        // reached disk, so callers can tell the user their data was not saved rather than reporting
        // a success that never happened. The temp file is cleaned up on failure.
        // ###########################################################################################
        public static bool Write<T>(string path, T value, JsonSerializerOptions options, string whatFailed)
        {
            if (string.IsNullOrEmpty(path))
            {
                Logger.Warning($"Failed to save {whatFailed}: no path given");
                return false;
            }

            string tempPath = path + ".tmp";

            try
            {
                var json = JsonSerializer.Serialize(value, options);
                File.WriteAllText(tempPath, json);

                // File.Replace, not File.Move: on Windows a rename-over cannot displace a file
                // that another handle (backup tool, sync client, antivirus) has open for reading,
                // while ReplaceFile swaps the names and succeeds. Move only ever runs for the
                // very first save, when no file exists yet to replace.
                if (File.Exists(path))
                    File.Replace(tempPath, path, destinationBackupFileName: null, ignoreMetadataErrors: true);
                else
                    File.Move(tempPath, path);

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to save {whatFailed}: [{path}] [{ex.Message}]");

                // Clean up the temp file if it was left behind. Best-effort: the real failure is
                // already logged above, and failing to delete a leftover temp file is not itself
                // worth surfacing - it is harmless and gets overwritten on the next save attempt.
                try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }

                return false;
            }
        }
    }
}
