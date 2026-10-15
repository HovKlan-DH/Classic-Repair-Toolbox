using CRT;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One user-drawn trace on a schematic, as stored in the traces file: its colour, whether it is
    // shown, and its points (each 0..1 across the image, so they survive a different image size).
    // ###########################################################################################
    public class TraceModel
    {
        public string Color { get; set; } = string.Empty;
        public bool Visible { get; set; } = true;
        public List<PointModel> Nodes { get; set; } = new();
    }

    public class PointModel
    {
        public double X { get; set; }
        public double Y { get; set; }
    }

    // ###########################################################################################
    // The traces the user draws on schematics (PolylineManagement), kept in ONE file for every
    // board: board key ("hardware|board") -> schematic name -> traces. Moved out of
    // PolylineManagement.cs (code review, 2026-10-13), which also fixed what it did wrong:
    //
    //   - it was written IN PLACE (File.WriteAllText), so a crash mid-write left a truncated file;
    //     it goes through AtomicJsonFile now, like the settings and the worklogs;
    //   - a file that failed to parse was read as "no traces at all" - and the next trace saved
    //     wrote that over every board's traces. It is moved aside as "<name>.corrupt-<date>" now
    //     (CorruptFileSetAside); a file that cannot even be read, or not be moved, is never
    //     written over this session;
    //   - a schematic image renamed lost its traces, which are filed by schematic name. They follow
    //     the rename now, with the user's other per-schematic data (LocalSchematicRename.CatchUp,
    //     through RenameSchematics);
    //   - its static constructor read the user's REAL file the first time anything touched it, a
    //     headless test included. Load() is called once at startup (App); tests use LoadFrom(path).
    //
    // Schematic names are looked up without regard to capitals, like every other schematic name.
    // Touched from the UI thread only, but locked anyway: it is cheap, and a static is shared.
    // ###########################################################################################
    public static class TraceStorage
    {
        private static readonly object Sync = new();
        private static readonly JsonSerializerOptions SaveOptions = new() { WriteIndented = true };

        // Empty until Load/LoadFrom - and then nothing is saved anywhere.
        private static string filePath = string.Empty;
        private static Dictionary<string, Dictionary<string, List<TraceModel>>> data = new();

        // ###########################################################################################
        // Resolves the traces file beside the settings and loads it. Call once at startup.
        // ###########################################################################################
        public static void Load()
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    AppConfig.AppFolderName);

                Directory.CreateDirectory(directory);
                TraceStorage.LoadFrom(Path.Combine(directory, AppConfig.TracesFileName));
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load user-drawn traces: [{ex.Message}] - none shown, and none saved this session");
            }
        }

        // ###########################################################################################
        // Loads the traces from `path` and makes it the save target - see the header for what
        // happens to a file that cannot be read or parsed.
        // ###########################################################################################
        internal static void LoadFrom(string path)
        {
            lock (TraceStorage.Sync)
            {
                TraceStorage.filePath = path;
                TraceStorage.data = new Dictionary<string, Dictionary<string, List<TraceModel>>>();

                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return;
                }

                string json;

                try
                {
                    json = File.ReadAllText(path);
                }
                catch (Exception ex)
                {
                    Logger.Warning($"User-drawn traces could not be read: [{path}] [{ex.Message}] - none shown, and the file is not written this session");
                    TraceStorage.filePath = string.Empty;
                    return;
                }

                Dictionary<string, Dictionary<string, List<TraceModel>>>? parsed = null;
                string reason = "it holds no traces";

                try
                {
                    parsed = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, List<TraceModel>>>>(json);
                }
                catch (Exception ex)
                {
                    reason = ex.Message;
                }

                if (parsed == null)
                {
                    string? keptAs = CorruptFileSetAside.SetAside(path, DateTime.Now);

                    if (keptAs == null)
                    {
                        Logger.Warning($"User-drawn traces could not be parsed [{reason}] nor moved aside: [{path}] - none shown, and the file is not written this session");
                        TraceStorage.filePath = string.Empty;
                    }
                    else
                    {
                        Logger.Warning($"User-drawn traces could not be parsed [{reason}] - the file was kept as [{keptAs}]");
                    }

                    return;
                }

                foreach (var (boardKey, boardTraces) in parsed)
                {
                    if (boardKey == null || boardTraces == null)
                    {
                        continue;
                    }

                    // System.Text.Json builds its own dictionaries, without the comparer.
                    //
                    // *** NAMES DIFFERING ONLY IN CAPITALS ARE ONE SCHEMATIC, AND KEEP BOTH LISTS (code
                    // review, 2026-10-13). *** An older store kept "Top" and "top" apart; keeping only
                    // the first and writing the file back lost the other's traces for good. They are
                    // joined, the rule RenameSchematics keeps - nothing is dropped.
                    var bySchematic = new Dictionary<string, List<TraceModel>>(StringComparer.OrdinalIgnoreCase);

                    foreach (var (schematicName, traces) in boardTraces)
                    {
                        if (schematicName == null || traces == null)
                        {
                            continue;
                        }

                        bySchematic[schematicName] = bySchematic.TryGetValue(schematicName, out var earlier)
                            ? [.. earlier, .. traces]
                            : traces;
                    }

                    TraceStorage.data[boardKey] = bySchematic;
                }
            }
        }

        public static List<TraceModel> GetTraces(string boardKey, string schematicName)
        {
            lock (TraceStorage.Sync)
            {
                if (TraceStorage.data.TryGetValue(boardKey, out var boardData) &&
                    boardData.TryGetValue(schematicName, out var traces))
                {
                    return traces;
                }

                return new List<TraceModel>();
            }
        }

        public static void SaveTraces(string boardKey, string schematicName, List<TraceModel> traces)
        {
            lock (TraceStorage.Sync)
            {
                if (!TraceStorage.data.TryGetValue(boardKey, out var boardData))
                {
                    boardData = new Dictionary<string, List<TraceModel>>(StringComparer.OrdinalIgnoreCase);
                    TraceStorage.data[boardKey] = boardData;
                }

                boardData[schematicName] = traces;
                TraceStorage.Save();
            }
        }

        // ###########################################################################################
        // Moves the board's traces to wherever `resolve` says each schematic's name went, all at
        // once (two images swapping names swap their traces), and saves. `before` is the board's
        // traces as they were, for RestoreBoard - null when nothing moved.
        //
        // A name moved onto one that already has traces keeps both lists - nothing is dropped.
        // Returns false, with nothing changed, when the file could not be written.
        // ###########################################################################################
        internal static bool RenameSchematics(
            string boardKey,
            Func<string, string?> resolve,
            out Dictionary<string, List<TraceModel>>? before)
        {
            lock (TraceStorage.Sync)
            {
                before = null;

                if (!TraceStorage.data.TryGetValue(boardKey, out var board) || board.Count == 0)
                {
                    return true;
                }

                var moves = board
                    .Select(pair => (Name: pair.Key, Now: resolve(pair.Key), Traces: pair.Value))
                    .ToList();

                if (moves.All(move => move.Now == null || string.Equals(move.Now, move.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }

                var renamed = new Dictionary<string, List<TraceModel>>(StringComparer.OrdinalIgnoreCase);

                foreach (var (name, _, traces) in moves.Where(move => move.Now == null))
                {
                    renamed[name] = traces;
                }

                foreach (var (_, now, traces) in moves.Where(move => move.Now != null))
                {
                    renamed[now!] = renamed.TryGetValue(now!, out var already)
                        ? already.Concat(traces).ToList()
                        : traces;
                }

                var previous = new Dictionary<string, List<TraceModel>>(board, StringComparer.OrdinalIgnoreCase);
                TraceStorage.data[boardKey] = renamed;

                if (!TraceStorage.Save())
                {
                    TraceStorage.data[boardKey] = previous;
                    return false;
                }

                before = previous;
                return true;
            }
        }

        // Puts a board's traces back as RenameSchematics found them.
        internal static void RestoreBoard(string boardKey, Dictionary<string, List<TraceModel>> before)
        {
            lock (TraceStorage.Sync)
            {
                TraceStorage.data[boardKey] = before;
                TraceStorage.Save();
            }
        }

        // Every board's traces to the file, atomically. False when the write failed. With no file
        // to write - not loaded, or one that must not be written over - there is nothing to fail.
        private static bool Save() =>
            string.IsNullOrEmpty(TraceStorage.filePath) ||
            AtomicJsonFile.Write(TraceStorage.filePath, TraceStorage.data, SaveOptions, "user-drawn traces");
    }
}
