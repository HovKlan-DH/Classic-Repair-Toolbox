using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Which WORKBOOK GENERATION the data tree is on, and what a file of that generation is called
    // (NewContributeStrategy.md, open question 3, answered 2026-09-21).
    //
    // A "generation" is a version stamped into a workbook's FILE NAME - the unversioned original,
    // then v2.0.0, then whatever comes next. It is NOT a published revision: a revision is one
    // merge's worth of change to a system (systems.current_revision), while a generation is a
    // compatibility target serving a range of application builds. Keep the two apart; the strategy
    // document records that conflating them sent an earlier reading of open question 5 astray.
    //
    // *** AN OLDER GENERATION IS FROZEN, NOT STALE. *** The owner's decision is that
    // publishing writes ONLY the newest generation and never touches an older one, because older
    // generations still serve older application builds and keep working precisely because nothing
    // writes to them. This class exists so that rule has one implementation rather than being
    // re-derived per call site.
    //
    // *** THE TARGET IS DISCOVERED, NEVER CONFIGURED. *** ResolveNewest reads the tree. A
    // configured generation would be a second place the truth lives, and its failure mode is
    // silent: the tree moves on, the setting does not, and published work lands in a frozen
    // generation that no current build reads, with nothing throwing.
    //
    // NOTE THE ASYMMETRY WITH THE READ PATH and keep it. DataManager.ResolveMainExcelFile picks
    // the newest generation AT OR BELOW the application's own version, because an app must not
    // read a workbook newer than itself. Publishing picks the newest that EXISTS, full stop -
    // there is no "own version" to compare against on a server, and the point of publishing is to
    // write the current generation.
    //
    // THE TWO NAMING CONVENTIONS ARE DIFFERENT AND BOTH ARE REAL. The master workbook separates
    // its version with a DOT ("Classic-Repair-Toolbox.v2.0.0.xlsx"); a board workbook separates it
    // with a SPACE ("Data C64 250407 v2.0.0.xlsx"). This is not a tidy-up opportunity - both
    // spellings are shipped, both are named in a master workbook's own ExcelDataFile column, and
    // renaming either breaks the tree for builds that look for it.
    // ###########################################################################################
    public static class DataGenerationRules
    {
        public const string WorkbookExtension = ".xlsx";

        // The master workbook, which is what names every board file of its own generation.
        public const string MasterStem = "Classic-Repair-Toolbox";

        // A board workbook's name always starts here; what follows is the board's own naming.
        public const string BoardStemPrefix = "Data ";

        // ###########################################################################################
        // The newest generation among the given file names, or null when only the unversioned
        // original is present (which IS a generation - the first one - and is represented by null
        // rather than by a fabricated 0.0.0, so a caller cannot accidentally order it against real
        // versions or print it as a version that never existed).
        //
        // Unparseable version-looking names are IGNORED rather than throwing: the tree is synced
        // and a stray file in it must not stop a publish.
        // ###########################################################################################
        public static Version? ResolveNewestGeneration(IEnumerable<string>? fileNames)
        {
            Version? newest = null;

            foreach (string candidate in fileNames ?? [])
            {
                Version? version = DataGenerationRules.TryReadGeneration(candidate);

                if (version != null && (newest == null || version > newest))
                {
                    newest = version;
                }
            }

            return newest;
        }

        // ###########################################################################################
        // The generation THE TREE is on, read from the MASTER workbooks at its root.
        //
        // *** THIS EXISTS BECAUSE A NEW SYSTEM HAS NO FILES TO READ A GENERATION FROM. ***
        // ResolveNewestGeneration answers from the files already in a system's folder, which is
        // right for an existing board and useless for one being created: an empty folder yields
        // null, and null means UNVERSIONED - so a brand-new system would publish as
        // "Data C64 250407.xlsx", landing in the frozen generation that serves every pre-2.0.0
        // build. That is precisely the file the project owner said must never be written.
        //
        // The master workbooks are the right source because they ARE the generations: the tree
        // root carries "Classic-Repair-Toolbox.xlsx" for the original and
        // "Classic-Repair-Toolbox.v2.0.0.xlsx" for the 2.0.0 generation, and each references its
        // own board files. A new system must join the newest of those.
        //
        // Returns null only when the root holds no versioned master at all, which is a tree that
        // has no 2.0.0 generation yet - and the caller must refuse rather than fall back to
        // unversioned.
        // ###########################################################################################
        public static Version? ResolveNewestGenerationFromTree(string? dataRoot)
        {
            if (string.IsNullOrWhiteSpace(dataRoot) || !Directory.Exists(dataRoot))
                return null;

            IEnumerable<string> masters;

            try
            {
                masters = Directory
                    .EnumerateFiles(dataRoot, DataGenerationRules.MasterStem + "*" + DataGenerationRules.WorkbookExtension)
                    .Select(Path.GetFileName)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Select(name => name!);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                // A tree that cannot be enumerated is not a tree to publish into. Null makes the
                // caller refuse, which is the safe direction.
                CrtLog.Warning($"Could not read the data root [{dataRoot}] to resolve its generation - [{exception.Message}]");
                return null;
            }

            return DataGenerationRules.ResolveNewestGeneration(masters);
        }

        // ###########################################################################################
        // The generation stamped into one file name, or null when it carries none (the unversioned
        // original) or cannot be parsed.
        //
        // Accepts BOTH separators deliberately - see the header. The separator is whatever
        // character precedes the "v", so this reads a master and a board file with one rule
        // instead of two that could disagree.
        // ###########################################################################################
        public static Version? TryReadGeneration(string? fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return null;

            string name = Path.GetFileNameWithoutExtension(fileName.Trim());

            if (string.IsNullOrWhiteSpace(name))
                return null;

            int marker = name.LastIndexOf('v');

            // A "v" must exist, must not be the first character (there would be no name before
            // it), and must be preceded by the dot or space that separates it from the name -
            // otherwise "Data VIC20 250403" would read its own "V" as a version marker.
            if (marker < 1)
                return null;

            char separator = name[marker - 1];

            if (separator != '.' && separator != ' ')
                return null;

            string versionPart = name[(marker + 1)..];

            return Version.TryParse(versionPart, out Version? parsed) ? parsed : null;
        }

        // ###########################################################################################
        // The master workbook's file name for a generation. Null generation means the unversioned
        // original.
        // ###########################################################################################
        public static string BuildMasterFileName(Version? generation) =>
            generation == null
                ? DataGenerationRules.MasterStem + DataGenerationRules.WorkbookExtension
                : $"{DataGenerationRules.MasterStem}.v{generation}{DataGenerationRules.WorkbookExtension}";

        // ###########################################################################################
        // A board workbook's file name for a generation, given the board's stem as it appears
        // without any version ("Data C64 250407").
        //
        // The stem is taken as given rather than rebuilt from manufacturer/hardware/board: board
        // file names do not follow the folder names mechanically ("Data C128DCR 250477" sits under
        // C128/250477, and "Data VIC20 250403" under VIC-20/250403), so deriving one would rename
        // files that are named in a master workbook and referenced by the sync manifest.
        // ###########################################################################################
        public static string BuildBoardFileName(string boardStem, Version? generation)
        {
            string stem = (boardStem ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(stem))
                return string.Empty;

            return generation == null
                ? stem + DataGenerationRules.WorkbookExtension
                : $"{stem} v{generation}{DataGenerationRules.WorkbookExtension}";
        }

        // ###########################################################################################
        // A board workbook's version-free stem, so a file of one generation can be renamed into
        // another. "Data C64 250407 v2.0.0.xlsx" gives "Data C64 250407".
        // ###########################################################################################
        public static string ReadBoardStem(string? fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return string.Empty;

            string name = Path.GetFileNameWithoutExtension(fileName.Trim());

            if (string.IsNullOrWhiteSpace(name))
                return string.Empty;

            if (DataGenerationRules.TryReadGeneration(fileName) == null)
                return name;

            int marker = name.LastIndexOf('v');

            // TryReadGeneration already proved there is a separator before the marker, so trimming
            // back past it cannot run off the front.
            return name[..(marker - 1)].TrimEnd();
        }

        // ###########################################################################################
        // Whether a file belongs to a generation OLDER than the target - the thing publishing must
        // never write. Expressed as its own question, rather than left to each call site to
        // compare versions, because getting it wrong is silent: the write succeeds and quietly
        // edits a frozen compatibility target.
        //
        // The unversioned original is older than every real generation, and equal to itself.
        // ###########################################################################################
        public static bool IsOlderGeneration(Version? candidate, Version? target)
        {
            if (candidate == null)
                return target != null;

            if (target == null)
                return false;

            return candidate < target;
        }
    }
}
