using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHICH file holds a system's published board (NewContributeStrategy.md Phase 5, task 3).
    //
    // Split out of PublishedBoardReader because this is the part worth testing: resolving a
    // system folder from untrusted identity values, picking the right workbook GENERATION, and
    // deciding what "there is no board" means. The reader around it only opens the file.
    //
    // *** THE IDENTITY IS UNTRUSTED, EVEN HERE. *** Manufacturer/Hardware/Board arrive in a
    // submission, and a manufacturer of "../.." would otherwise relocate the system folder
    // outside the data tree - which on a READ means disclosing an arbitrary file's contents to a
    // maintainer, and would make this endpoint a file-disclosure hole rather than a review screen.
    // So every part goes through SubmissionPathRules exactly as the write paths do.
    //
    // *** IT PICKS THE NEWEST GENERATION, the same rule publishing writes with. *** Reading an
    // older, frozen generation would show the maintainer a diff against a board no current build
    // uses, and the generation gap would appear as changes the contributor never made.
    // ###########################################################################################
    public static class PublishedBoardLocator
    {
        // ###########################################################################################
        // Where this system's published board is, and whether it is there at all.
        //
        // A BLANK OR UNSAFE IDENTITY YIELDS "does not exist" rather than throwing. The manifest's
        // own validation already reports a bad identity with a message the contributor can act
        // on; throwing here would turn a maintainer opening that submission into an error page
        // instead of showing them the findings that explain it.
        // ###########################################################################################
        public static PublishedBoardLocation Locate(string? dataTreeRoot, SubmissionManifest manifest)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            if (string.IsNullOrWhiteSpace(dataTreeRoot))
                return PublishedBoardLocation.None;

            string relative = string.Join('/',
                new[] { manifest.Manufacturer, manifest.Hardware, manifest.Board }
                    .Where(part => !string.IsNullOrWhiteSpace(part)));

            if (string.IsNullOrWhiteSpace(relative))
                return PublishedBoardLocation.None;

            if (!SubmissionPathRules.TryResolve(dataTreeRoot, relative, out string systemFolder, out _))
                return PublishedBoardLocation.None;

            if (!Directory.Exists(systemFolder))
                return PublishedBoardLocation.None;

            // Only the workbooks, and only their names - the generation lives in the file name.
            string[] fileNames = Directory
                .EnumerateFiles(systemFolder, "*" + DataGenerationRules.WorkbookExtension)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrEmpty(name))
                .Select(name => name!)
                .ToArray();

            if (fileNames.Length == 0)
                return PublishedBoardLocation.None;

            Version? generation = DataGenerationRules.ResolveNewestGeneration(fileNames);

            // The board file's stem does not follow the folder names mechanically
            // ("Data C128DCR 250477" lives under C128/250477), so it is READ off the files that
            // are there rather than rebuilt from the identity. Any workbook of the target
            // generation answers it, since they all share one stem per system.
            string? match = fileNames.FirstOrDefault(name =>
                DataGenerationRules.TryReadGeneration(name) == generation);

            if (match is null)
                return PublishedBoardLocation.None;

            return new PublishedBoardLocation(
                Path.Combine(systemFolder, match),
                systemFolder,
                generation,
                Exists: true);
        }
    }

    // ###########################################################################################
    // Where a published board is - or that there is none, which is a NEW SYSTEM and a first-class
    // answer rather than a failure.
    // ###########################################################################################
    public sealed record PublishedBoardLocation(
        string WorkbookPath,
        string SystemFolder,
        Version? Generation,
        bool Exists)
    {
        public static readonly PublishedBoardLocation None =
            new(string.Empty, string.Empty, null, Exists: false);
    }
}
