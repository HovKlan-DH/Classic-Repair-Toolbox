using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The boards the published tree holds, as (manufacturer, hardware, board) - so a maintainer can
    // be assigned to a shipped board BEFORE anything has ever been submitted to it (Phase 6
    // roles, 2026-09-25).
    //
    // The `systems` table only gains a row when a submission arrives or a publish happens, so on
    // a fresh deployment it is nearly empty while the tree holds twenty-odd boards. The
    // administrator assigns maintainers by looking at the tree, and the row is created on the first
    // assignment (ISubmissionStore.EnsureSystemAsync).
    //
    // A board is a folder three levels down holding at least one workbook. The two shared-file
    // folders are skipped by name: "Generic shared files" at the top and "Shared files" beside a
    // manufacturer's boards are not systems and never get a maintainer. Skipped by
    // SubmissionFileScopes.IsSharedFolderName, the rule SubmissionValidator refuses such board names
    // by - so no board it skips can ever be published. A thin filesystem walk, tested against a
    // temp tree.
    // ###########################################################################################
    public static class PublishedSystemLister
    {
        public sealed record KnownSystem(string SystemId, string Manufacturer, string Hardware, string Board);

        public static IReadOnlyList<KnownSystem> List(string? dataTreeRoot)
        {
            var systems = new List<KnownSystem>();

            if (string.IsNullOrWhiteSpace(dataTreeRoot) || !Directory.Exists(dataTreeRoot))
                return systems;

            try
            {
                foreach (string manufacturerFolder in Directory.EnumerateDirectories(dataTreeRoot))
                {
                    string manufacturer = Path.GetFileName(manufacturerFolder);

                    if (SubmissionFileScopes.IsSharedFolderName(manufacturer))
                        continue;

                    foreach (string hardwareFolder in Directory.EnumerateDirectories(manufacturerFolder))
                    {
                        string hardware = Path.GetFileName(hardwareFolder);

                        if (SubmissionFileScopes.IsSharedFolderName(hardware))
                            continue;

                        foreach (string boardFolder in Directory.EnumerateDirectories(hardwareFolder))
                        {
                            string board = Path.GetFileName(boardFolder);

                            if (!Directory.EnumerateFiles(boardFolder, "*.xlsx").Any())
                                continue;

                            systems.Add(new KnownSystem(
                                SystemDescriptorRules.BuildSystemId(manufacturer, hardware, board),
                                manufacturer,
                                hardware,
                                board));
                        }
                    }
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A tree that cannot be fully read lists what it could. The administrator sees
                // fewer boards than expected, which is visible; nothing is written on this path.
            }

            return systems.OrderBy(system => system.SystemId, StringComparer.Ordinal).ToList();
        }
    }
}
