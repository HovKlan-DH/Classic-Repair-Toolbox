using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The boards the published tree holds, as (manufacturer, hardware, board) - so a maintainer can
    // be assigned to a shipped board BEFORE anything has ever been submitted to it (Phase 6
    // roles, 2026-09-25).
    //
    // The `boards` table only gains a row when a submission arrives or a publish happens, so on
    // a fresh deployment it is nearly empty while the tree holds twenty-odd boards. The
    // administrator assigns maintainers by looking at the tree, and the row is created on the first
    // assignment (ISubmissionStore.EnsureBoardAsync).
    //
    // A board is a folder three levels down holding at least one workbook. The two shared-file
    // folders are skipped by name: "Generic shared files" at the top and "Shared files" beside a
    // manufacturer's boards are not boards and never get a maintainer. Skipped by
    // SubmissionFileScopes.IsSharedFolderName, the rule SubmissionValidator refuses such board names
    // by - so no board it skips can ever be published. A thin filesystem walk, tested against a
    // temp tree.
    // ###########################################################################################
    public static class PublishedBoardLister
    {
        public sealed record KnownBoard(string BoardId, string Manufacturer, string Hardware, string Board);

        public static IReadOnlyList<KnownBoard> List(string? dataTreeRoot)
        {
            var boards = new List<KnownBoard>();

            if (string.IsNullOrWhiteSpace(dataTreeRoot) || !Directory.Exists(dataTreeRoot))
                return boards;

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

                            boards.Add(new KnownBoard(
                                BoardDescriptorRules.BuildBoardId(manufacturer, hardware, board),
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

            return boards.OrderBy(board => board.BoardId, StringComparer.Ordinal).ToList();
        }
    }
}
