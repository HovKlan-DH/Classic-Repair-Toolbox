using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE ID A NEW ROW OF A RECORD'S OWN LIST TAKES: one past the highest it holds, 1 for the first - a worklog's links,
    // comments and work done, a device's comments, links and photos, held or saved (code review, 2026-11-08: ten copies of
    // this one line). A list's own ids are only told apart within it, so one freed by a delete may come back.
    //
    // NOT the workbook, worklog and device ids (WorklogManager): those name folders and exported files, and are never
    // handed out again - a counter of their own.
    // ###########################################################################################
    internal static class RecordIds
    {
        public static int Next(IEnumerable<int> ids) => ids.DefaultIfEmpty(0).Max() + 1;
    }
}
