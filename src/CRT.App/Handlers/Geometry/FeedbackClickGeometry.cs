using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Which row of the Feedback tab's field a click on its empty space is on (code review,
    // 2026-10-16: a click beside a picture in the middle of the text sent the caret to the very
    // end). The parts and pictures are the rows, top to bottom, as (Top, Bottom): a click within
    // one is on it, a click in the gap between two on the nearer (the upper on a tie), above them
    // all on the first and below them all on the last. -1 when there are no rows.
    // ###########################################################################################
    internal static class FeedbackClickGeometry
    {
        public static int RowAt(double y, IReadOnlyList<(double Top, double Bottom)> rows)
        {
            ArgumentNullException.ThrowIfNull(rows);

            for (int i = 0; i < rows.Count; i++)
            {
                if (y < rows[i].Top)
                    return i == 0 || rows[i].Top - y < y - rows[i - 1].Bottom ? i : i - 1;

                if (y <= rows[i].Bottom)
                    return i;
            }

            return rows.Count - 1;
        }
    }
}
