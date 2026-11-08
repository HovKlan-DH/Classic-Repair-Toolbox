using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // HOW WIDE THE WORKLOG BAR'S DROP-DOWNS ARE (owner request, 2026-11-06: "please have the two drop-downs
    // resize, so it will show as much as the text in the drop-downs as possible. For now it looks like they are
    // statically set, with device being the smallest one"). Each wants the width of the line it shows; with room
    // for all, each gets it. Short of room, the room is shared as water fills: a box wanting less than an even
    // share keeps all of it and leaves the rest to the others, so the most text shows - a short device name whole
    // beside a long workbook title cut short, never both cut. Pure: BarPickersPanel lays the bar out by it.
    // ###########################################################################################
    internal static class BarPickersGeometry
    {
        // Each box's width, in the order `wanted` names them, for `room` to share.
        public static double[] Widths(double room, IReadOnlyList<double> wanted)
        {
            var widths = new double[wanted.Count];
            double left = Math.Max(0, room);

            List<int> leastFirst = Enumerable.Range(0, wanted.Count).OrderBy(index => wanted[index]).ToList();

            for (int place = 0; place < leastFirst.Count; place++)
            {
                int index = leastFirst[place];
                double evenShare = left / (leastFirst.Count - place);

                widths[index] = Math.Min(Math.Max(0, wanted[index]), evenShare);
                left -= widths[index];
            }

            return widths;
        }
    }
}
