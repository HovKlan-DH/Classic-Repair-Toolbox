using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE ORDER "My devices" AND THE WORKLOG BAR'S "Device" DROP-DOWN LIST THE DEVICES IN - the Configuration tab's
    // "Sorting of devices" (owner request, 2026-11-07: "Sort automatically by name", "Sort automatically by
    // hardware/board/name", "Sort manually in list (drag up/down)", by name the default):
    //   - ByName              the device's name, numbers as numbers ("#2" before "#10");
    //   - ByHardwareBoardName its hardware, then its board, then its name - each as the drop-downs name it;
    //   - Manual              the order dragged in "My devices" (OwnedDeviceRecord.DisplayOrder), a device never
    //                         dragged after those that were, in the order they were added.
    //
    // A tie - two devices of one name - goes by the order they were added, so the list never jumps between two reads.
    // ###########################################################################################
    public static class OwnedDeviceSorting
    {
        public const string ByName = "Name";

        public const string ByHardwareBoardName = "HardwareBoardName";

        public const string Manual = "Manual";

        // A stored value as one of the three - anything else is by name, the default.
        public static string Normalise(string? sorting) =>
            string.Equals(sorting, ByHardwareBoardName, StringComparison.OrdinalIgnoreCase) ? ByHardwareBoardName
            : string.Equals(sorting, Manual, StringComparison.OrdinalIgnoreCase) ? Manual
            : ByName;

        // ###########################################################################################
        // The devices in the order `sorting` asks for. `namesOf` gives a board key's hardware and board as the
        // drop-downs name them - what the "hardware/board" order reads, and only it.
        // ###########################################################################################
        public static List<OwnedDeviceRecord> Sort(
            IEnumerable<OwnedDeviceRecord> devices,
            string? sorting,
            Func<string, (string Hardware, string Board)> namesOf)
        {
            ArgumentNullException.ThrowIfNull(devices);
            ArgumentNullException.ThrowIfNull(namesOf);

            return Normalise(sorting) switch
            {
                Manual => devices
                    .OrderBy(device => device.DisplayOrder > 0 ? 0 : 1)
                    .ThenBy(device => device.DisplayOrder)
                    .ThenBy(device => device.Id)
                    .ToList(),

                ByHardwareBoardName => devices
                    .Select(device => (Device: device, Names: namesOf(device.BoardKey ?? string.Empty)))
                    .OrderBy(each => each.Names.Hardware?.Trim() ?? string.Empty, NaturalLabelComparer.Instance)
                    .ThenBy(each => each.Names.Board?.Trim() ?? string.Empty, NaturalLabelComparer.Instance)
                    .ThenBy(each => each.Device.Name?.Trim() ?? string.Empty, NaturalLabelComparer.Instance)
                    .ThenBy(each => each.Device.Id)
                    .Select(each => each.Device)
                    .ToList(),

                _ => devices
                    .OrderBy(device => device.Name?.Trim() ?? string.Empty, NaturalLabelComparer.Instance)
                    .ThenBy(device => device.Id)
                    .ToList(),
            };
        }

        // ###########################################################################################
        // The whole order once the devices LISTED were dragged into `listedNow`: the places the listed devices
        // stood in are theirs again, in their new order, and every device not listed - another board's, under the
        // scope, or one the search left out - keeps its place. `allInOrder` is every device in the manual order
        // (Sort's), `listedNow` the ids listed, as dropped. A listed id not in the whole order is left out.
        // ###########################################################################################
        public static List<int> Reordered(IReadOnlyList<int> allInOrder, IReadOnlyList<int> listedNow)
        {
            ArgumentNullException.ThrowIfNull(allInOrder);
            ArgumentNullException.ThrowIfNull(listedNow);

            var listed = new HashSet<int>(listedNow.Where(allInOrder.Contains));
            var next = new Queue<int>(listedNow.Where(listed.Contains).Distinct());
            var order = new List<int>(allInOrder.Count);

            foreach (int id in allInOrder)
            {
                order.Add(listed.Contains(id) && next.Count > 0 ? next.Dequeue() : id);
            }

            return order;
        }
    }
}
