using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE "My Devices" TAB LISTS - the Configuration tab's "Scope in "My Devices" tab" (owner request,
    // 2026-11-06: "This should now follow devices and not workbooks, so the two choices should be: Show all devices,
    // Show only devices for selected board. And then add a new option (in the middle) for 'Show only devices for
    // selected hardware'"; and asked, "Follow the same scope" - the workbook list narrows as the devices do):
    //   - AllBoards       every device and every workbook;
    //   - CurrentHardware those on any board of the hardware shown;
    //   - CurrentBoard    those on the board shown - the default, as the workbooks always were.
    //
    // Stored where the workbooks' scope always was, `workbooksScope` in the settings file, which CRT 3.x reads too:
    // it knows "AllBoards" and takes anything else as the board shown, so "CurrentHardware" is safe there.
    // A board key is "Hardware|Board", capitals aside, as every board-key lookup of a device is.
    // ###########################################################################################
    public static class MyDevicesScope
    {
        public const string AllBoards = "AllBoards";

        public const string CurrentHardware = "CurrentHardware";

        public const string CurrentBoard = "CurrentBoard";

        // A stored or handed value as one of the three - anything else is the board shown, the default.
        public static string Normalise(string? scope) =>
            string.Equals(scope, AllBoards, StringComparison.OrdinalIgnoreCase) ? AllBoards
            : string.Equals(scope, CurrentHardware, StringComparison.OrdinalIgnoreCase) ? CurrentHardware
            : CurrentBoard;

        // Whether something on the board `boardKey` names is listed while `boardShown` is shown.
        public static bool Holds(string? scope, string? boardKey, string? boardShown) => Normalise(scope) switch
        {
            AllBoards => true,
            CurrentHardware => BoardKeyParts.Split(boardKey).Hardware is { Length: > 0 } hardware &&
                               string.Equals(hardware, BoardKeyParts.Split(boardShown).Hardware, StringComparison.OrdinalIgnoreCase),
            _ => (boardKey?.Trim() ?? string.Empty).Length > 0 &&
                 string.Equals(boardKey!.Trim(), boardShown?.Trim(), StringComparison.OrdinalIgnoreCase),
        };
    }
}
