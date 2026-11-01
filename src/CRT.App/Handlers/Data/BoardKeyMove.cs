namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THIS MACHINE KEEPS UNDER A BOARD'S NAME, CARRIED OVER TO ITS NEW ONE (a new board renamed on
    // the Contribute tab, owner request 2026-11-01): worklogs, traces, devices and settings are filed by
    // "{hardware}|{board}", not by the board's folder, so a board renamed would otherwise come back with
    // none of them. A part that could not follow is said in the log, and stays under the old name -
    // nothing is thrown away.
    // ###########################################################################################
    internal static class BoardKeyMove
    {
        public static void Run(string from, string to)
        {
            if (string.Equals(from, to, System.StringComparison.Ordinal))
            {
                return;
            }

            UserSettings.MoveBoard(from, to);

            if (!TraceStorage.MoveBoard(from, to))
            {
                Logger.Warning($"The drawn traces of [{from}] could not follow the board to [{to}]");
            }

            if (!WorklogManager.MoveBoard(from, to))
            {
                Logger.Warning($"Not every workbook of [{from}] could follow the board to [{to}]");
            }

            if (!OwnedDeviceStore.MoveBoard(from, to))
            {
                Logger.Warning($"Not every device on [{from}] could follow the board to [{to}]");
            }
        }
    }
}
