using System.IO;
using System.Text;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A CRT 4.0 main file for a test's data root - the board list and the oscilloscopes, as
// "Classic-Repair-Toolbox.v4.0.0.json" (DataTreeLayout.MainFileName) - and the identity CRT gives a
// board it lists ("<id>/board.json", DataTreeLayout.BoardFilePath).
// ###########################################################################################
internal static class MainFiles
{
    public const string FileName = DataTreeLayout.MainFileName;

    // One board of the list: its id ("Commodore/C64/250407") and the names the drop-downs show.
    public sealed record Board(string BoardId, string HardwareName, string BoardName, string HardwareNotes = "")
    {
        public string Identity => DataTreeLayout.BoardFilePath(this.BoardId);
    }

    // Writes the main file listing `boards` under `dataRoot` and returns its file name.
    public static string Write(string dataRoot, params Board[] boards) =>
        MainFiles.Write(dataRoot, boards, []);

    public static string Write(string dataRoot, IEnumerable<Board> boards, IEnumerable<MainDocument.OscilloscopeRow> oscilloscopes)
    {
        var main = new MainDocument
        {
            Boards = boards.Select(board => new MainDocument.BoardListingRow
            {
                BoardId = board.BoardId,
                HardwareName = board.HardwareName,
                BoardName = board.BoardName,
                HardwareNotes = board.HardwareNotes,
            }).ToList(),
            Oscilloscopes = oscilloscopes.ToList(),
        };

        Directory.CreateDirectory(dataRoot);
        File.WriteAllText(Path.Combine(dataRoot, MainFiles.FileName), DataTreeJson.Write(main), new UTF8Encoding(false));

        return MainFiles.FileName;
    }
}
