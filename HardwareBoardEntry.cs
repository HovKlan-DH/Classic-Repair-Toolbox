namespace CRT
{
    // ###########################################################################################
    // Represents a single hardware/board entry from the main Excel data file.
    // Used to populate hardware and board drop-down selectors in the UI.
    // ###########################################################################################
    public class HardwareBoardEntry
    {
        public string HardwareName { get; init; } = string.Empty;
        public string BoardName { get; init; } = string.Empty;
        public string ExcelDataFile { get; init; } = string.Empty;
        public string HardwareNotes { get; init; } = string.Empty;

        public override string ToString() => $"{HardwareName} - {BoardName}";
    }
}