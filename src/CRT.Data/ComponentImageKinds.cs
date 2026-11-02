namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A "Component images" ROW IS - the one definition (CRT 4.0, 2026-10-24).
    //
    // A row that names a pin is an oscilloscope BASELINE; every other row is a picture of the
    // component. The component popup asks here (ComponentImageQueries.IsOscilloscopeImage), and the
    // board data converter splits the sheet by it into pictures and scope baselines - so the popup
    // and the converted data can never disagree about which rows are baselines.
    //
    // *** T/DIV, V/DIV AND T.LVL DO NOT DECIDE IT SINCE 2026-11-02 (owner request: "there is no
    // requirement that one of T/DIV or V/DIV or Trigger level must be set"). *** A row had to supply
    // one of them until then. Measured on the shipped data that day: all 33,390 rows naming a pin
    // have one of them too, so no published row changes kind.
    // ###########################################################################################
    public static class ComponentImageKinds
    {
        public static bool IsScopeBaseline(ComponentImageEntry? entry) =>
            entry != null && !string.IsNullOrWhiteSpace(entry.Pin);
    }
}
