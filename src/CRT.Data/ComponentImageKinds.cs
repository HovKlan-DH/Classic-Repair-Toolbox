namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A "Component images" ROW IS - the one definition (CRT 4.0, 2026-10-24).
    //
    // A row that names a pin AND supplies at least one scope setting (T/DIV, V/DIV, T.LVL) is an
    // oscilloscope BASELINE; every other row is a picture of the component. It is the rule the
    // component popup has always used (ComponentImageQueries.IsOscilloscopeImage asks here), and
    // the board data converter splits the sheet by it into pictures and scope baselines - so the
    // popup and the converted data can never disagree about which rows are baselines.
    // ###########################################################################################
    public static class ComponentImageKinds
    {
        public static bool IsScopeBaseline(ComponentImageEntry? entry) =>
            entry != null &&
            !string.IsNullOrWhiteSpace(entry.Pin) &&
            (!string.IsNullOrWhiteSpace(entry.TimeDiv) ||
             !string.IsNullOrWhiteSpace(entry.VoltsDiv) ||
             !string.IsNullOrWhiteSpace(entry.TriggerLevelVolts));
    }
}
