namespace Handlers.DataHandling
{
    // A window as it was left: "Normal" or "Maximized", and its Normal size and position - "Capture from oscilloscope pin-by-pin"'s
    // (UserSettings.BaselineCaptureWindowPlace; owner request, 2026-11-20: "remember size and position") and a collection's
    // (UserSettings.CollectionWindowPlace; owner request, 2026-11-28: "do save its size and position, so it will open same size and
    // place the next time"). Kept and used by CRT's KeptWindowPlace. BaselineCaptureWindowPlace until 2026-11-28.
    public sealed record WindowPlace(string State, double Width, double Height, int X, int Y);
}
