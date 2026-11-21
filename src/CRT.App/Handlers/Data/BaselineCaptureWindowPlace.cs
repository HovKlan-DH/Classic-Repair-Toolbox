namespace Handlers.DataHandling
{
    // "Capture from oscilloscope pin-by-pin"'s window as it was left: "Normal" or "Maximized", and its Normal size and position
    // (UserSettings.BaselineCaptureWindowPlace; owner request, 2026-11-20: "remember size and position").
    public sealed record BaselineCaptureWindowPlace(string State, double Width, double Height, int X, int Y);
}
