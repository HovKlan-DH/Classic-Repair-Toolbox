using Avalonia.Input;

namespace CRT;

// ###########################################################################################
// The MODE HINT: the khaki "what to do next" label pinned to the window's upper-left corner,
// shown while a mode is waiting for the user to act on the schematic image.
//
// The problem it solves is general rather than specific to the worklog. Clicking "Add worklog"
// puts the Schematics tab into area-marking mode and then says nothing - the cursor becomes a
// crosshair and the user is expected to know they should now drag a rectangle. The label editor
// and KiCad calibration mode leave the user in the same position. This gives any such mode one
// line to say what it is waiting for.
//
// It lives in the main window rather than in the tab because it floats above the WHOLE window
// (the LeftPanel sidebar included, not just the tab area), and because a hint is about the
// application's state - which mode is active - rather than about anything the tab is drawing.
//
// Deliberately NOT dismissible and NOT remembered in settings. It clears itself on the first
// pointer press anywhere, so it costs an experienced user one click they were going to make
// anyway, and there is no stuck state and no "reset hints" control to build.
//
// To use it from a new mode: call ShowModeHint("...") when the mode starts and HideModeHint()
// when it ends. Nothing else is needed.
// ###########################################################################################
public partial class Main
{
    // The wording for the worklog area-marking mode. Kept beside the machinery that shows it so
    // the text a user reads is not scattered across the modes that trigger it.
    public const string WorklogAreaModeHint =
        "Now mark an area on the schematics image, to select the components in scope of your worklog.";

    // ###########################################################################################
    // Shows the hint. Calling it again just replaces the wording, so a mode that changes what it
    // is waiting for can update the line in place.
    // ###########################################################################################
    public void ShowModeHint(string text)
    {
        this.ModeHintText.Text = text;
        this.ModeHintBorder.IsVisible = true;
    }

    // ###########################################################################################
    // Hides the hint. Safe to call when it is not showing, so a mode's exit path can call it
    // unconditionally without first asking whether it ever appeared.
    // ###########################################################################################
    public void HideModeHint()
    {
        this.ModeHintBorder.IsVisible = false;
    }

    // ###########################################################################################
    // Clears the hint on the first pointer press anywhere in the window.
    //
    // Handled on the TUNNELLING route, so it runs before the press reaches whatever was clicked and
    // cannot be swallowed by a control that marks the event handled - the schematic image itself
    // does exactly that when a drag begins, which is the most likely first click of all.
    //
    // It does NOT mark the event handled: this only hides a label, and the click must still do
    // whatever the user actually pressed it for.
    // ###########################################################################################
    private void OnModeHintDismissPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (this.ModeHintBorder.IsVisible)
        {
            this.HideModeHint();
        }
    }
}
