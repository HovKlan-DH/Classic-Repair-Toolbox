using Avalonia.Input;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Covers the KiCad trace calibration drag's pointer capture being RELEASED by whatever ends the
// drag - not only by the matching PointerReleased branch.
//
// The bug: OnSchematicsPointerReleased releases the capture inside its
// "thisKiCadTraceCalibrationDragMode != None" branch, but ESC (CancelKiCadTraceCalibrationMode)
// and Apply (ApplyKiCadTraceCalibration) both clear that mode while the button can still be down.
// The release that follows then falls through to the "not panning, nothing to tear down" early
// return, and the capture outlives the gesture: every later pointer event routes to
// SchematicsContainer even once the pointer has left it, recoverable only by switching board -
// the same class of leak the worklog-resize and worklog-drawing branches each carry a comment
// about having already had to fix.
//
// WHAT THESE DO NOT COVER. Entering calibration mode needs a decoded full-resolution bitmap, a
// loaded KiCad project and a resolvable view, so the drag itself is out of reach headlessly.
// These pin the contract the fix rests on instead - the recorded capture is released and
// forgotten, and releasing with nothing captured is harmless - via the seams in
// TabSchematics.KiCad.Calibration.TestSeams.cs.
// ###########################################################################################
[Collection("HeadlessUi")]
public class KiCadCalibrationPointerCaptureTests
{
    // Avalonia's own Pointer, not a fake: IPointer is explicitly not implementable outside the
    // framework (it carries a hidden member that says so), and Pointer is a plain public class
    // whose Capture/Captured need no window. Capturing to a real control and reading Captured back
    // is therefore the actual behaviour, not a stand-in for it.
    private static Pointer NewPointer() => new(1, PointerType.Mouse, isPrimary: true);

    [Fact]
    public void Releasing_the_calibration_capture_clears_it_and_releases_the_pointer()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var pointer = NewPointer();

            pointer.Capture(tab);
            tab.SetKiCadTraceCalibrationCapturedPointerForTests(pointer);
            Assert.Same(pointer, tab.KiCadTraceCalibrationCapturedPointerForTests);

            tab.ReleaseKiCadTraceCalibrationPointerCaptureForTests();

            Assert.Null(pointer.Captured);

            // Forgotten as well as released, so a later release cannot act on a pointer whose
            // gesture is long over.
            Assert.Null(tab.KiCadTraceCalibrationCapturedPointerForTests);
        });
    }

    // Every path that ends a calibration drag calls this unconditionally - including the ones that
    // end a mode in which no drag was ever started - so it has to be safe with nothing captured.
    [Fact]
    public void Releasing_with_nothing_captured_is_harmless()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            Assert.Null(tab.KiCadTraceCalibrationCapturedPointerForTests);

            tab.ReleaseKiCadTraceCalibrationPointerCaptureForTests();
            tab.ReleaseKiCadTraceCalibrationPointerCaptureForTests();

            Assert.Null(tab.KiCadTraceCalibrationCapturedPointerForTests);
        });
    }

    // The second release must not re-touch the first pointer: once forgotten, it is no longer this
    // tab's to act on, and a stale Capture(null) would tear down whatever captured it since.
    [Fact]
    public void A_second_release_does_not_touch_the_pointer_already_let_go()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var pointer = NewPointer();

            pointer.Capture(tab);
            tab.SetKiCadTraceCalibrationCapturedPointerForTests(pointer);
            tab.ReleaseKiCadTraceCalibrationPointerCaptureForTests();

            // Something else captures it afterwards - a pan, another control.
            var otherElement = new TabSchematics();
            pointer.Capture(otherElement);

            tab.ReleaseKiCadTraceCalibrationPointerCaptureForTests();

            Assert.Same(otherElement, pointer.Captured);
        });
    }
}
