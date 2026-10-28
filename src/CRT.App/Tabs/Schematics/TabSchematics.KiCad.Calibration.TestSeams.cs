using Avalonia.Input;
using Handlers.DataHandling;
using System;
using System.Threading.Tasks;

namespace CRT;

// ###########################################################################################
// Test seams for the KiCad trace calibration drag's POINTER CAPTURE.
//
// Same convention as TabSchematics.Worklog.TestSeams.cs: every member here delegates straight to
// the shipped field or method, nothing reimplements behaviour.
//
// WHY THESE EXIST. Entering calibration mode needs a decoded full-resolution bitmap, a loaded
// KiCad project and a resolvable view, none of which a headless test can stand up - so the drag
// itself is out of reach. What IS reachable, and what actually went wrong, is narrower: the press
// records the captured pointer, and every path that ends the drag has to release it. The release
// handler releases it only inside its "drag mode is not None" branch, and ESC and Apply both clear
// that mode while the button can still be down, after which the release falls through to the
// "nothing to tear down" early return and the capture outlives the gesture - every later pointer
// event then routes to SchematicsContainer even once the pointer has left it, recoverable only by
// switching board.
//
// So these seams pin the contract the fix rests on: a recorded capture is released and forgotten
// by whatever ends the drag, and calling the release with nothing captured is harmless.
//
// Since then calibration mode CAN be entered headlessly (UseKiCadProjectForTests, a stand-in bitmap),
// so its save has a stand-in too - "Done - next picture" (step 3, slice 3d) is pressed by a test.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // Stands in for the calibration's save into the draft: the row in, a failure out - or null for saved.
    internal Func<KiCadCalibrationEntry, Task<string?>>? KiCadCalibrationSaveOverrideForTests { get; set; }

    internal IPointer? KiCadTraceCalibrationCapturedPointerForTests =>
        this.thisKiCadTraceCalibrationCapturedPointer;

    internal void SetKiCadTraceCalibrationCapturedPointerForTests(IPointer? pointer) =>
        this.thisKiCadTraceCalibrationCapturedPointer = pointer;

    internal void ReleaseKiCadTraceCalibrationPointerCaptureForTests() =>
        this.ReleaseKiCadTraceCalibrationPointerCapture();
}
