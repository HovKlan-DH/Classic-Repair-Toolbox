using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A capture waiting in the user's folder is read for its picture OFF the UI thread (BaselineCaptureTarget.WaitingPictureRead) - and
// let go of ("Capture again", another component shown, the window closed) while that read may still be running. Windows refuses to
// delete a file another thread has open unless it was opened to allow that, and the capture then stayed in the user's folder for
// good (ComponentPaneBaselineRunTests failed on it at random under load, 2026-11-26).
// ###########################################################################################
public sealed class BaselineCaptureTargetTests
{
    [Fact]
    public void A_capture_being_read_for_its_picture_can_still_be_deleted()
    {
        using var workspace = new TempWorkspace();
        string capture = workspace.WriteFile("capture.png", "not a real picture");

        using (BaselineCaptureTarget.OpenCapture(capture))
        {
            File.Delete(capture);
            Assert.False(File.Exists(capture));
        }
    }
}
