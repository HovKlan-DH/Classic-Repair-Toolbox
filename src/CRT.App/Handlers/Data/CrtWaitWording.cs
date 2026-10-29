using System.Globalization;
using CRT;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What CRT's "please wait" overlay says, and what the user is told when a wait ran into the
    // two-minute limit (owner decision, 2026-09-28: "I want this method everywhere in the entire
    // project where there is a Wait ... it must be solid in validating if it did finish").
    //
    // Pure and tested. Each sentence names WHAT is happening, for a hobbyist at the bench rather than
    // for a log. The after-timeout sentences build on WaitWording, which the Maintainer tab uses
    // too, so the limit reads the same on every screen.
    //
    // *** SOME THINGS CANNOT BE CHECKED AFTERWARDS, AND SAY SO. *** Feedback goes to a page that
    // answers nothing back, so after a timeout nobody can know whether it arrived - the sentence
    // says it MAY have, and what to do, rather than guessing either way.
    // ###########################################################################################
    public static class CrtWaitWording
    {
        public const string SendingFeedback = "Sending your feedback...";

        public const string CheckingSubmissions = "Checking your contributions with the server...";

        public const string SavingDraft = "Saving to your local draft...";

        public const string SavingLabels = "Saving the component labels to your draft...";

        public const string OpeningTable = "Opening the table...";

        public const string OpeningComponent = "Opening the component...";

        // A picture, a screenshot or a file added in the component pane - copied in and saved.
        public const string AddingFile = "Adding the file to your draft...";

        public const string CapturingOscilloscope = "Taking the screenshot from the oscilloscope...";

        // "Copy from another board" (slice 5) - the other boards' files read for the same part.
        public const string LookingAtOtherBoards = "Looking for the same part on the other boards...";

        // "Also on other boards" for a file replaced - the other boards' file lists read for its bytes.
        public const string LookingForFilesOnOtherBoards = "Looking for the same files on the other boards...";

        // "Also on other boards" for a part AND its files - one read of the other boards for both.
        public const string LookingForPartAndFilesOnOtherBoards = "Looking for the same part and files on the other boards...";

        // "Also on other boards" (slice 5) - each board's draft made when it has none, changed and saved.
        public const string ChangingOtherBoards = "Putting the change into the other boards' drafts...";

        public const string MakingDraft = "Making your draft of this board...";

        // Before the first draft of a run, when nothing has asked yet (Main.UpdateRequired.cs).
        public const string CheckingDraftsWithServer = "Checking with the server that this version of CRT can send drafts...";

        public const string ComparingWithOfficial = "Comparing your draft with the published data...";

        // "Review changes", and the Submit and Discard draft dialogs, reading what the draft changes
        // (DraftChangeSummary) - step 1's last slice, 2026-10-25.
        public const string ReadingChanges = "Reading what your draft changes...";

        // Submit's credit suggestions are worked out against the board as published (code review, 2026-10-26).
        public const string ReadingPublishedBoard = "Reading the board as published...";

        public const string UndoingComponentChanges = "Putting the component back to its published values...";

        public const string BringingDraftUpToDate = "Bringing your draft up to date with the published data...";

        // The Contribute tab's "KiCad data", reading the board before its window opens (2026-10-10).
        public const string OpeningBoardFiles = "Reading the board...";

        public const string ImportingKiCad = "Importing the KiCad data...";

        public const string CheckingKiCadMatches = "Checking what the KiCad data lines up with...";

        // "Add N from KiCad" on a new board's checklist (step 3).
        public const string AddingKiCadComponents = "Adding the components from the KiCad data to your draft...";

        // "Mark all automatically" on a picture aligned with KiCad (step 3).
        public const string MarkingFromKiCad = "Drawing the components' rectangles from the KiCad data...";

        // A suggested match used on a new board's checklist (step 3).
        public const string MatchingPictures = "Matching the pictures to their KiCad views...";

        // Files dropped on the Contribute tab, looked at before their window opens (step 3).
        public const string ReadingDroppedFiles = "Looking at the files dropped...";

        // A device's photos copied in, and a device deleted with its photos (step 4).
        public const string AddingDevicePhotos = "Adding the photos to the device...";

        public const string DeletingDevice = "Deleting the device...";

        public const string ConnectingOscilloscope = "Connecting to the oscilloscope...";

        public const string DownloadingUpdate = "Downloading the update...";

        public static string ExportingWorkbook(bool asZip) =>
            asZip ? "Exporting the workbook to a ZIP file..." : "Exporting the workbook to a PDF...";

        public static string ExportingDeviceHistory(bool asZip) =>
            asZip ? "Exporting the device's history to a ZIP file..." : "Exporting the device's history to a PDF...";

        public static string CheckingSubmission(int number, int total) =>
            string.Create(CultureInfo.InvariantCulture, $"Checking your contributions with the server ({number} of {total})...");

        public static string DownloadingUpdateAt(int percent) =>
            string.Create(CultureInfo.InvariantCulture, $"Downloading the update: {percent}%");

        // ###########################################################################################
        // The feedback upload as it goes, then - once every byte is sent - what is waited for: the
        // server unpacks the attached files and mails them before it answers, which for a large zip
        // takes a while with nothing more to report (code review, 2026-10-04).
        // ###########################################################################################
        public static string SendingFeedbackAt(int percent) =>
            percent >= 100
                ? "Sent - waiting for the server to save the files and mail the feedback..."
                : string.Create(CultureInfo.InvariantCulture, $"Sending to server... {percent}%");

        // ---- After the limit ------------------------------------------------------------------

        // Feedback: nothing can be read back, so it may or may not have arrived. The text is kept.
        public static string FeedbackNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. Your feedback may still have arrived - " +
            "it is kept below, so wait a while and send it again only if you think it did not.";

        // My submissions: whatever was checked before the limit is on screen.
        public static string SubmissionsNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. The rows show what was checked before that, and nothing has been lost.";

        public static string OscilloscopeNoAnswer =>
            $"The oscilloscope did not answer within {WaitWording.Limit}. Check that it is switched on and on the network, then connect again.";

        public static string UpdateNoProgress =>
            $"The update download stopped moving for {WaitWording.Limit}. Check your internet connection and try again.";
    }
}
