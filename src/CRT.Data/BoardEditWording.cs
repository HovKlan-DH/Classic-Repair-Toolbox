namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHY A BOARD CANNOT BE CHANGED ON THE BOARDS SCREEN WHILE THIS ACCOUNT'S LAST CHANGE OF IT
    // STILL WAITS. A change made there that the server could not publish stays in the queue as a
    // submission, and a newer submission from the same account would replace it - so the table
    // opens read-only, naming the submission to decide.
    //
    // CRT.Server sends it as the table's read-only reason (BoardEditFlow.WhyNotEditableAsync). The
    // Maintainer tab says it too when it reopens the table on such a change without reading it
    // again (BoardDetailView.ReopenAfterSentAsync; code review, 2026-10-10) - in CRT.Data so the
    // two are the same words.
    // ###########################################################################################
    public static class BoardEditWording
    {
        public static string AlreadyWaitingMessage(long submissionId) =>
            $"Your earlier change to this board, submission #{submissionId}, is still waiting under {MaintainerScreenWording.ContributorQueueQuoted}. " +
            "Approve or reject it there first - or make this change in its table there. A new change from here would replace it.";
    }
}
