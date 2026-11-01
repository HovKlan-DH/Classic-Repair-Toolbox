namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The words of "Board name" on the Contribute tab (owner request, 2026-11-01).
    // ###########################################################################################
    internal static class BoardNameWording
    {
        public const string Intro = "How the board is listed in the hardware and board lists at the top left.";

        // Above a published board's name. Its folder and name are how every CRT and the server find it.
        public const string Published = "This board is published, so its name cannot be changed here - it is how every CRT finds the board.";

        // Above a new board's name once a submission of it is with a maintainer.
        public const string SentForReview = "This board has been sent for review under this name, so it keeps it until the maintainer has decided.";

        public const string RenameMoves = "A new name moves your draft to a folder of that name, and clears its Undo steps.";

        public const string Saved = "Saved.";

        public const string NotSaved = "Nothing was changed - what the draft holds could not be saved first.";
    }
}
