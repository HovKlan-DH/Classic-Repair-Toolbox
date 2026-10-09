namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** ONLY THE ADMINISTRATOR PUBLISHES TO THE STABLE SOURCE, FOR NOW (owner request, 2026-10-05:
    // "I do not want to pollute the stable yet. Only me, as admin, should be able to publish to
    // stable"). ***
    //
    // While CRT.Server's ProductionPublishingAdministratorsOnly is on, the server refuses a
    // maintainer's publish from BETA to the stable source with this sentence, and sends it as the
    // plan's refusal - which the Maintainer tab shows under a greyed-out publish button. Pushing a
    // board back and rejecting it stay open to the maintainer, so the sentence says so.
    //
    // In CRT.Data so the server's refusal and the Maintainer tab's tests are the same words.
    // ###########################################################################################
    public static class StablePublishing
    {
        public const string AdministratorsOnlyMessage =
            "For now, only the administrator publishes to the stable source. If something is not right in BETA, " +
            "you can still push it back to the queue or reject it.";
    }
}
