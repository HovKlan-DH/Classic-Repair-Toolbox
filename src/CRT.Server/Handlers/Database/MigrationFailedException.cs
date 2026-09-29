namespace CRT.Server.Handlers.Database
{
    // ###########################################################################################
    // A migration that failed, or a set of migrations MigrationPlan refused (2026-09-27).
    //
    // Its own type so Program.Main can tell it from a database that simply cannot be reached yet
    // (ServerExitCodes.ForStartupFailure): this one needs a person to look at the database, and
    // restarting the service only repeats it; an unreachable database is fixed by waiting. Derived
    // from InvalidOperationException, which is what MigrationRunner threw before it had a type.
    // ###########################################################################################
    public sealed class MigrationFailedException : InvalidOperationException
    {
        public MigrationFailedException(string message)
            : base(message)
        {
        }

        public MigrationFailedException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
