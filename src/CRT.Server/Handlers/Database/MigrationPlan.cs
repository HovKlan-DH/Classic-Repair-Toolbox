using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace CRT.Server.Handlers.Database
{
    // ###########################################################################################
    // Decides WHICH migrations to run and IN WHAT ORDER. Pure: it takes the migrations found on
    // disk and the rows already recorded in schema_migrations, and returns a plan. It opens no
    // database, reads no file and logs nothing, so every rule below is a unit test.
    //
    // The I/O half - reading the .sql files, executing them, inserting the rows - is
    // MigrationRunner. That split is the same one the csproj header requires of endpoints, and for
    // the same reason: the decisions that can go wrong silently are the ones that must be testable
    // without a server.
    //
    // THE RULES, AND WHY EACH EXISTS:
    //
    // 1. Migrations apply in FILENAME order, and the filename must start with a zero-padded
    //    number. Ordering by anything else - a timestamp inside the file, the order the filesystem
    //    happens to return - means two machines can apply the same set in different orders and end
    //    up with different schemas from identical inputs.
    //
    // 2. A migration ALREADY APPLIED is never re-run, identified by its number, not its filename.
    //    Renaming 0002_accounts.sql to 0002_accounts_and_sessions.sql must not re-run it.
    //
    // 3. AN ALREADY-APPLIED MIGRATION WHOSE CONTENT CHANGED IS AN ERROR, not something to re-run
    //    and not something to ignore. The recorded checksum is what catches a migration edited
    //    after it shipped - the single most common way a team ends up with two databases that
    //    disagree while both claim to be at version N. The fix is always a NEW migration; this
    //    rule makes that non-optional. See VerifyExistingUnchanged.
    //
    // 4. A GAP IS AN ERROR. If 0001 and 0003 exist but 0002 does not, someone's migration did not
    //    get committed, or a merge dropped it. Applying 0003 onto a database missing 0002 produces
    //    a schema nobody has ever tested. Refuse instead.
    //
    // 5. A MIGRATION NUMBERED BELOW THE HIGHEST APPLIED ONE, BUT NOT ITSELF APPLIED, IS AN ERROR.
    //    This is the merge case: two branches each add "the next" migration, one lands first, and
    //    the other now wants to insert itself into the past. Silently applying it out of order is
    //    how a database ends up in a state the migration sequence cannot reproduce.
    //
    // Rules 3, 4 and 5 all fail the whole run rather than skipping the offending file. A partially
    // migrated database is worse than one that refused to start, for exactly the reason
    // ServerOptions gives for having no defaults: an outage is noticed, a silent divergence is not.
    // ###########################################################################################
    public static class MigrationPlan
    {
        // ###########################################################################################
        // Builds the plan, or the reasons there cannot be one.
        //
        // available: every migration found on disk, in any order - this method sorts them.
        // applied:   what schema_migrations already records.
        //
        // A plan with any failure has an EMPTY ToApply, so a caller that ignores Failures cannot
        // accidentally half-migrate. That is deliberate defensiveness: the honest caller checks
        // Failures first, and the careless one does nothing rather than damage.
        // ###########################################################################################
        public static MigrationPlanResult Build(
            IReadOnlyList<MigrationScript> available,
            IReadOnlyList<AppliedMigration> applied)
        {
            ArgumentNullException.ThrowIfNull(available);
            ArgumentNullException.ThrowIfNull(applied);

            var failures = new List<string>();

            MigrationPlan.VerifyNoDuplicateNumbers(available, failures);
            MigrationPlan.VerifyNoGaps(available, failures);
            MigrationPlan.VerifyExistingUnchanged(available, applied, failures);
            MigrationPlan.VerifyNoneInsertedIntoThePast(available, applied, failures);
            MigrationPlan.VerifyNothingAppliedIsMissing(available, applied, failures);

            if (failures.Count > 0)
                return new MigrationPlanResult(Array.Empty<MigrationScript>(), failures);

            var appliedNumbers = applied.Select(row => row.Number).ToHashSet();

            IReadOnlyList<MigrationScript> toApply = available
                .Where(script => !appliedNumbers.Contains(script.Number))
                .OrderBy(script => script.Number)
                .ToList();

            return new MigrationPlanResult(toApply, Array.Empty<string>());
        }

        // ###########################################################################################
        // The checksum recorded against an applied migration, and compared on every later run.
        //
        // SHA-256 of the UTF-8 bytes, with line endings normalised to \n first. The normalisation
        // is not cosmetic: these files are edited on Windows and executed on Linux, and git's
        // autocrlf can rewrite them in transit. Without it, every migration would appear modified
        // the first time the repository was checked out on the other platform - which would train
        // the project owner to ignore rule 3, the one rule most worth heeding.
        //
        // This is an integrity check against accidental edits, not a security control. Nobody is
        // defending against an attacker who can already rewrite migration files AND the
        // schema_migrations rows.
        // ###########################################################################################
        public static string ComputeChecksum(string sql)
        {
            ArgumentNullException.ThrowIfNull(sql);

            string normalised = sql.Replace("\r\n", "\n").Replace("\r", "\n");
            byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(normalised));

            return Convert.ToHexStringLower(hash);
        }

        // ###########################################################################################
        // Reads the leading number off a migration filename: "0001_initial.sql" -> 1.
        //
        // The underscore is required, so "0001.sql" is rejected: a migration with no descriptive
        // name is one nobody can identify in a schema_migrations listing months later.
        //
        // Returns false rather than throwing - the caller turns that into a named failure, which is
        // more useful than an exception from a directory walk.
        // ###########################################################################################
        public static bool TryParseNumber(string fileName, out int number)
        {
            number = 0;

            if (string.IsNullOrWhiteSpace(fileName))
                return false;

            int underscore = fileName.IndexOf('_');
            if (underscore <= 0)
                return false;

            string prefix = fileName[..underscore];

            // Every character must be a digit. int.TryParse would accept "-1" and " 1", both of
            // which would sort in ways nobody intends.
            foreach (char character in prefix)
            {
                if (!char.IsAsciiDigit(character))
                    return false;
            }

            return int.TryParse(prefix, NumberStyles.None, CultureInfo.InvariantCulture, out number)
                && number > 0;
        }

        // -------------------------------------------------------------------------------------
        // Rule checks. Each adds its own failures; none short-circuits, so one run reports every
        // problem - the same reasoning ServerOptionsValidator gives for accumulating failures.
        // -------------------------------------------------------------------------------------

        private static void VerifyNoDuplicateNumbers(
            IReadOnlyList<MigrationScript> available,
            List<string> failures)
        {
            foreach (var duplicate in available.GroupBy(script => script.Number).Where(group => group.Count() > 1))
            {
                string names = string.Join(", ", duplicate.Select(script => script.FileName).OrderBy(name => name));

                failures.Add(
                    $"Migration number {duplicate.Key} is used by more than one file ({names}). " +
                    "Each migration must have a unique number, because the number is what identifies " +
                    "it in schema_migrations.");
            }
        }

        private static void VerifyNoGaps(
            IReadOnlyList<MigrationScript> available,
            List<string> failures)
        {
            if (available.Count == 0)
                return;

            var numbers = available.Select(script => script.Number).Distinct().OrderBy(number => number).ToList();

            // Migrations start at 1 by convention: 0001_initial.sql. A set starting at 2 means the
            // first one is missing, which is a gap like any other.
            for (int expected = 1; expected <= numbers[^1]; expected++)
            {
                if (!numbers.Contains(expected))
                {
                    failures.Add(
                        $"Migration {expected:D4} is missing, but higher-numbered migrations exist " +
                        $"(up to {numbers[^1]:D4}). Applying those onto a database without it would " +
                        "produce a schema that no migration sequence can reproduce.");
                }
            }
        }

        private static void VerifyExistingUnchanged(
            IReadOnlyList<MigrationScript> available,
            IReadOnlyList<AppliedMigration> applied,
            List<string> failures)
        {
            var byNumber = available.ToLookup(script => script.Number);

            foreach (AppliedMigration row in applied)
            {
                MigrationScript? script = byNumber[row.Number].FirstOrDefault();

                // A recorded migration whose file is gone is handled by VerifyNothingAppliedIsMissing.
                if (script is null)
                    continue;

                if (!string.Equals(script.Checksum, row.Checksum, StringComparison.OrdinalIgnoreCase))
                {
                    failures.Add(
                        $"Migration {row.Number:D4} ({script.FileName}) has changed since it was " +
                        $"applied on {row.AppliedAtUtc:yyyy-MM-dd HH:mm} UTC. An applied migration is " +
                        "history and must never be edited - the database it already ran against " +
                        "cannot be changed retroactively. Revert the file and add a NEW migration " +
                        "with the change instead.");
                }
            }
        }

        private static void VerifyNoneInsertedIntoThePast(
            IReadOnlyList<MigrationScript> available,
            IReadOnlyList<AppliedMigration> applied,
            List<string> failures)
        {
            if (applied.Count == 0)
                return;

            int highestApplied = applied.Max(row => row.Number);
            var appliedNumbers = applied.Select(row => row.Number).ToHashSet();

            foreach (MigrationScript script in available.Where(script => !appliedNumbers.Contains(script.Number)))
            {
                if (script.Number < highestApplied)
                {
                    failures.Add(
                        $"Migration {script.Number:D4} ({script.FileName}) has never been applied, but " +
                        $"migration {highestApplied:D4} already has. Applying it now would run it out " +
                        "of order. This usually means two branches each added 'the next' migration - " +
                        "renumber this one above the highest applied migration.");
                }
            }
        }

        private static void VerifyNothingAppliedIsMissing(
            IReadOnlyList<MigrationScript> available,
            IReadOnlyList<AppliedMigration> applied,
            List<string> failures)
        {
            var availableNumbers = available.Select(script => script.Number).ToHashSet();

            foreach (AppliedMigration row in applied.Where(row => !availableNumbers.Contains(row.Number)))
            {
                failures.Add(
                    $"Migration {row.Number:D4} ({row.FileName}) is recorded as applied but its file is " +
                    "no longer present. Either the deployment is older than the database, or a " +
                    "migration was deleted. Deploying an older build over a newer database is the " +
                    "usual cause.");
            }
        }
    }

    // ###########################################################################################
    // A migration file found on disk. FileName is carried alongside Number purely so failures can
    // name the actual file the project owner has to go and look at.
    // ###########################################################################################
    public sealed record MigrationScript(int Number, string FileName, string Sql, string Checksum);

    // ###########################################################################################
    // A row from schema_migrations.
    // ###########################################################################################
    public sealed record AppliedMigration(int Number, string FileName, string Checksum, DateTime AppliedAtUtc);

    // ###########################################################################################
    // What Build decided. ToApply is empty whenever Failures is non-empty - see Build's header.
    // ###########################################################################################
    public sealed record MigrationPlanResult(
        IReadOnlyList<MigrationScript> ToApply,
        IReadOnlyList<string> Failures)
    {
        public bool CanProceed => this.Failures.Count == 0;
    }
}
