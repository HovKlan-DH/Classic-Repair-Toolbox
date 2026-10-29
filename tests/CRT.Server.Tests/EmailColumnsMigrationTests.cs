using System.Text.RegularExpressions;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Holds every EMAIL COLUMN the migrations declare to the longest address the server takes
    // (EmailAddressRules.MaximumLength). Every flow test runs against in-memory stores, so a column
    // narrower than the rule passed the whole suite, and a long address that passed the check failed
    // the INSERT on the server with a 500 - 0024's contact_email was 255 (code review, 2026-10-29).
    // This reads the migrations, as MaintainerRoleMigrationTests does, and fails on either half moving
    // alone.
    // ###########################################################################################
    public sealed class EmailColumnsMigrationTests
    {
        // Every migration in order, with "--" comments removed.
        private static string MigrationsText() =>
            Regex.Replace(
                string.Join(
                    "\n",
                    Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Migrations"), "*.sql")
                        .OrderBy(path => path, StringComparer.Ordinal)
                        .Select(File.ReadAllText)),
                "--[^\n]*",
                string.Empty);

        [Fact]
        public void Every_email_column_holds_the_longest_address_the_server_takes()
        {
            MatchCollection columns = Regex.Matches(
                EmailColumnsMigrationTests.MigrationsText(),
                @"\b(?<name>\w*email\w*)\s+VARCHAR\s*\(\s*(?<width>\d+)\s*\)",
                RegexOptions.IgnoreCase);

            Assert.NotEmpty(columns);

            Assert.All(columns, column => Assert.True(
                int.Parse(column.Groups["width"].Value) >= EmailAddressRules.MaximumLength,
                $"{column.Groups["name"].Value} is VARCHAR({column.Groups["width"].Value}); an address may be {EmailAddressRules.MaximumLength} characters."));
        }
    }
}
