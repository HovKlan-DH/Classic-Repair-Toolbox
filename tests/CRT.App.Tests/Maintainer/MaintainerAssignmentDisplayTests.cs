using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers MaintainerAssignmentDisplay - how the administrator's "Maintainers" screen reads, and
// ReviewApiParser's two admin lists (Phase 6 roles, 2026-09-25).
//
// The wording is tested rather than eyeballed because this screen is where publish rights are
// handed out: a line that hid WHY an account cannot be granted, or a system whose empty pool read
// as a blank, would each send the administrator down the wrong path.
public sealed class MaintainerAssignmentDisplayTests
{
    private static MaintainerRow Anna() => new(1, "Anna", "anna@example.com");

    private static ReviewSystemRow System(params MaintainerRow[] maintainers) =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", "2026-May-14", true, maintainers);

    private static ReviewAccountRow Account(
        bool administrator = false, bool verified = true, bool locked = false) =>
        new(7, "bob@example.com", "Bob", administrator, verified, locked);

    // -----------------------------------------------------------------------------------
    // Systems
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_system_line_names_the_board_in_words_and_counts_its_maintainers()
    {
        Assert.Equal(
            "Commodore / C64 / 250407  -  1 maintainer",
            MaintainerAssignmentDisplay.SystemLine(MaintainerAssignmentDisplayTests.System(MaintainerAssignmentDisplayTests.Anna())));
    }

    [Fact]
    public void A_system_with_NOBODY_assigned_says_so_in_words()
    {
        // An unassigned board is exactly what the administrator opens this screen to find, and
        // "0 maintainers" reads as a count while "nobody assigned" reads as a to-do.
        Assert.EndsWith("nobody assigned", MaintainerAssignmentDisplay.SystemLine(MaintainerAssignmentDisplayTests.System()));
    }

    [Fact]
    public void Two_or_more_maintainers_pluralise()
    {
        Assert.Equal("2 maintainers", MaintainerAssignmentDisplay.CountPhrase(2));
        Assert.Equal("1 maintainer", MaintainerAssignmentDisplay.CountPhrase(1));
    }

    [Fact]
    public void A_system_with_no_name_parts_falls_back_to_its_id()
    {
        var bare = new ReviewSystemRow("X/Y/Z", "", "", "", null, true, []);

        Assert.StartsWith("X/Y/Z", MaintainerAssignmentDisplay.SystemLine(bare));
    }

    // -----------------------------------------------------------------------------------
    // People
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_maintainer_line_carries_the_address_because_names_are_not_unique()
    {
        Assert.Equal("Anna (anna@example.com)", MaintainerAssignmentDisplay.MaintainerLine(MaintainerAssignmentDisplayTests.Anna()));
    }

    [Fact]
    public void A_grantable_account_is_listed_plainly()
    {
        Assert.Equal("Bob (bob@example.com)", MaintainerAssignmentDisplay.AccountChoice(MaintainerAssignmentDisplayTests.Account()));
        Assert.Null(MaintainerAssignmentDisplay.WhyNotGrantable(MaintainerAssignmentDisplayTests.Account()));
    }

    [Fact]
    public void An_account_the_server_would_REFUSE_says_why_in_the_list_itself()
    {
        // Read before the button is pressed, not after a round trip - and the same three reasons
        // MaintainerAssignmentRules gives on the server.
        Assert.Contains("administrator", MaintainerAssignmentDisplay.AccountChoice(MaintainerAssignmentDisplayTests.Account(administrator: true)), StringComparison.Ordinal);
        Assert.Contains("not verified", MaintainerAssignmentDisplay.AccountChoice(MaintainerAssignmentDisplayTests.Account(verified: false)), StringComparison.Ordinal);
        Assert.Contains("locked", MaintainerAssignmentDisplay.AccountChoice(MaintainerAssignmentDisplayTests.Account(locked: true)), StringComparison.Ordinal);
    }

    // -----------------------------------------------------------------------------------
    // The parsers behind the screen
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_systems_answer_yields_each_system_with_its_maintainers()
    {
        ReviewSystemsResponse? systems = ReviewApiParser.ParseSystems("""
            {"systems":[
              {"systemId":"Commodore/C64/250407","manufacturer":"Commodore","hardware":"C64","board":"250407",
               "currentRevision":"2026-May-14","isAccepting":true,
               "maintainers":[{"accountId":1,"displayName":"Anna","email":"anna@example.com"}]},
              {"systemId":"Amstrad/CPC/464","manufacturer":"Amstrad","hardware":"CPC","board":"464",
               "currentRevision":null,"isAccepting":true,"maintainers":[]}]}
            """);

        Assert.NotNull(systems);
        Assert.Equal(2, systems!.Systems.Count);

        ReviewSystemRow c64 = systems.Systems[0];
        Assert.Equal("Commodore/C64/250407", c64.SystemId);
        Assert.Equal("2026-May-14", c64.CurrentRevision);
        Assert.Equal("Anna", Assert.Single(c64.Maintainers).DisplayName);

        Assert.Empty(systems.Systems[1].Maintainers);
        Assert.Null(systems.Systems[1].CurrentRevision);
    }

    [Fact]
    public void The_accounts_answer_yields_each_account_with_its_three_flags()
    {
        ReviewAccountsResponse? accounts = ReviewApiParser.ParseAccounts("""
            {"accounts":[
              {"id":1,"email":"admin@example.com","displayName":"Admin","isAdministrator":true,"isVerified":true,"isLocked":false},
              {"id":2,"email":"new@example.com","displayName":"New","isAdministrator":false,"isVerified":false,"isLocked":false}]}
            """);

        Assert.NotNull(accounts);
        Assert.Equal(2, accounts!.Accounts.Count);
        Assert.True(accounts.Accounts[0].IsAdministrator);
        Assert.False(accounts.Accounts[1].IsVerified);
    }

    [Fact]
    public void An_unreadable_answer_is_NULL_and_a_bad_row_is_skipped_rather_than_failing_the_list()
    {
        Assert.Null(ReviewApiParser.ParseSystems("not json"));
        Assert.Null(ReviewApiParser.ParseAccounts("""{"nope":[]}"""));

        ReviewSystemsResponse? systems = ReviewApiParser.ParseSystems("""
            {"systems":[{"manufacturer":"no id"},{"systemId":"A/B/C"},"a string"]}
            """);

        Assert.Equal("A/B/C", Assert.Single(systems!.Systems).SystemId);
    }
}
