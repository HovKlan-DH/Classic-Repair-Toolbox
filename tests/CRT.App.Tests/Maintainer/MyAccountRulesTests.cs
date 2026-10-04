using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// MyAccountRules - what "My account" lets the maintainer send (2026-10-03; the "Your account"
// window until 2026-10-04). A button is on only when pressing it could change something;
// everything the server judges is left to the server, so these never restate a password or name
// rule. No current password is asked for anywhere (owner decision, 2026-10-03).
// ###########################################################################################
public sealed class MyAccountRulesTests
{
    // ###########################################################################################
    // "Logged in as:<br /><b>Dennis</b> (dennis@...dk)" (owner request, 2026-10-04): the name bold
    // and the address after it in brackets, not bold - or, with no name, the address alone and bold,
    // since it is then what names the account.
    // ###########################################################################################
    [Fact]
    public void Logged_in_as_puts_the_name_in_bold_and_the_address_after_it()
    {
        Assert.Equal(
            [new ReviewNoteRun("Dennis", true), new ReviewNoteRun(" (dh@example.com)", false)],
            MyAccountRules.LoggedInAs(" Dennis ", "dh@example.com"));

        Assert.Equal([new ReviewNoteRun("dh@example.com", true)], MyAccountRules.LoggedInAs("  ", "dh@example.com"));
        Assert.Equal([new ReviewNoteRun("dh@example.com", true)], MyAccountRules.LoggedInAs(null, "dh@example.com"));
    }

    [Theory]
    [InlineData("Dennis", "Dennis H", true)]
    [InlineData("Dennis", "  Dennis H  ", true)]
    [InlineData("Dennis", "Dennis", false)]
    [InlineData("Dennis", "  Dennis ", false)]
    [InlineData("Dennis", "   ", false)]
    [InlineData("Dennis", null, false)]
    [InlineData("Dennis", "dennis", true)]
    public void A_name_can_be_saved_when_it_is_typed_and_different(string current, string? typed, bool expected)
    {
        Assert.Equal(expected, MyAccountRules.CanSaveName(current, typed));
    }

    // The same address in other capitals IS a change - the server makes it at once, with no code.
    [Theory]
    [InlineData("bench@example.com", true)]
    [InlineData("DH@example.com", true)]
    [InlineData("dh@example.com", false)]
    [InlineData(" dh@example.com ", false)]
    [InlineData("  ", false)]
    [InlineData(null, false)]
    public void A_code_can_be_asked_for_any_different_address(string? typed, bool expected)
    {
        Assert.Equal(expected, MyAccountRules.CanAskForCode("dh@example.com", typed));
    }

    [Fact]
    public void A_code_can_be_sent_back_once_one_is_typed()
    {
        Assert.True(MyAccountRules.CanConfirmCode("iAXr2z0PPffPwpmzHR-bOFo5ZPCPfHEK1hZbFAKAoYQ"));
        Assert.False(MyAccountRules.CanConfirmCode("   "));
        Assert.False(MyAccountRules.CanConfirmCode(null));
    }

    // ###########################################################################################
    // The two new passwords must agree before anything is sent - and the view says so only once
    // both are typed, since half of the second one is not a mistake yet.
    // ###########################################################################################
    [Fact]
    public void A_password_change_needs_the_new_one_twice_alike()
    {
        Assert.True(MyAccountRules.CanChangePassword("a new one here", "a new one here"));
        Assert.False(MyAccountRules.CanChangePassword("a new one here", "a new one her"));
        Assert.False(MyAccountRules.CanChangePassword("", ""));
        Assert.False(MyAccountRules.CanChangePassword(null, null));

        Assert.Equal(MyAccountRules.PasswordsDiffer, MyAccountRules.PasswordProblem("a new one here", "a new one her"));
        Assert.Null(MyAccountRules.PasswordProblem("a new one here", ""));
        Assert.Null(MyAccountRules.PasswordProblem("a new one here", "a new one here"));
    }
}
