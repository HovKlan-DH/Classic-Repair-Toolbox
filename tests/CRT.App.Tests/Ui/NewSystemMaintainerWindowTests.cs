using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// NewSystemMaintainerWindow - the agreement to become a new system's maintainer, which must be
// accepted before "Create system" creates anything (maintainer request, 2026-09-24).
//
// The keyboard tests go through the real headless input stack, the same way
// DiscardDraftWindowTests does, because the thing they guard is the Tunnel-vs-bubbling trap: a
// focused Button handles Enter itself, so a bubbling KeyDown handler would let a reflexive Enter
// on "Accept and create" accept an agreement nobody read. These fail against that version.
//
// Whether Create is actually gated on this window is NewSystemWindowTests' job - see its
// "maintainer agreement" section, which drives this dialog through the real create path.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class NewSystemMaintainerWindowTests
{
    private static NewSystemMaintainerWindow BuildWindow()
    {
        var window = new NewSystemMaintainerWindow();
        window.Initialize("HW4 - Board4");

        // Show() so the visual tree is built and focus can actually land on a button.
        window.Show();
        return window;
    }

    private static Button ButtonWithContent(NewSystemMaintainerWindow window, string content) =>
        window.GetVisualDescendants()
            .OfType<Button>()
            .First(b => (b.Content as string) == content);

    private static void PressKey(NewSystemMaintainerWindow window, Key key, PhysicalKey physicalKey) =>
        window.KeyPress(key, RawInputModifiers.None, physicalKey, keySymbol: null);

    [Fact]
    public void Enter_declines_even_when_the_accept_button_has_focus()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();
            var acceptButton = ButtonWithContent(window, "Accept and create");

            bool accepted = false;
            acceptButton.Click += (_, _) => accepted = true;

            acceptButton.Focus();
            PressKey(window, Key.Enter, PhysicalKey.Enter);

            Assert.False(accepted);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Escape_declines()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            bool accepted = false;
            ButtonWithContent(window, "Accept and create").Click += (_, _) => accepted = true;

            PressKey(window, Key.Escape, PhysicalKey.Escape);

            Assert.False(accepted);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void The_agreement_names_the_system_on_its_own_bold_line()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var nameBlock = window.GetControl<TextBlock>("SystemNameText");

            Assert.Equal("HW4 - Board4", nameBlock.Text);
            Assert.Equal(FontWeight.Bold, nameBlock.FontWeight);
        });
    }

    // What is being agreed to has to be on the screen: the role, what it involves, and that the
    // system is not created without it.
    [Fact]
    public void The_text_states_the_role_what_it_involves_and_that_it_is_required()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            string all = string.Join(
                "\n",
                window.GetVisualDescendants().OfType<TextBlock>().Select(t => t.Text ?? string.Empty));

            Assert.Contains("registered as its reviewer", all, StringComparison.Ordinal);
            Assert.Contains("reviewing the changes other members of the community submit", all, StringComparison.Ordinal);
            Assert.Contains("can only be created if you accept this", all, StringComparison.Ordinal);
        });
    }

    // Both choices are offered by name - an agreement with only an OK button is not a choice.
    [Fact]
    public void Both_Decline_and_Accept_are_offered()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            Assert.NotNull(ButtonWithContent(window, "Decline"));
            Assert.NotNull(ButtonWithContent(window, "Accept and create"));
        });
    }
}
