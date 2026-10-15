using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Contribute tab's three buttons (owner request, 2026-10-09): "Add new component", "Edit board
// as draft" and "Add a new board", in that order. The first two need a board shown; adding a new
// board never does - it is what someone does when theirs is not in the lists.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabContributeTests
{
    [Fact]
    public void Edit_board_as_draft_sits_between_the_other_two_and_needs_a_board_shown()
    {
        UiTest.Run(() =>
        {
            var tab = new TabContribute();

            Button addComponent = tab.GetControl<Button>("AddNewComponentButton");
            Button editAsDraft = tab.GetControl<Button>("EditBoardAsDraftButton");
            Button addBoard = tab.GetControl<Button>("AddNewBoardButton");

            var panel = (StackPanel)editAsDraft.Parent!;
            Assert.Equal([addComponent, editAsDraft, addBoard], panel.Children.OfType<Button>());
            Assert.Equal("Edit board as draft", editAsDraft.Content);

            tab.LoadData(null, "PAL");
            Assert.False(editAsDraft.IsEnabled);
            Assert.True(addBoard.IsEnabled);

            tab.LoadData(new BoardData(), "PAL");
            Assert.True(editAsDraft.IsEnabled);
            Assert.True(addComponent.IsEnabled);
        });
    }

    // ###########################################################################################
    // *** ALL THREE IN THE SAME COLOUR (owner request, 2026-10-09). *** Only "Add new component" had
    // it. IndianRed until 2026-10-13, then the green Accept colour, when IndianRed was kept for
    // buttons that throw something away (owner decision; App.axaml). Read off a SHOWN window, where
    // the styles are applied, against the theme's own Button_Ok_Bg.
    // ###########################################################################################
    [Fact]
    public void All_three_buttons_are_in_the_green_Accept_colour()
    {
        UiTest.Run(() =>
        {
            var tab = new TabContribute();
            var window = new Window { Content = tab, Width = 900, Height = 500 };
            window.Show();

            Assert.True(Avalonia.Application.Current!.TryGetResource(
                "Button_Ok_Bg", window.ActualThemeVariant, out object? resource));

            Avalonia.Media.Color accept = ((Avalonia.Media.ISolidColorBrush)resource!).Color;

            foreach (string name in new[] { "AddNewComponentButton", "EditBoardAsDraftButton", "AddNewBoardButton" })
            {
                Button button = tab.GetControl<Button>(name);

                Assert.Contains("ContributeAction", button.Classes);
                Assert.Contains("Accept", button.Classes);
                Assert.Equal(accept, ((Avalonia.Media.ISolidColorBrush)button.Background!).Color);
                Assert.NotEqual(Avalonia.Media.Colors.IndianRed, ((Avalonia.Media.ISolidColorBrush)button.Background!).Color);
            }

            window.Close();
        });
    }

    [Fact]
    public void A_problem_making_the_draft_is_said_under_the_text_and_cleared_again()
    {
        UiTest.Run(() =>
        {
            var tab = new TabContribute();
            TextBlock problem = tab.GetControl<TextBlock>("DraftProblemText");

            Assert.False(problem.IsVisible);

            tab.ShowDraftProblem("No draft could be made of this board: the published board could not be read.");
            Assert.True(problem.IsVisible);
            Assert.StartsWith("No draft could be made", problem.Text);

            tab.ShowDraftProblem(null);
            Assert.False(problem.IsVisible);
        });
    }

    // The problem is about the board it was pressed on: loading another board clears it (code
    // review, 2026-10-09 - it stayed under the next board until the button was pressed again).
    [Fact]
    public void A_problem_making_the_draft_is_cleared_when_another_board_is_loaded()
    {
        UiTest.Run(() =>
        {
            var tab = new TabContribute();
            tab.LoadData(new BoardData(), "PAL");
            tab.ShowDraftProblem("No draft could be made of this board: the published board could not be read.");

            tab.LoadData(new BoardData(), "PAL");

            Assert.False(tab.GetControl<TextBlock>("DraftProblemText").IsVisible);
        });
    }
}
