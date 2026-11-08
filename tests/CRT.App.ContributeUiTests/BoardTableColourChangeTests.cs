using Avalonia.Controls;
using Avalonia.LogicalTree;
using Avalonia.Media;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A changed color cell's card (owner request, 2026-11-01: "In 'Maintainer' tab: If the color has
// changed, then I would like a popup showing the colors 'Changed from' and 'Changed to'"): the two
// colors drawn beside their names, in the file hover card's own popup. Which cells, and their two
// sides, are pinned in CRT.Data.Tests' BoardTableColourCellsTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableColourChangeTests
{
    private static BoardSchematicEntry Picture(string color) =>
        new() { SchematicName = "Top", SchematicImageFile = "Images/top.png", SchematicHighlightColor = color };

    private static BoardTableDocument Document(string from, string to) =>
        BoardTableDocument.Create(new BoardData { Schematics = [Picture(from)] }, new BoardData { Schematics = [Picture(to)] });

    private static int ColourColumn(BoardTableSheet sheet) =>
        sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColSchematicHighlightColor);

    private static List<string?> TextsOf(Control card) =>
        card.GetLogicalDescendants().OfType<TextBlock>().Select(text => text.Text).ToList();

    [Fact]
    public void The_card_draws_both_colors_beside_their_names()
    {
        UiTest.Run(() =>
        {
            BoardTableSheet sheet = Document("DeepSkyBlue", "#FF0000").FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

            Control card = BoardTableEditor.BuildColourChangeCard(sheet.Rows.Single(), ColourColumn(sheet))!;

            Assert.Equal(["Changed from", "DeepSkyBlue", "Changed to", "#FF0000"], TextsOf(card));
            Assert.Equal(
                [Colors.DeepSkyBlue, Colors.Red],
                card.GetLogicalDescendants().OfType<Border>().Select(swatch => ((ISolidColorBrush)swatch.Background!).Color));
        });
    }

    // A side with no color of its own says so, its swatch empty - CRT draws it in its default.
    [Fact]
    public void A_color_not_set_says_so()
    {
        UiTest.Run(() =>
        {
            BoardTableSheet sheet = Document(string.Empty, "Red").FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

            Control card = BoardTableEditor.BuildColourChangeCard(sheet.Rows.Single(), ColourColumn(sheet))!;

            Assert.Equal(["Changed from", BoardTableEditor.ColourNotSet, "Changed to", "Red"], TextsOf(card));
            Assert.Equal(Colors.Transparent, ((ISolidColorBrush)card.GetLogicalDescendants().OfType<Border>().First().Background!).Color);
        });
    }

    [Fact]
    public void No_card_for_a_color_that_did_not_change_or_any_other_cell()
    {
        UiTest.Run(() =>
        {
            BoardTableSheet same = Document("Red", "Red").FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

            Assert.Null(BoardTableEditor.BuildColourChangeCard(same.Rows.Single(), ColourColumn(same)));
            Assert.Null(BoardTableEditor.BuildColourChangeCard(same.Rows.Single(), 0));
            Assert.Null(BoardTableEditor.BuildColourChangeCard(same.Rows.Single(), 999));
        });
    }

    // ###########################################################################################
    // It opens in the file card's popup - with no file source too, as it needs no bytes - and goes as
    // that does.
    // ###########################################################################################
    [Fact]
    public void The_card_opens_in_the_file_cards_popup_without_a_file_source()
    {
        UiTest.Run(() =>
        {
            BoardTableDocument document = Document("DeepSkyBlue", "Red");
            var editor = new BoardTableEditor();
            editor.Open(document);

            var window = new Window { Content = editor, Width = 900, Height = 600 };
            window.Show();

            BoardTableSheet sheet = document.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

            Assert.Null(editor.ShowFilePreview(sheet.Rows.Single(), ColourColumn(sheet), editor));

            Control shown = Assert.IsAssignableFrom<Decorator>(editor.FilePreviewFrameForTests).Child!;
            Assert.Contains("Changed to", TextsOf(shown));

            window.Close();
        });
    }
}
