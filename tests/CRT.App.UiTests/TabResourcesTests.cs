using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "Resources" tab shows a board file's and a board link's NOTE under its name (owner request,
// 2026-11-06: "I would like to add an additional 'Note' ... describe what this link or file is about.
// It should of course also show in the 'Resources' tab"), its lines kept; a row without one shows none.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabResourcesTests
{
    // A block's words - the plain-text fast path sets Text, a note with a web address Inlines.
    private static string TextOf(TextBlock block) =>
        block.Text ?? string.Concat(block.Inlines?.OfType<Run>().Select(run => run.Text) ?? []);

    private static List<TextBlock> NotesIn(ItemsControl list) =>
        list.GetLogicalDescendants().OfType<TextBlock>().Where(block => block.IsSet(Handlers.Theming.TextLinkRenderer.PlainLinkTextProperty)).ToList();

    // ###########################################################################################
    // A note is typed into a plain box, so it is shown as typed - its links clickable, nothing read as formatting (code
    // review, 2026-11-07: "use __only__ with rev A" showed underlined with its underscores gone, a line starting "- " as a
    // bullet - neither of which the box it was typed into showed).
    // ###########################################################################################
    [Fact]
    public void A_note_is_shown_as_typed_with_its_links_clickable()
    {
        UiTest.Run(() =>
        {
            var tab = new TabResources();
            var window = new Window { Content = tab, Width = 900, Height = 700 };
            window.Show();

            const string Note = "use __only__ with rev A\n- pins *4* and *5*, see https://x.dk/a__b__";
            tab.LoadData([new ResourceItem("Manuals", "Service manual", "manual.pdf", note: Note)], []);
            Dispatcher.UIThread.RunJobs();

            TextBlock note = Assert.Single(NotesIn(tab.LocalFilesItemsControl));
            Assert.Equal(Note, TextOf(note));

            List<Run> runs = note.Inlines!.OfType<Run>().ToList();
            Assert.Equal("https://x.dk/a__b__", Assert.Single(runs, run => run.TextDecorations is not null).Text);
            Assert.All(runs, run => Assert.NotEqual(Avalonia.Media.FontStyle.Italic, run.FontStyle));

            window.Close();
        });
    }

    [Fact]
    public void A_file_or_a_link_shows_its_note_under_its_name_and_one_without_shows_none()
    {
        UiTest.Run(() =>
        {
            var tab = new TabResources();
            var window = new Window { Content = tab, Width = 900, Height = 700 };
            window.Show();

            tab.LoadData(
                [
                    new ResourceItem("Manuals", "Service manual", "manual.pdf", note: "Page 12 has the PLA.\nThe rest is the C64C."),
                    new ResourceItem("Manuals", "Dead test", "dead.pdf"),
                ],
                [new ResourceItem("Forums", "Lemon64", "https://lemon64.com", note: "Search for 250407")]);

            Dispatcher.UIThread.RunJobs();

            List<TextBlock> fileNotes = NotesIn(tab.LocalFilesItemsControl);
            List<TextBlock> linkNotes = NotesIn(tab.WebLinksItemsControl);

            Assert.Equal(2, fileNotes.Count);
            Assert.Equal("Page 12 has the PLA.\nThe rest is the C64C.", TextOf(fileNotes[0]));
            Assert.True(fileNotes[0].IsVisible);
            Assert.False(fileNotes[1].IsVisible);

            Assert.Equal("Search for 250407", TextOf(Assert.Single(linkNotes)));
            Assert.True(linkNotes[0].IsVisible);

            window.Close();
        });
    }
}
