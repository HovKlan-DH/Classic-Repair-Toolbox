using System.Xml.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// EVERY ENTRY THAT MOVES UP AND DOWN ON THE "Contribute" TAB STANDS IN A FRAME OF ITS OWN (owner request, 2026-11-08:
// "everything that can be moved up/down must have its own panel with a border. E.g. the "Files" and "Links" in both
// board and component can be moved, so make sure each entry is having its own panel, so it makes it easier to view its
// boundaries"). What moves is what has a drag handle (Border.RowDragHandle, ListRowDrag), and its frame is
// Border.MovableRow - so a list made movable later is held to it as soon as it gets its handle.
// ###########################################################################################
public sealed class ContributeMovableRowsTests
{
    private static readonly XNamespace Avalonia = "https://github.com/avaloniaui";

    [Fact]
    public void Every_drag_handle_on_the_Contribute_tab_stands_in_a_frame_of_its_own()
    {
        var handles = new List<string>();
        var unframed = new List<string>();

        foreach (string file in ContributeMovableRowsTests.ContributeMarkup())
        {
            XDocument markup = XDocument.Load(file, LoadOptions.SetLineInfo);

            foreach (XElement handle in markup.Descendants().Where(element => ContributeMovableRowsTests.HasClass(element, "RowDragHandle")))
            {
                string where = $"{Path.GetFileName(file)}:{((System.Xml.IXmlLineInfo)handle).LineNumber}";
                handles.Add(where);

                // Its frame lies between it and the template it is a row of.
                bool framed = handle.Ancestors()
                    .TakeWhile(ancestor => ancestor.Name != Avalonia + "DataTemplate")
                    .Any(ancestor => ancestor.Name == Avalonia + "Border" && ContributeMovableRowsTests.HasClass(ancestor, "MovableRow"));

                if (!framed)
                {
                    unframed.Add(where);
                }
            }
        }

        // The seven lists of 2026-11-08: a component's files and links, the board's files, links, credits, important
        // signals and categories. Fewer found means the detector has gone blind.
        Assert.True(handles.Count >= 7, $"Only {handles.Count} drag handles found - the detector is broken");
        Assert.True(unframed.Count == 0, "A row that moves stands in a Border.MovableRow:\n" + string.Join("\n", unframed));
    }

    private static bool HasClass(XElement element, string name) =>
        (element.Attribute("Classes")?.Value ?? string.Empty)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Contains(name, StringComparer.Ordinal);

    private static IEnumerable<string> ContributeMarkup() =>
        Directory.EnumerateFiles(ContributeMovableRowsTests.RepositoryPath(Path.Combine("src", "CRT.App", "Tabs", "Drafts")), "*.axaml");

    // Walks up from the test binary to the repository root (the folder holding the solution).
    private static string RepositoryPath(string relative)
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
        {
            folder = Path.GetDirectoryName(folder);
        }

        Assert.NotNull(folder);
        return Path.Combine(folder!, relative);
    }
}
