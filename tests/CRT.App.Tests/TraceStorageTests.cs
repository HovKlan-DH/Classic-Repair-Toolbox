using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// TraceStorage - the traces the user draws on schematics, one file for every board (code review,
// 2026-10-13: it was written in place, a file that failed to parse was read as "no traces" and the
// next save wrote that over every board's traces, and a renamed schematic lost its traces).
//
// TraceStorage is a static, so this joins the "UserSettings" collection with the other tests of
// this machine's own files (LocalSchematicRenameTests uses it too); each test loads its own temp
// file and Dispose points it at nothing again.
[Collection("UserSettings")]
public sealed class TraceStorageTests : IDisposable
{
    private const string Board = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        TraceStorage.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private string NewTracesPath() => this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".traces.json");

    private static List<TraceModel> Traces(params string[] colors) =>
        colors.Select(color => new TraceModel
        {
            Color = color,
            Nodes = [new PointModel { X = 0.25, Y = 0.5 }, new PointModel { X = 0.75, Y = 0.5 }],
        }).ToList();

    private static string[] ColorsOf(List<TraceModel> traces) => traces.Select(trace => trace.Color).ToArray();

    private static string[] CorruptCopiesOf(string path) =>
        Directory.GetFiles(Path.GetDirectoryName(path)!, Path.GetFileName(path) + ".corrupt-*");

    [Fact]
    public void Saved_traces_read_back_after_a_reload()
    {
        string path = this.NewTracesPath();
        TraceStorage.LoadFrom(path);

        TraceStorage.SaveTraces(Board, "Schematic 1", Traces("#FF0000", "#00FF00"));
        TraceStorage.LoadFrom(path);

        List<TraceModel> loaded = TraceStorage.GetTraces(Board, "Schematic 1");
        Assert.Equal(["#FF0000", "#00FF00"], ColorsOf(loaded));
        Assert.Equal(0.75, loaded[0].Nodes[1].X);
        Assert.Empty(TraceStorage.GetTraces(Board, "Schematic 2"));
        Assert.Empty(TraceStorage.GetTraces("Commodore 128|310378", "Schematic 1"));
    }

    // ###########################################################################################
    // *** TWO SPELLINGS OF ONE SCHEMATIC KEEP BOTH LISTS (code review, 2026-10-13). *** An older
    // store kept "Top" and "top" apart; reading kept the first and dropped the other, and the next
    // save wrote the file without it. Both lists are now joined, and survive the save.
    // ###########################################################################################
    [Fact]
    public void Traces_under_names_differing_only_in_capitals_are_joined_not_dropped()
    {
        string path = this.NewTracesPath();
        var stored = new Dictionary<string, Dictionary<string, List<TraceModel>>>
        {
            [Board] = new()
            {
                ["Top"] = Traces("#FF0000"),
                ["top"] = Traces("#00FF00", "#0000FF"),
            },
        };
        File.WriteAllText(path, JsonSerializer.Serialize(stored));

        TraceStorage.LoadFrom(path);

        Assert.Equal(["#FF0000", "#00FF00", "#0000FF"], ColorsOf(TraceStorage.GetTraces(Board, "TOP")));

        // Saved and read again: all three still there.
        TraceStorage.SaveTraces("Commodore 128|310378", "Sheet", Traces("#123456"));
        TraceStorage.LoadFrom(path);

        Assert.Equal(["#FF0000", "#00FF00", "#0000FF"], ColorsOf(TraceStorage.GetTraces(Board, "Top")));
    }

    // Schematic names are looked up without regard to capitals - after a reload too, where the JSON
    // reader builds its own dictionaries without the comparer.
    [Fact]
    public void Schematic_names_match_without_regard_to_capitals_after_a_reload()
    {
        string path = this.NewTracesPath();
        TraceStorage.LoadFrom(path);
        TraceStorage.SaveTraces(Board, "PCB Top", Traces("#FF0000"));

        TraceStorage.LoadFrom(path);

        Assert.Single(TraceStorage.GetTraces(Board, "pcb top"));
    }

    // The save goes through AtomicJsonFile: no temp file is left, and the file parses.
    [Fact]
    public void A_save_leaves_a_whole_file_and_no_temp_file()
    {
        string path = this.NewTracesPath();
        TraceStorage.LoadFrom(path);

        TraceStorage.SaveTraces(Board, "Schematic 1", Traces("#FF0000"));

        Assert.False(File.Exists(path + ".tmp"));
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(path));
        Assert.True(document.RootElement.TryGetProperty(Board, out _));
    }

    [Fact]
    public void A_file_that_cannot_be_parsed_is_kept_aside_and_never_written_over()
    {
        string path = this.NewTracesPath();
        File.WriteAllText(path, "{ \"Commodore 64|250407\": { \"Schematic 1\": [ { \"Color\": ");

        TraceStorage.LoadFrom(path);

        Assert.Empty(TraceStorage.GetTraces(Board, "Schematic 1"));
        string kept = Assert.Single(CorruptCopiesOf(path));

        // A trace drawn now starts a new file - the kept copy is untouched.
        TraceStorage.SaveTraces(Board, "Schematic 2", Traces("#0000FF"));

        Assert.Equal("{ \"Commodore 64|250407\": { \"Schematic 1\": [ { \"Color\": ", File.ReadAllText(kept));
        TraceStorage.LoadFrom(path);
        Assert.Single(TraceStorage.GetTraces(Board, "Schematic 2"));
    }

    [Fact]
    public void A_file_holding_null_counts_as_unparsable()
    {
        string path = this.NewTracesPath();
        File.WriteAllText(path, "null");

        TraceStorage.LoadFrom(path);

        Assert.Single(CorruptCopiesOf(path));
    }

    // Not loaded at all (a headless test, say): nothing is read, and nothing is written anywhere.
    [Fact]
    public void Nothing_is_written_before_a_file_is_loaded()
    {
        TraceStorage.LoadFrom(string.Empty);

        TraceStorage.SaveTraces(Board, "Schematic 1", Traces("#FF0000"));

        Assert.Single(TraceStorage.GetTraces(Board, "Schematic 1"));
        Assert.Empty(Directory.GetFiles(this.thisWorkspace.Root));
    }

    [Fact]
    public void Renaming_moves_traces_swaps_all_at_once_and_keeps_both_lists_on_a_shared_name()
    {
        string path = this.NewTracesPath();
        TraceStorage.LoadFrom(path);
        TraceStorage.SaveTraces(Board, "A", Traces("#A"));
        TraceStorage.SaveTraces(Board, "B", Traces("#B"));
        TraceStorage.SaveTraces(Board, "C", Traces("#C"));
        TraceStorage.SaveTraces(Board, "Kept", Traces("#K"));

        // A and B swap; C goes onto "Kept", which keeps its own trace as well.
        var moves = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["A"] = "B", ["B"] = "A", ["C"] = "Kept" };

        Assert.True(TraceStorage.RenameSchematics(Board, name => moves.GetValueOrDefault(name), out var before));

        Assert.NotNull(before);
        Assert.Equal(["#B"], ColorsOf(TraceStorage.GetTraces(Board, "A")));
        Assert.Equal(["#A"], ColorsOf(TraceStorage.GetTraces(Board, "B")));
        Assert.Empty(TraceStorage.GetTraces(Board, "C"));
        Assert.Equal(["#K", "#C"], ColorsOf(TraceStorage.GetTraces(Board, "Kept")));

        // On disk too.
        TraceStorage.LoadFrom(path);
        Assert.Equal(["#B"], ColorsOf(TraceStorage.GetTraces(Board, "A")));

        // And put back as it was.
        TraceStorage.RestoreBoard(Board, before!);
        Assert.Equal(["#A"], ColorsOf(TraceStorage.GetTraces(Board, "A")));
        Assert.Equal(["#C"], ColorsOf(TraceStorage.GetTraces(Board, "C")));
    }

    [Fact]
    public void Renaming_nothing_reports_nothing_to_put_back()
    {
        TraceStorage.LoadFrom(this.NewTracesPath());
        TraceStorage.SaveTraces(Board, "A", Traces("#A"));

        Assert.True(TraceStorage.RenameSchematics(Board, _ => null, out var before));
        Assert.Null(before);
        Assert.True(TraceStorage.RenameSchematics("Other|Board", _ => "X", out before));
        Assert.Null(before);
    }
}
