using System.Collections;
using System.Reflection;
using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardKeyMove - what this machine keeps under a board's name, carried over when a new board is renamed
// (owner request, 2026-11-01). Worklogs, traces, devices and settings are filed by "{hardware}|{board}",
// so a board renamed would otherwise come back with none of them.
//
// In the "UserSettings" collection with the other tests of this machine's own files; the worklog's
// root and the traces' file are each its own temp file, pointed at nothing again afterwards.
// ###########################################################################################
[Collection("UserSettings")]
public sealed class BoardKeyMoveTests : IDisposable
{
    private const string From = "C64|999001";
    private const string To = "Commodore 64|999001A";

    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        TraceStorage.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Detached-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    // ###########################################################################################
    // EVERY "...ByBoard" SETTING MOVES - found by its name, so a setting added later cannot be
    // forgotten - with the splitter ratios; a board's catalogue keys and each of its schematics' too.
    // Another board's are left alone.
    // ###########################################################################################
    [Fact]
    public void Every_setting_kept_per_board_follows_it()
    {
        var data = new UserSettingsData();
        List<PropertyInfo> perBoard = typeof(UserSettingsData).GetProperties()
            .Where(property => property.Name.EndsWith("ByBoard", StringComparison.Ordinal) || property.Name == nameof(UserSettingsData.SchematicsSplitterRatios))
            .ToList();

        Assert.True(perBoard.Count >= 10, $"Only {perBoard.Count} found - were they renamed?");

        foreach (PropertyInfo property in perBoard)
        {
            var settings = (IDictionary)property.GetValue(data)!;
            Type valueType = property.PropertyType.GetGenericArguments()[1];
            object value = valueType == typeof(string) ? "x" : Activator.CreateInstance(valueType)!;

            settings[From] = value;
            settings["Other|board"] = value;
        }

        data.CatalogueUncheckedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { From, From + "|Top", "Other|board" };
        data.CatalogueCollapsedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { From };
        data.LastBoardByHardware["C64"] = "999001";

        string path = this.thisWorkspace.Path_("settings.json");
        File.WriteAllText(path, JsonSerializer.Serialize(data));
        UserSettings.LoadFrom(path);

        UserSettings.MoveBoard(From, To);

        UserSettingsData saved = JsonSerializer.Deserialize<UserSettingsData>(File.ReadAllText(path))!;

        foreach (PropertyInfo property in perBoard)
        {
            var settings = (IDictionary)property.GetValue(saved)!;

            Assert.False(settings.Contains(From), $"{property.Name} still has the old name");
            Assert.True(settings.Contains(To), $"{property.Name} did not follow");
            Assert.True(settings.Contains("Other|board"), $"{property.Name} lost another board's");
        }

        Assert.Equal(new HashSet<string> { To, To + "|Top", "Other|board" }, saved.CatalogueUncheckedKeys.ToHashSet());
        Assert.Equal([To], saved.CatalogueCollapsedKeys);
        Assert.Equal("999001A", saved.LastBoardByHardware["Commodore 64"]);
        Assert.False(saved.LastBoardByHardware.ContainsKey("C64"));
    }

    [Fact]
    public void Drawn_traces_follow_it()
    {
        string path = this.thisWorkspace.Path_("traces.json");
        TraceStorage.LoadFrom(path);
        TraceStorage.SaveTraces(From, "Top", [new TraceModel()]);

        Assert.True(TraceStorage.MoveBoard(From, To));

        Assert.Empty(TraceStorage.GetTraces(From, "Top"));
        Assert.Single(TraceStorage.GetTraces(To, "Top"));

        // And in the file.
        TraceStorage.LoadFrom(path);
        Assert.Single(TraceStorage.GetTraces(To, "Top"));

        // A board with none moves nothing, and that is no failure.
        Assert.True(TraceStorage.MoveBoard("No|such", To));
        Assert.Single(TraceStorage.GetTraces(To, "Top"));
    }

    [Fact]
    public void Workbooks_and_devices_on_it_follow_it()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks"));

        WorkbookRecord ours = WorklogManager.CreateWorkbook(From, "A repair", string.Empty)!;
        WorkbookRecord other = WorklogManager.CreateWorkbook("Other|board", "Another", string.Empty)!;
        OwnedDeviceRecord device = OwnedDeviceStore.Create(new OwnedDeviceDetails("Breadbin", From, string.Empty, string.Empty, string.Empty, string.Empty))!;

        Assert.True(WorklogManager.MoveBoard(From, To));
        Assert.True(OwnedDeviceStore.MoveBoard(From, To));

        Assert.Empty(WorklogManager.GetWorkbooksForBoard(From));
        Assert.Equal(ours.Id, Assert.Single(WorklogManager.GetWorkbooksForBoard(To)).Id);
        Assert.Equal(other.Id, Assert.Single(WorklogManager.GetWorkbooksForBoard("Other|board")).Id);

        Assert.Equal(To, OwnedDeviceStore.Get(device.Id)!.BoardKey);
    }
}
