using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// CRT 4 keeps using the "Data" folder (owner request, 2026-11-22; DataFolders): CRT 3's Excel tree
// in it is set aside as "Data CRT 3" at the first 4.0 launch, a pre-release's "Data v4" takes the
// name "Data", and "Data CRT 3" is deleted once nothing in it is needed any more. Each test works
// in an app folder of its own - never the user's.
// ###########################################################################################
public sealed class DataFoldersTests : IDisposable
{
    private readonly TempWorkspace thisSpace = new();

    public void Dispose() => this.thisSpace.Dispose();

    private string App => this.thisSpace.Root;

    private string Data => Path.Combine(this.App, AppConfig.DataFolderName);

    private string PreRelease => Path.Combine(this.App, AppConfig.PreReleaseDataFolderName);

    private string Aside => Path.Combine(this.App, AppConfig.Crt3DataFolderName);

    // A data folder as CRT 3 left it: its main workbooks at the top, a board under it.
    private void WriteCrt3Tree(string folder)
    {
        File.WriteAllText(this.Under(folder, "Classic-Repair-Toolbox.v3.0.0.xlsx"), "CRT 3's main workbook");
        File.WriteAllText(this.Under(folder, "Classic-Repair-Toolbox.xlsx"), "CRT 1's main workbook");
        File.WriteAllText(this.Under(folder, "Commodore", "C64", "250407", "Data C64 250407 v3.0.0.xlsx"), "a board");
    }

    // A data folder as CRT 4 keeps it: its main file at the top, a board file under it.
    private void WriteCrt4Tree(string folder)
    {
        File.WriteAllText(this.Under(folder, "Classic-Repair-Toolbox.v4.0.0.json"), "CRT 4's main file");
        File.WriteAllText(this.Under(folder, "Commodore", "C64", "250407", "board.json"), "a board file");
    }

    private string Under(string folder, params string[] parts)
    {
        string path = Path.Combine([folder, .. parts]);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        return path;
    }

    // ------------------------------------------------------------------ which folder is which

    [Fact]
    public void A_folder_with_CRT_3s_main_workbook_is_CRT_3s_and_one_with_CRT_4s_main_file_is_not()
    {
        this.WriteCrt3Tree(this.Data);
        this.WriteCrt4Tree(this.PreRelease);

        Assert.True(DataFolders.IsCrt3Data(this.Data));
        Assert.False(DataFolders.IsCrt3Data(this.PreRelease));
        Assert.False(DataFolders.IsCrt3Data(Path.Combine(this.App, "not there")));
    }

    // CRT 3 installed again and run on CRT 4's folder downloads its workbooks into it: the folder holds
    // CRT 4's main file still, and stays CRT 4's - nothing of it is set aside.
    [Fact]
    public void A_CRT_4_folder_CRT_3_also_wrote_into_stays_CRT_4s()
    {
        this.WriteCrt4Tree(this.Data);
        this.WriteCrt3Tree(this.Data);

        Assert.False(DataFolders.IsCrt3Data(this.Data));

        DataFolders.Layout layout = DataFolders.Arrange(this.App);

        Assert.Equal(this.Data, layout.DataRoot);
        Assert.Equal(string.Empty, layout.Crt3Root);
        Assert.True(File.Exists(Path.Combine(this.Data, "Classic-Repair-Toolbox.v4.0.0.json")));
    }

    // ------------------------------------------------------------------ arranging them

    [Fact]
    public void On_a_new_machine_CRT_uses_Data_and_there_is_no_CRT_3_folder()
    {
        DataFolders.Layout layout = DataFolders.Arrange(this.App);

        Assert.Equal(new DataFolders.Layout(this.Data, string.Empty, false, false, string.Empty), layout);
        Assert.False(Directory.Exists(this.Data));
    }

    // The ordinary update from CRT 3: its folder is renamed, nothing copied, and CRT 4's "Data" is new -
    // DataManager seeds it from the install package.
    [Fact]
    public void CRT_3s_Data_is_set_aside_as_Data_CRT_3_and_CRT_4_uses_Data()
    {
        this.WriteCrt3Tree(this.Data);

        DataFolders.Layout layout = DataFolders.Arrange(this.App);

        Assert.Equal(new DataFolders.Layout(this.Data, this.Aside, true, true, string.Empty), layout);
        Assert.False(Directory.Exists(this.Data));
        Assert.Equal("CRT 3's main workbook", File.ReadAllText(Path.Combine(this.Aside, "Classic-Repair-Toolbox.v3.0.0.xlsx")));
    }

    [Fact]
    public void A_pre_releases_Data_v4_becomes_Data_once_CRT_3s_is_set_aside()
    {
        this.WriteCrt3Tree(this.Data);
        this.WriteCrt4Tree(this.PreRelease);

        DataFolders.Layout layout = DataFolders.Arrange(this.App);

        Assert.Equal(new DataFolders.Layout(this.Data, this.Aside, true, true, string.Empty), layout);
        Assert.False(Directory.Exists(this.PreRelease));
        Assert.Equal("CRT 4's main file", File.ReadAllText(Path.Combine(this.Data, "Classic-Repair-Toolbox.v4.0.0.json")));
        Assert.True(File.Exists(Path.Combine(this.Aside, "Classic-Repair-Toolbox.v3.0.0.xlsx")));
    }

    // Every launch after the first: nothing is moved, and CRT 3's folder is read until it goes.
    [Fact]
    public void Once_arranged_a_launch_moves_nothing()
    {
        this.WriteCrt4Tree(this.Data);
        this.WriteCrt3Tree(this.Aside);

        DataFolders.Layout layout = DataFolders.Arrange(this.App, (from, to) => throw new InvalidOperationException($"moved [{from}] -> [{to}]"));

        Assert.Equal(new DataFolders.Layout(this.Data, this.Aside, true, false, string.Empty), layout);
    }

    // A file in CRT 3's folder held open - by Excel, say - stops the rename on Windows: this run is as the
    // pre-releases ran, on "Data v4" with "Data" read as CRT 3's, and it is never deleted from there.
    [Fact]
    public void When_CRT_3s_Data_cannot_be_set_aside_this_run_uses_Data_v4_and_reads_Data()
    {
        this.WriteCrt3Tree(this.Data);

        DataFolders.Layout layout = DataFolders.Arrange(this.App, (_, _) => throw new IOException("held open"));

        Assert.Equal(new DataFolders.Layout(this.PreRelease, this.Data, false, false, string.Empty), layout);
        Assert.True(File.Exists(Path.Combine(this.Data, "Classic-Repair-Toolbox.v3.0.0.xlsx")));
    }

    [Fact]
    public void When_Data_v4_cannot_become_Data_it_is_used_where_it_is()
    {
        this.WriteCrt3Tree(this.Data);
        this.WriteCrt4Tree(this.PreRelease);

        DataFolders.Layout layout = DataFolders.Arrange(this.App, (from, to) =>
        {
            if (from == this.PreRelease)
                throw new UnauthorizedAccessException("in use");

            Directory.Move(from, to);
        });

        Assert.Equal(new DataFolders.Layout(this.PreRelease, this.Aside, true, true, string.Empty), layout);
        Assert.True(Directory.Exists(this.PreRelease));
        Assert.False(Directory.Exists(this.Data));
    }

    // CRT 3 installed again, its "Data" made anew, while "Data CRT 3" is still there: neither is moved
    // over the other, and neither is deleted from this run.
    [Fact]
    public void A_CRT_3_Data_while_Data_CRT_3_is_still_there_is_left_where_it_is_and_read()
    {
        this.WriteCrt3Tree(this.Data);
        this.WriteCrt3Tree(this.Aside);

        DataFolders.Layout layout = DataFolders.Arrange(this.App);

        Assert.Equal(new DataFolders.Layout(this.PreRelease, this.Data, false, false, string.Empty), layout);
        Assert.True(Directory.Exists(this.Data));
        Assert.True(Directory.Exists(this.Aside));
    }

    [Fact]
    public void A_delete_stopped_half_way_is_finished_and_never_read_as_CRT_3s_folder()
    {
        this.WriteCrt4Tree(this.Data);
        this.WriteCrt3Tree(this.Aside + DataFolders.DeletingSuffix);

        DataFolders.Layout layout = DataFolders.Arrange(this.App);

        Assert.Equal(new DataFolders.Layout(this.Data, string.Empty, false, false, this.Aside + DataFolders.DeletingSuffix), layout);
    }

    // "--data-root=" moves nothing: CRT 3's folder is read wherever it is.
    [Fact]
    public void A_run_on_another_data_folder_reads_CRT_3s_folder_wherever_it_is()
    {
        Assert.Equal(string.Empty, DataFolders.Crt3RootToRead(this.App));

        this.WriteCrt3Tree(this.Data);
        Assert.Equal(this.Data, DataFolders.Crt3RootToRead(this.App));

        this.WriteCrt3Tree(this.Aside);
        Assert.Equal(this.Aside, DataFolders.Crt3RootToRead(this.App));
    }

    // ------------------------------------------------------------------ the capture folder

    // Chosen inside CRT 3's "Data", the folder went with it - the setting must not name a folder inside
    // CRT 4's new "Data" instead.
    [Fact]
    public void A_capture_folder_chosen_inside_CRT_3s_Data_follows_it_and_any_other_stays()
    {
        Assert.Equal(
            Path.Combine(this.Aside, "My captures"),
            DataFolders.CaptureFolderAfterSettingAside(Path.Combine(this.Data, "My captures"), this.Data, this.Aside));
        Assert.Equal(this.Aside, DataFolders.CaptureFolderAfterSettingAside(this.Data, this.Data, this.Aside));
        Assert.Null(DataFolders.CaptureFolderAfterSettingAside(Path.Combine(this.App, "Oscilloscope images"), this.Data, this.Aside));
        Assert.Null(DataFolders.CaptureFolderAfterSettingAside(string.Empty, this.Data, this.Aside));
    }

    // ------------------------------------------------------------------ when CRT 3's folder may go

    // Everything taken over: the own-models file made, the drafts folder holding nothing of CRT 3's still
    // to do, nothing but CRT 3's published data in the folder.
    private (string OwnModels, string Drafts, string Captures) TakenOver()
    {
        this.WriteCrt3Tree(this.Aside);
        string ownModels = this.thisSpace.WriteFile("Classic-Repair-Toolbox.own-oscilloscopes.json", "{}");
        string drafts = this.thisSpace.Path_("Drafts", "x");
        Directory.CreateDirectory(Path.GetDirectoryName(drafts)!);

        return (ownModels, Path.GetDirectoryName(drafts)!, Path.Combine(this.App, "Oscilloscope images"));
    }

    [Fact]
    public void With_everything_taken_over_CRT_3s_folder_may_go()
    {
        (string ownModels, string drafts, string captures) = this.TakenOver();

        Assert.Empty(DataFolders.ReasonsToKeep(this.Aside, drafts, ownModels, [], captures));
    }

    // The file is made once the models are carried over - not yet when CRT 3's main workbook could not
    // be opened (OwnOscilloscopeFile), and the folder is where they are read from.
    [Fact]
    public void Own_oscilloscope_models_not_carried_over_yet_keep_it()
    {
        (string ownModels, string drafts, string captures) = this.TakenOver();
        File.Delete(ownModels);

        Assert.Contains(
            DataFolders.ReasonsToKeep(this.Aside, drafts, ownModels, [], captures),
            reason => reason.Contains("oscilloscope models", StringComparison.Ordinal));
    }

    [Fact]
    public void A_capture_folder_inside_it_keeps_it()
    {
        (string ownModels, string drafts, _) = this.TakenOver();

        Assert.Contains(
            DataFolders.ReasonsToKeep(this.Aside, drafts, ownModels, [], Path.Combine(this.Aside, "My captures")),
            reason => reason.Contains("captures are saved", StringComparison.Ordinal));
    }

    // CRT 3's data had no picture at its top: a capture 3.0.0 saved there that could not be moved out, or
    // a photo of the user's, keeps the folder.
    [Theory]
    [InlineData("U8_12.png")]
    [InlineData("holiday.jpg")]
    public void A_picture_at_its_top_keeps_it(string name)
    {
        (string ownModels, string drafts, string captures) = this.TakenOver();
        File.WriteAllText(Path.Combine(this.Aside, name), "a picture");

        Assert.Contains(
            DataFolders.ReasonsToKeep(this.Aside, drafts, ownModels, [], captures),
            reason => reason.Contains($"[{name}]", StringComparison.Ordinal));
    }

    [Fact]
    public void Anything_of_CRT_3s_drafts_still_to_take_over_keeps_it()
    {
        (string ownModels, string drafts, string captures) = this.TakenOver();
        File.WriteAllText(this.Under(drafts, "Commodore", "C64", "250407", "Data C64 250407.xlsx"), "an Excel-era board put there by hand");

        Assert.Contains(
            DataFolders.ReasonsToKeep(this.Aside, drafts, ownModels, [], captures),
            reason => reason.Contains("not taken in as a draft yet", StringComparison.Ordinal));
    }

    // ------------------------------------------------------------------ deleting it

    [Fact]
    public void Going_it_is_renamed_first_and_then_deleted_whole()
    {
        this.WriteCrt3Tree(this.Aside);

        string? deleting = DataFolders.SetAsideForDeleting(this.Aside);

        Assert.Equal(this.Aside + DataFolders.DeletingSuffix, deleting);
        Assert.False(Directory.Exists(this.Aside));

        DataFolders.Delete(deleting!);

        Assert.False(Directory.Exists(deleting));
    }

    [Fact]
    public void A_folder_that_cannot_be_renamed_stays_as_it_is_for_the_next_launch()
    {
        this.WriteCrt3Tree(this.Aside);

        Assert.Null(DataFolders.SetAsideForDeleting(this.Aside, (_, _) => throw new IOException("held open")));
        Assert.True(File.Exists(Path.Combine(this.Aside, "Classic-Repair-Toolbox.v3.0.0.xlsx")));
    }
}
