using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Carrying "Show changes only" out of the separate maintainer application's settings file, once
// (2026-09-29: it became CRT's Maintainer tab). The file is gone afterwards either way, and
// nothing about it may stop CRT starting.
//
// Every test points the migration at a temp folder; none reaches the real AppData folder.
// ###########################################################################################
public sealed class MaintainerSettingsMigrationTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string LegacyPath => this.thisWorkspace.Path_(MaintainerSettingsMigration.LegacyFileName);

    // What the old application wrote: System.Text.Json's default names, placement and all.
    private void WriteLegacy(string json) => File.WriteAllText(this.LegacyPath, json);

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void The_old_files_choice_is_carried_over_and_the_file_removed(bool showChangesOnly)
    {
        this.WriteLegacy(
            "{ \"HasWindowPlacement\": true, \"WindowMaximized\": true, \"WindowX\": 10, \"WindowY\": 20, " +
            $"\"ShowChangesOnly\": {(showChangesOnly ? "true" : "false")} }}");

        bool? carried = null;

        bool handled = MaintainerSettingsMigration.Apply(this.thisWorkspace.Root, value => carried = value);

        Assert.True(handled);
        Assert.Equal(showChangesOnly, carried);
        Assert.False(File.Exists(this.LegacyPath));
    }

    // No file is the ordinary case for everybody who never ran the Maintainer tab.
    [Fact]
    public void No_old_file_changes_nothing()
    {
        bool called = false;

        Assert.False(MaintainerSettingsMigration.Apply(this.thisWorkspace.Root, _ => called = true));
        Assert.False(called);
    }

    // A half-written or hand-edited file carries nothing, throws nothing, and is removed so the
    // failure is not repeated on every launch.
    [Theory]
    [InlineData("not json at all")]
    [InlineData("")]
    [InlineData("[1, 2, 3]")]
    [InlineData("{ \"ShowChangesOnly\": \"yes\" }")]
    [InlineData("{ \"WindowX\": 10 }")]
    public void A_file_with_nothing_usable_carries_nothing_and_is_removed(string json)
    {
        this.WriteLegacy(json);

        bool called = false;

        bool handled = MaintainerSettingsMigration.Apply(this.thisWorkspace.Root, _ => called = true);

        Assert.True(handled);
        Assert.False(called);
        Assert.False(File.Exists(this.LegacyPath));
    }

    // A blank folder (AppData could not be resolved) is not an error either.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_folder_changes_nothing(string folder)
    {
        bool called = false;

        Assert.False(MaintainerSettingsMigration.Apply(folder, _ => called = true));
        Assert.False(called);
    }
}
