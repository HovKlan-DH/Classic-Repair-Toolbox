using System;
using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Reading and writing system.json (NewContributeStrategy.md Phase 4, task 7).
//
// THE THEME OF THIS FILE IS THAT A DESCRIPTOR IS UNTRUSTED. It sits in the synced Data tree on
// the user's own disk and arrives over the network, so every read has to survive a file that is
// absent, truncated, hand-edited or written by a build that does not exist yet - and none of
// those may stop a board loading.
// ###########################################################################################
public sealed class SystemDescriptorStoreTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string SystemFolder => Path.Combine(this.thisWorkspace.Root, "Commodore", "C64", "250407");

    private string DescriptorPath => Path.Combine(this.SystemFolder, SystemDescriptorStore.FileName);

    private static SystemDescriptor Descriptor(string board = "250407") => SystemDescriptorRules.Build(
        "Commodore", "C64", board, "2026-09-21",
        new DateTimeOffset(2026, 9, 21, 10, 0, 0, TimeSpan.Zero),
        new[] { "Dennis" },
        SystemDescriptorRules.SystemOrigin.Contributed,
        new[] { new SystemContentEntry("a.png", "aaa") });

    [Fact]
    public void A_written_descriptor_reads_back_whole()
    {
        SystemDescriptor written = Descriptor();

        SystemDescriptorStore.Write(this.SystemFolder, written);

        SystemDescriptor? read = SystemDescriptorStore.Read(this.SystemFolder);

        Assert.NotNull(read);
        Assert.Equal("Commodore/C64/250407", read!.SystemId);
        Assert.Equal(written.SystemId, read.SystemId);
        Assert.Equal("Commodore", read.Manufacturer);
        Assert.Equal("C64", read.Hardware);
        Assert.Equal("250407", read.Board);
        Assert.Equal("2026-09-21", read.Revision);
        Assert.Equal(written.PublishedUtc, read.PublishedUtc);
        Assert.Equal(new[] { "Dennis" }, read.Maintainers);
        Assert.Equal("contributed", read.Origin);
        Assert.Equal(written.ContentHash, read.ContentHash);
    }

    [Fact]
    public void The_file_is_written_inside_the_system_folder_under_its_documented_name()
    {
        SystemDescriptorStore.Write(this.SystemFolder, Descriptor());

        Assert.True(File.Exists(this.DescriptorPath));
        Assert.Equal("system.json", SystemDescriptorStore.FileName);
    }

    // ###########################################################################################
    // EVERY FAILURE MODE READS AS "no descriptor", never as an exception.
    //
    // Every board that shipped before this file existed has no system.json, and that is the
    // ordinary case rather than an error. A throw here would stop those boards loading at all.
    // ###########################################################################################
    [Fact]
    public void A_system_with_no_descriptor_reads_as_nothing()
    {
        Directory.CreateDirectory(this.SystemFolder);

        Assert.Null(SystemDescriptorStore.Read(this.SystemFolder));
    }

    [Fact]
    public void A_folder_that_does_not_exist_reads_as_nothing()
    {
        Assert.Null(SystemDescriptorStore.Read(Path.Combine(this.thisWorkspace.Root, "nope")));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void A_blank_folder_reads_as_nothing(string? folder)
    {
        Assert.Null(SystemDescriptorStore.Read(folder!));
    }

    [Fact]
    public void An_unparseable_descriptor_reads_as_nothing_rather_than_throwing()
    {
        Directory.CreateDirectory(this.SystemFolder);
        File.WriteAllText(this.DescriptorPath, "{ not json at all");

        Assert.Null(SystemDescriptorStore.Read(this.SystemFolder));
    }

    // ###########################################################################################
    // A MALFORMED SystemId IS REJECTED ON READ, because that value reaches a lookup key.
    //
    // This is the one field where "take it as written" is not good enough: a hand-edited or
    // truncated id would be carried into whatever keys off it. Everything else in the file is
    // display information and is accepted as-is.
    // ###########################################################################################
    [Fact]
    public void A_descriptor_whose_system_id_is_malformed_is_ignored()
    {
        Directory.CreateDirectory(this.SystemFolder);

        File.WriteAllText(this.DescriptorPath, """
            { "SystemId": "not-a-real-id", "Manufacturer": "Commodore" }
            """);

        Assert.Null(SystemDescriptorStore.Read(this.SystemFolder));
    }

    [Fact]
    public void Writing_a_descriptor_with_a_malformed_system_id_is_refused()
    {
        var bad = new SystemDescriptor { SystemId = "nope", Manufacturer = "Commodore" };

        Assert.Throws<ArgumentException>(() => SystemDescriptorStore.Write(this.SystemFolder, bad));
    }

    // ###########################################################################################
    // FORWARDS AND BACKWARDS COMPATIBILITY, both of which will really happen.
    //
    // A descriptor written by a NEWER build carries fields this one has never heard of, and a
    // system.json this build only half understands is still worth more than none - so unknown
    // members are skipped rather than failing the parse. A descriptor written by an OLDER build is
    // missing fields, which must read as their defaults rather than as a corrupt file.
    // ###########################################################################################
    [Fact]
    public void A_descriptor_from_a_newer_build_still_loads_and_keeps_what_this_build_understands()
    {
        Directory.CreateDirectory(this.SystemFolder);

        File.WriteAllText(this.DescriptorPath, """
            {
              "SystemId": "Commodore/C64/250407",
              "Manufacturer": "Commodore",
              "Hardware": "C64",
              "SomethingFromTheFuture": { "nested": [1, 2, 3] },
              "Origin": "contributed"
            }
            """);

        SystemDescriptor? read = SystemDescriptorStore.Read(this.SystemFolder);

        Assert.NotNull(read);
        Assert.Equal("Commodore", read!.Manufacturer);
        Assert.Equal("contributed", read.Origin);
    }

    [Fact]
    public void A_descriptor_missing_fields_reads_them_as_empty_rather_than_failing()
    {
        Directory.CreateDirectory(this.SystemFolder);

        File.WriteAllText(this.DescriptorPath, """
            { "SystemId": "Commodore/C64/250407" }
            """);

        SystemDescriptor? read = SystemDescriptorStore.Read(this.SystemFolder);

        Assert.NotNull(read);
        Assert.Equal(string.Empty, read!.Manufacturer);
        Assert.Equal(string.Empty, read.Origin);
        Assert.NotNull(read.Maintainers);
        Assert.Empty(read.Maintainers);
    }

    [Fact]
    public void Writing_twice_replaces_rather_than_appends()
    {
        SystemDescriptorStore.Write(this.SystemFolder, Descriptor());
        SystemDescriptorStore.Write(this.SystemFolder, Descriptor());

        SystemDescriptor? read = SystemDescriptorStore.Read(this.SystemFolder);

        Assert.NotNull(read);
        Assert.Equal("Commodore/C64/250407", read!.SystemId);
    }

    [Fact]
    public void Writing_creates_the_system_folder_when_it_is_not_there_yet()
    {
        Assert.False(Directory.Exists(this.SystemFolder));

        SystemDescriptorStore.Write(this.SystemFolder, Descriptor());

        Assert.True(File.Exists(this.DescriptorPath));
    }
}
