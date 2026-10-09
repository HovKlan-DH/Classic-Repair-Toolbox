using System;
using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftMarkerStore - the ".crt-draft.json" marker that makes a folder under Drafts/ a draft.
//
// *** THE MARKER KEEPS THE NAMES IT HAD BEFORE "SYSTEM" BECAME "BOARD" (owner decision,
// 2026-10-09). *** Every draft already on a contributor's disk has a marker naming its board
// "SystemKey" and a draft-only board's registration "NewSystem". A marker that no longer read
// would turn the draft into an unknown folder - the contributor's edits would stop being theirs -
// so the code says BoardKey and NewBoard while the file keeps the old names.
// ###########################################################################################
public sealed class DraftMarkerStoreTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    [Fact]
    public void A_marker_written_before_the_rename_still_names_its_board_and_registration()
    {
        string path = this.thisWorkspace.WriteFile(
            Path.Combine("Drafts", "Commodore", "C64", "250407", DraftFolderLayout.DraftMarkerFileName),
            """
            {
              "SystemKey": "Commodore/C64/250407/C64 250407.xlsx",
              "BaseRevision": "2026-October-1",
              "NewSystem": {
                "HardwareName": "Commodore 64",
                "BoardName": "250407",
                "ExcelDataFile": "Commodore/C64/250407/C64 250407.xlsx"
              },
              "CreatedUtc": "2026-10-01T12:00:00.0000000Z"
            }
            """);

        DraftMarker? marker = DraftMarkerStore.Load(path);

        Assert.NotNull(marker);
        Assert.Equal("Commodore/C64/250407/C64 250407.xlsx", marker!.BoardKey);
        Assert.True(marker.IsNewBoard);
        Assert.Equal(("Commodore 64", "250407"), (marker.NewBoard!.HardwareName, marker.NewBoard.BoardName));
    }

    [Fact]
    public void A_marker_is_written_with_the_names_every_CRT_reads()
    {
        string path = this.thisWorkspace.Path_("Drafts", "Commodore", "C64", "250407", DraftFolderLayout.DraftMarkerFileName);

        DraftMarkerStore.Save(path, new DraftMarker
        {
            BoardKey = "Commodore/C64/250407/C64 250407.xlsx",
            NewBoard = new NewBoardRegistration { HardwareName = "Commodore 64", BoardName = "250407" },
        });

        string written = File.ReadAllText(path);

        Assert.Contains("\"SystemKey\"", written, StringComparison.Ordinal);
        Assert.Contains("\"NewSystem\"", written, StringComparison.Ordinal);
        Assert.DoesNotContain("BoardKey", written, StringComparison.Ordinal);
        Assert.DoesNotContain("\"NewBoard\"", written, StringComparison.Ordinal);
        Assert.Equal("Commodore/C64/250407/C64 250407.xlsx", DraftMarkerStore.Load(path)!.BoardKey);
    }
}
