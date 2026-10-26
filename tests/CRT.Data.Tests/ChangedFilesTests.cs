using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ChangedFiles - "Also on other boards" for files (CRT 4.0 step 2): the files a draft replaced - a
// row of the same name citing other bytes - and the other boards that list the bytes it replaced,
// under their own names, where those boards show them as a component's file or picture or a board
// file. And what each other board's draft makes of the new file going in.
// ###########################################################################################
public sealed class ChangedFilesTests
{
    private const string OldPdf = "1111111111111111111111111111111111111111111111111111111111111111";
    private const string NewPdf = "2222222222222222222222222222222222222222222222222222222222222222";
    private const string OldPng = "3333333333333333333333333333333333333333333333333333333333333333";
    private const string NewPng = "4444444444444444444444444444444444444444444444444444444444444444";
    private const string Other = "5555555555555555555555555555555555555555555555555555555555555555";

    private static Dictionary<string, BoardDocument.StoredFileRow> Files(params (string Name, string Sha)[] files) =>
        files.ToDictionary(file => file.Name, file => new BoardDocument.StoredFileRow { Name = file.Name, Sha256 = file.Sha }, StringComparer.OrdinalIgnoreCase);

    // U8 with a datasheet, a picture and a scope baseline - as published.
    private static BoardData Published() => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC" }],
        ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Component local files/6510.pdf" }],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/U8.png" },
            new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/U8_1.png", TimeDiv = "50us" },
        ],
    };

    private static readonly Dictionary<string, BoardDocument.StoredFileRow> PublishedFiles = Files(
        ("Component local files/6510.pdf", OldPdf),
        ("Component images/U8.png", OldPng),
        ("Scope baseline/U8_1.png", Other));

    // The draft: the datasheet a new file under a new name, the picture replaced under its own name.
    private static BoardData Draft() => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC" }],
        ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = " datasheet ", File = "Component local files/6510-2.pdf" }],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/U8.png" },
            new ComponentImageEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", File = "Scope baseline/U8_1.png", TimeDiv = "50us" },
        ],
    };

    private static readonly Dictionary<string, BoardDocument.StoredFileRow> DraftFiles = Files(
        ("Component local files/6510-2.pdf", NewPdf),
        ("Component images/U8.png", NewPng),
        ("Scope baseline/U8_1.png", Other));

    [Fact]
    public void A_file_and_a_picture_replaced_are_found_by_their_row_name()
    {
        List<ChangedFile> changed = ChangedFiles.OfComponent(Draft(), DraftFiles, Published(), PublishedFiles, "U8", "PAL");

        Assert.Equal(
            [
                new ChangedFile(ChangedFileKind.ComponentFile, "Datasheet", "Component local files/6510.pdf", OldPdf, "Component local files/6510-2.pdf", NewPdf),
                new ChangedFile(ChangedFileKind.ComponentPicture, "Top", "Component images/U8.png", OldPng, "Component images/U8.png", NewPng),
            ],
            changed);

        Assert.Equal("Datasheet (6510.pdf)", changed[0].DisplayText);
    }

    // A scope baseline is measured on its board: one replaced is never offered to another.
    [Fact]
    public void A_scope_baseline_replaced_is_no_changed_file()
    {
        Dictionary<string, BoardDocument.StoredFileRow> draftFiles = Files(
            ("Component local files/6510.pdf", OldPdf),
            ("Component images/U8.png", OldPng),
            ("Scope baseline/U8_1.png", NewPng));

        Assert.Empty(ChangedFiles.OfComponent(Published(), draftFiles, Published(), PublishedFiles, "U8", "PAL"));
    }

    // Nothing pairs a row renamed AND given a new file - nor one without a name, nor one with the
    // same bytes under another name.
    [Fact]
    public void Rows_that_cannot_be_paired_or_hold_the_same_bytes_are_no_change()
    {
        BoardData renamed = Draft();
        renamed.ComponentLocalFiles[0] = new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Pinout", File = "Component local files/6510-2.pdf" };
        renamed.ComponentImages[0] = new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Component images/U8-copy.png" };

        Dictionary<string, BoardDocument.StoredFileRow> files = Files(
            ("Component local files/6510-2.pdf", NewPdf),
            ("Component images/U8-copy.png", OldPng));

        Assert.Empty(ChangedFiles.OfComponent(renamed, files, Published(), PublishedFiles, "U8", "PAL"));

        BoardData unnamed = Published();
        unnamed.ComponentLocalFiles[0] = new ComponentLocalFileEntry { BoardLabel = "U8", Name = "", File = "Component local files/6510.pdf" };
        Assert.DoesNotContain(
            ChangedFiles.OfComponent(Draft(), DraftFiles, unnamed, PublishedFiles, "U8", "PAL"),
            change => change.Kind == ChangedFileKind.ComponentFile);
    }

    // A file the draft lists nothing for yet, or no published board at all, is nothing to say.
    [Fact]
    public void Without_both_file_lists_nothing_is_changed()
    {
        Assert.Empty(ChangedFiles.OfComponent(Draft(), DraftFiles, null, PublishedFiles, "U8", "PAL"));
        Assert.Empty(ChangedFiles.OfComponent(Draft(), null, Published(), PublishedFiles, "U8", "PAL"));
        Assert.Empty(ChangedFiles.OfComponent(Draft(), Files(), Published(), PublishedFiles, "U8", "PAL"));
    }

    // One old file two rows took two different ways: which one to offer is not known - neither is.
    [Fact]
    public void Old_bytes_that_went_two_ways_are_left_out()
    {
        BoardData published = Published();
        published.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Manual", File = "Component local files/6510.pdf" });

        BoardData draft = Draft();
        draft.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Manual", File = "Component local files/6510-3.pdf" });

        Dictionary<string, BoardDocument.StoredFileRow> files = Files(
            ("Component local files/6510-2.pdf", NewPdf),
            ("Component local files/6510-3.pdf", Other),
            ("Component images/U8.png", NewPng));

        Assert.Equal(["Top"], ChangedFiles.OfComponent(draft, files, published, PublishedFiles, "U8", "PAL").Select(change => change.RowName));
    }

    [Fact]
    public void A_board_file_replaced_is_found_by_its_name()
    {
        var published = new BoardData { BoardLocalFiles = [new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual.pdf" }] };
        var draft = new BoardData { BoardLocalFiles = [new BoardLocalFileEntry { Category = "Resources", Name = "Service manual", File = "Board local files/manual-2.pdf" }] };

        ChangedFile changed = Assert.Single(ChangedFiles.OfBoardFiles(
            draft,
            Files(("Board local files/manual-2.pdf", NewPdf)),
            published,
            Files(("Board local files/manual.pdf", OldPdf))));

        Assert.Equal(ChangedFileKind.BoardFile, changed.Kind);
        Assert.Equal((OldPdf, NewPdf), (changed.PublishedSha256, changed.DraftSha256));
    }

    // ###########################################################################################
    // The other boards: those listing the replaced bytes, under their own name - but only where a
    // component file, a component picture or a board file cites it. Bytes another board uses as a
    // scope baseline, or lists without any row citing them, are not offered.
    // ###########################################################################################
    [Fact]
    public void The_other_boards_holding_the_old_bytes_are_found_under_their_own_names()
    {
        ChangedFile datasheet = ChangedFiles.OfComponent(Draft(), DraftFiles, Published(), PublishedFiles, "U8", "PAL")[0];

        var asFile = new BoardData { ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = "Component local files/6510.pdf" }] };
        var asBoardFile = new BoardData { BoardLocalFiles = [new BoardLocalFileEntry { Name = "CPU", File = "Board local files/MOS 6510.pdf" }] };
        var asBaseline = new BoardData { ComponentImages = [new ComponentImageEntry { BoardLabel = "U7", Pin = "1", File = "Scope baseline/odd.pdf", TimeDiv = "1us" }] };
        var unCited = new BoardData();
        var otherBytes = new BoardData { ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = "Component local files/6510.pdf" }] };

        List<ChangedFileElsewhere> found = ChangedFiles.Elsewhere(
            [datasheet],
            [
                ("C64 250425", "Commodore/C64/250425/board.json", asFile, Files(("Component local files/6510.pdf", OldPdf))),
                ("C128 310378", "Commodore/C128/310378/board.json", asBoardFile, Files(("Board local files/MOS 6510.pdf", OldPdf))),
                ("C64 250466", "Commodore/C64/250466/board.json", asBaseline, Files(("Scope baseline/odd.pdf", OldPdf))),
                ("C64 250469", "Commodore/C64/250469/board.json", unCited, Files(("Component local files/6510.pdf", OldPdf))),
                ("C64 326298", "Commodore/C64/326298/board.json", otherBytes, Files(("Component local files/6510.pdf", Other))),
                ("VIC-20", "Commodore/VIC-20/250403/board.json", asFile, null),
            ]);

        Assert.Equal(
            [("C64 250425", "Component local files/6510.pdf"), ("C128 310378", "Board local files/MOS 6510.pdf")],
            found.Select(where => (where.BoardName, where.NameThere)));

        Assert.Equal("C64 250425", found[0].DisplayText);
        Assert.Equal("C128 310378 - as MOS 6510.pdf", found[1].DisplayText);
    }

    // ###########################################################################################
    // The other board's draft decides: the new file goes in only where it still has the published
    // bytes - never over a version of its own, and not where its rows no longer cite the name.
    // ###########################################################################################
    [Fact]
    public void The_other_draft_takes_the_new_file_only_over_the_published_bytes()
    {
        ChangedFile datasheet = ChangedFiles.OfComponent(Draft(), DraftFiles, Published(), PublishedFiles, "U8", "PAL")[0];
        var target = new ChangedFileElsewhere(datasheet, "C64 250425", "Commodore/C64/250425/board.json", "Component local files/6510.pdf");
        var board = new BoardData { ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = "Component local files/./6510.pdf" }] };

        Assert.Equal(ChangedFilePut.Put, ChangedFiles.CanPut(board, Files(("Component local files/6510.pdf", OldPdf)), target));
        Assert.Equal(ChangedFilePut.AlreadySo, ChangedFiles.CanPut(board, Files(("Component local files/6510.pdf", NewPdf)), target));
        Assert.Equal(ChangedFilePut.OwnVersion, ChangedFiles.CanPut(board, Files(("Component local files/6510.pdf", Other)), target));
        Assert.Equal(ChangedFilePut.NoLongerThere, ChangedFiles.CanPut(board, Files(), target));
        Assert.Equal(ChangedFilePut.NoLongerThere, ChangedFiles.CanPut(new BoardData(), Files(("Component local files/6510.pdf", OldPdf)), target));
    }

    // ###########################################################################################
    // The new file goes in under a name of that board's own (code review, 2026-10-26): in its folder
    // for the old file, under the new file's own name - its own TYPE with it, so text never sits under
    // a PDF's name - and never a name that board has.
    // ###########################################################################################
    [Fact]
    public void The_new_file_gets_a_name_of_that_boards_own_of_its_own_type()
    {
        ChangedFile asText = new(ChangedFileKind.ComponentFile, "Datasheet", "Component local files/6510.pdf", OldPdf, "Datasheets/6510 rev B.txt", NewPdf);
        var target = new ChangedFileElsewhere(asText, "C64 250425", "Commodore/C64/250425/board.json", "Component local files/MOS 6510.pdf");

        Assert.Equal("Component local files/6510 rev B.txt", ChangedFiles.NameFor(target, ["Component local files/MOS 6510.pdf"]));
        Assert.Equal(
            "Component local files/6510 rev B (2).txt",
            ChangedFiles.NameFor(target, ["Component local files/MOS 6510.pdf", "component local files/6510 REV B.txt"]));

        // A name at the board's top keeps no folder.
        Assert.Equal("6510 rev B.txt", ChangedFiles.NameFor(target with { NameThere = "6510.pdf" }, []));
    }

    // ###########################################################################################
    // Every component file, picture and board file citing the old file is pointed at the new one -
    // every other field kept, every other row the very same object; a scope baseline and a schematic
    // keep the old file, and a board citing nothing of it is the same board.
    // ###########################################################################################
    [Fact]
    public void Every_row_showing_the_old_file_is_pointed_at_the_new_one()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Odd", SchematicImageFile = "Component local files/6510.pdf" }],
            ComponentLocalFiles =
            [
                new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = "Component local files/./6510.pdf" },
                new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "component local files/6510.PDF" },
                new ComponentLocalFileEntry { BoardLabel = "U9", Name = "Other", File = "Component local files/other.pdf" },
            ],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U7", Name = "Pinout", File = "Component local files/6510.pdf", Note = "From the sheet" },
                new ComponentImageEntry { BoardLabel = "U7", Pin = "1", File = "Component local files/6510.pdf", TimeDiv = "1us" },
            ],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "Resources", Name = "CPU", File = "Component local files/6510.pdf" }],
        };

        BoardData pointed = ChangedFiles.Repointed(board, "Component local files/6510.pdf", "Component local files/6510 rev B.pdf");

        Assert.Equal(
            ["Component local files/6510 rev B.pdf", "Component local files/6510 rev B.pdf", "Component local files/other.pdf"],
            pointed.ComponentLocalFiles.Select(row => row.File));
        Assert.Same(board.ComponentLocalFiles[2], pointed.ComponentLocalFiles[2]);
        Assert.Equal(("Pinout", "From the sheet", "Component local files/6510 rev B.pdf"), (pointed.ComponentImages[0].Name, pointed.ComponentImages[0].Note, pointed.ComponentImages[0].File));
        Assert.Same(board.ComponentImages[1], pointed.ComponentImages[1]);
        Assert.Equal(("Resources", "CPU", "Component local files/6510 rev B.pdf"), (pointed.BoardLocalFiles[0].Category, pointed.BoardLocalFiles[0].Name, pointed.BoardLocalFiles[0].File));
        Assert.Same(board.Schematics, pointed.Schematics);

        Assert.Same(board, ChangedFiles.Repointed(board, "Component local files/none.pdf", "x.pdf"));
    }

    // A picture's row is copied field by field - a field added to it later would be lost here unless
    // copied too, so every one is checked by name.
    [Fact]
    public void A_picture_pointed_at_the_new_file_keeps_every_other_field()
    {
        var row = new ComponentImageEntry();

        foreach (System.Reflection.PropertyInfo property in typeof(ComponentImageEntry).GetProperties().Where(property => property.PropertyType == typeof(string) && property.Name != nameof(ComponentImageEntry.Pin) && property.Name != nameof(ComponentImageEntry.TimeDiv) && property.Name != nameof(ComponentImageEntry.VoltsDiv) && property.Name != nameof(ComponentImageEntry.TriggerLevelVolts)))
        {
            property.SetValue(row, property.Name == nameof(ComponentImageEntry.File) ? "old.png" : property.Name);
        }

        ComponentImageEntry pointed = ChangedFiles.Repointed(new BoardData { ComponentImages = [row] }, "old.png", "new.png").ComponentImages.Single();

        Assert.Equal("new.png", pointed.File);
        Assert.All(
            typeof(ComponentImageEntry).GetProperties().Where(property => property.Name != nameof(ComponentImageEntry.File)),
            property => Assert.Equal(property.GetValue(row), property.GetValue(pointed)));
    }

    // Put in before, the board's rows cite the new bytes under a name of their own: it is so already.
    [Fact]
    public void A_board_whose_rows_cite_the_new_bytes_has_it_already()
    {
        ChangedFile datasheet = ChangedFiles.OfComponent(Draft(), DraftFiles, Published(), PublishedFiles, "U8", "PAL")[0];
        var target = new ChangedFileElsewhere(datasheet, "C64 250425", "Commodore/C64/250425/board.json", "Component local files/6510.pdf");
        var board = new BoardData { ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = "Component local files/6510-2.pdf" }] };

        Assert.Equal(ChangedFilePut.AlreadySo, ChangedFiles.CanPut(board, Files(("Component local files/6510-2.pdf", NewPdf)), target));
        Assert.Equal(ChangedFilePut.NoLongerThere, ChangedFiles.CanPut(board, Files(("Component local files/6510-2.pdf", Other)), target));
    }

    // ###########################################################################################
    // A name another board's file list gives that climbs out of its folder is never offered (code
    // review, 2026-10-26): the new file would be copied to wherever it points, and a file already
    // there set aside and deleted. The server refuses such names, so only a tampered tree has one.
    // ###########################################################################################
    [Fact]
    public void A_name_that_climbs_out_of_the_boards_folder_is_never_offered()
    {
        ChangedFile datasheet = ChangedFiles.OfComponent(Draft(), DraftFiles, Published(), PublishedFiles, "U8", "PAL")[0];
        const string Climbing = "Component local files/../../../../Documents/x.pdf";
        var board = new BoardData { ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = Climbing }] };

        Assert.Empty(ChangedFiles.Elsewhere([datasheet], [("C64 250425", "Commodore/C64/250425/board.json", board, Files((Climbing, OldPdf)))]));
    }
}
