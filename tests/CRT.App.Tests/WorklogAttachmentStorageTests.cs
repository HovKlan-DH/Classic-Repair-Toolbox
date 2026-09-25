using Handlers.DataHandling;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// Tests for WorklogAttachmentStorage - the naming and vetting decisions made when a user picks or
// drops a photo/file onto a worklog entry, extracted from WorklogEntryEditorWindow so they can be
// tested without a window, a file dialog or a drop event.
//
// The collision rules matter more than they look: attachments are stored by name in one folder per
// entry, so a naming bug does not throw, it silently overwrites a photo the user already added.
public class WorklogAttachmentStorageTests
{
    // ------------------------------------------------------------------- ValidateSourceFile

    [Fact]
    public void A_real_image_file_is_accepted()
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_("photo.png");
        File.WriteAllText(path, "not really a png, but the extension and existence are what is checked");

        var problem = WorklogAttachmentStorage.ValidateSourceFile(path, WorklogAttachmentStorage.AttachmentKind.Photo);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.None, problem);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void A_blank_path_reports_that_nothing_was_selected(string? path)
    {
        var problem = WorklogAttachmentStorage.ValidateSourceFile(path, WorklogAttachmentStorage.AttachmentKind.Photo);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.NoFileSelected, problem);
    }

    // A file the user dragged in and then moved or deleted before pressing Add.
    [Fact]
    public void A_missing_file_is_reported_as_not_found()
    {
        using var workspace = new TempWorkspace();

        var problem = WorklogAttachmentStorage.ValidateSourceFile(workspace.Path_("gone.png"), WorklogAttachmentStorage.AttachmentKind.Photo);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.FileNotFound, problem);
    }

    // The photos section draws its attachments with Avalonia's Bitmap, so a format Bitmap cannot
    // decode has to be refused up front - it would otherwise be added and then show as a blank
    // frame forever. .svg is the interesting case: ExternalTargetLauncher allows it (the OS shell
    // opens it) but the app cannot draw it, so it must be refused here.
    [Theory]
    [InlineData("notes.txt")]
    [InlineData("datasheet.pdf")]
    [InlineData("diagram.svg")]
    [InlineData("archive.zip")]
    public void A_file_the_app_cannot_draw_is_refused_for_photos(string fileName)
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_(fileName);
        File.WriteAllText(path, "x");

        var problem = WorklogAttachmentStorage.ValidateSourceFile(path, WorklogAttachmentStorage.AttachmentKind.Photo);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.NotDisplayableImage, problem);
    }

    // The same file is fine for the Files section, which never tries to draw it - it is handed to
    // the OS shell instead, so the wider openable-document set applies.
    [Fact]
    public void A_document_is_accepted_as_a_file()
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_("datasheet.pdf");
        File.WriteAllText(path, "x");

        var problem = WorklogAttachmentStorage.ValidateSourceFile(path, WorklogAttachmentStorage.AttachmentKind.File);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.None, problem);
    }

    // Format is judged before existence: telling the user the type is wrong is more useful than
    // "not found" when they picked a .txt that is sitting right there.
    [Fact]
    public void A_wrong_type_is_reported_even_when_the_file_does_not_exist()
    {
        using var workspace = new TempWorkspace();

        var problem = WorklogAttachmentStorage.ValidateSourceFile(workspace.Path_("missing.txt"), WorklogAttachmentStorage.AttachmentKind.Photo);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.NotDisplayableImage, problem);
    }

    [Fact]
    public void Every_refusal_carries_a_message()
    {
        foreach (var problem in new[]
        {
            WorklogAttachmentStorage.AttachmentProblem.NoFileSelected,
            WorklogAttachmentStorage.AttachmentProblem.FileNotFound,
            WorklogAttachmentStorage.AttachmentProblem.NotDisplayableImage
        })
        {
            Assert.False(string.IsNullOrWhiteSpace(WorklogAttachmentStorage.DescribeProblem(problem)));
        }
    }

    // "None" is not a refusal, so it deliberately has nothing to say.
    [Fact]
    public void An_accepted_file_has_no_message()
    {
        Assert.Equal(string.Empty, WorklogAttachmentStorage.DescribeProblem(WorklogAttachmentStorage.AttachmentProblem.None));
    }

    // The Files section hands its attachments to the OS shell, which RUNS these rather than
    // displaying them - so an attachment that would become code execution on click is refused at
    // the point of attaching, not left to fail (or worse, succeed) at open time.
    [Theory]
    [InlineData("installer.exe")]
    [InlineData("script.bat")]
    [InlineData("shortcut.lnk")]
    [InlineData("macro.vbs")]
    [InlineData("shell.ps1")]
    [InlineData("no-extension")]
    public void An_executable_or_script_is_refused_as_a_file(string fileName)
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_(fileName);
        File.WriteAllText(path, "x");

        var problem = WorklogAttachmentStorage.ValidateSourceFile(path, WorklogAttachmentStorage.AttachmentKind.File);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.NotOpenableFile, problem);
    }

    // An image is legitimate as a File too - the user may prefer to attach it as a document rather
    // than as a photo, and the launcher opens it happily.
    [Theory]
    [InlineData("board.png")]
    [InlineData("manual.pdf")]
    [InlineData("readings.csv")]
    [InlineData("notes.txt")]
    [InlineData("diagram.svg")]
    public void A_document_or_image_is_accepted_as_a_file(string fileName)
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_(fileName);
        File.WriteAllText(path, "x");

        var problem = WorklogAttachmentStorage.ValidateSourceFile(path, WorklogAttachmentStorage.AttachmentKind.File);

        Assert.Equal(WorklogAttachmentStorage.AttachmentProblem.None, problem);
    }

    [Fact]
    public void The_unopenable_file_refusal_carries_a_message()
    {
        Assert.False(string.IsNullOrWhiteSpace(
            WorklogAttachmentStorage.DescribeProblem(WorklogAttachmentStorage.AttachmentProblem.NotOpenableFile)));
    }

    // ------------------------------------------------------------------ GetDisplayFileName

    [Fact]
    public void A_photo_display_name_drops_the_prefix_and_id()
    {
        Assert.Equal("board.png", WorklogAttachmentStorage.GetDisplayFileName("photo_3_board.png", WorklogAttachmentStorage.PhotoFilePrefix, 3));
    }

    [Fact]
    public void A_file_display_name_drops_the_prefix_and_id()
    {
        Assert.Equal("manual.pdf", WorklogAttachmentStorage.GetDisplayFileName("file_2_manual.pdf", WorklogAttachmentStorage.FileFilePrefix, 2));
    }

    // The prefix is stripped once, not repeatedly: a source file the user had already named
    // "2_schematic.png" becomes "photo_1_2_schematic.png" when stored as photo #1, and must read
    // back as the name they chose rather than losing another segment.
    [Fact]
    public void Only_one_id_segment_is_stripped()
    {
        Assert.Equal("2_schematic.png", WorklogAttachmentStorage.GetDisplayFileName("photo_1_2_schematic.png", WorklogAttachmentStorage.PhotoFilePrefix, 1));
    }

    // An attachment recorded before this naming scheme has no prefix to strip and must still show
    // something sensible rather than losing its first segment.
    [Theory]
    [InlineData("board.png")]
    [InlineData("my_notes.txt")]
    [InlineData("_leading.png")]
    public void A_name_without_an_id_prefix_is_returned_unchanged(string storedName)
    {
        Assert.Equal(storedName, WorklogAttachmentStorage.GetDisplayFileName(storedName, WorklogAttachmentStorage.PhotoFilePrefix, 3));
    }

    // Attachments stored by a build before an underscore was added between the prefix and the id
    // ("photo3_board.png" rather than "photo_3_board.png" - see BuildStoredFileName's own comment)
    // are still on disk and still named that way in entries.json, since existing attachments are
    // not migrated. The older form must keep reading back correctly, or every attachment stored
    // before the change shows its raw storage name forever the moment the app upgrades.
    [Fact]
    public void A_photo_stored_in_the_older_no_underscore_form_still_has_it_stripped()
    {
        Assert.Equal("board.png", WorklogAttachmentStorage.GetDisplayFileName("photo3_board.png", WorklogAttachmentStorage.PhotoFilePrefix, 3));
    }

    [Fact]
    public void A_file_stored_in_the_older_no_underscore_form_still_has_it_stripped()
    {
        Assert.Equal("manual.pdf", WorklogAttachmentStorage.GetDisplayFileName("file2_manual.pdf", WorklogAttachmentStorage.FileFilePrefix, 2));
    }

    // Photos stored by the build where PhotoFilePrefix was "" are named bare, as "3_board.png".
    // Those files are still on disk and still named in entries.json, so the bare form has to be
    // recognised too - otherwise every photo attached before the prefix was introduced shows its
    // raw storage name permanently, which is exactly the noise this method removes.
    [Fact]
    public void A_photo_stored_with_the_legacy_bare_id_prefix_still_has_it_stripped()
    {
        Assert.Equal("board.png", WorklogAttachmentStorage.GetDisplayFileName("3_board.png", WorklogAttachmentStorage.PhotoFilePrefix, 3));
    }

    // The id still has to be the owning record's. A user's own file that happens to start with
    // some OTHER number is left alone, so the legacy allowance above does not become a blanket
    // "strip any digits then underscore" rule.
    [Fact]
    public void A_legacy_bare_prefix_belonging_to_another_id_is_left_alone()
    {
        Assert.Equal("3_board.png", WorklogAttachmentStorage.GetDisplayFileName("3_board.png", WorklogAttachmentStorage.PhotoFilePrefix, 7));
    }

    // Files never had a bare-id era - they were introduced with the "file" prefix already in
    // place - but the same allowance applies to them for consistency, keyed to their own id.
    [Fact]
    public void A_name_that_is_only_a_prefix_keeps_its_raw_form()
    {
        // Stripping would leave an empty display name, which shows a row with no filename at all.
        Assert.Equal("photo_3_", WorklogAttachmentStorage.GetDisplayFileName("photo_3_", WorklogAttachmentStorage.PhotoFilePrefix, 3));
    }

    // A File's name must not have a bare id stripped as though it were a Photo's, and vice versa.
    [Fact]
    public void A_mismatched_prefix_leaves_the_name_alone()
    {
        Assert.Equal("file_2_manual.pdf", WorklogAttachmentStorage.GetDisplayFileName("file_2_manual.pdf", WorklogAttachmentStorage.PhotoFilePrefix, 2));
    }

    // The mangling this guards against: "file_2_backup.pdf" is BOTH a name a user may legitimately
    // choose and exactly what file #2 is stored as - nothing in the text tells them apart. Matching
    // the owning record's own id is what resolves it: read as file #7, this is a user-chosen name
    // and is left whole. (Read as file #2, the test below shows it strips.)
    [Fact]
    public void A_user_file_whose_own_name_starts_with_the_prefix_is_not_mangled()
    {
        Assert.Equal(
            "file_2_backup.pdf",
            WorklogAttachmentStorage.GetDisplayFileName("file_2_backup.pdf", WorklogAttachmentStorage.FileFilePrefix, 7));
    }

    // The same file once actually stored: its real name carries this record's id, so it strips back
    // to exactly the name the user chose - one segment removed, not two.
    [Fact]
    public void That_same_name_stored_properly_strips_back_to_the_users_name()
    {
        string stored = WorklogAttachmentStorage.BuildStoredFileName("file_2_backup.pdf", WorklogAttachmentStorage.FileFilePrefix, 7);

        Assert.Equal("file_7_file_2_backup.pdf", stored);
        Assert.Equal("file_2_backup.pdf", WorklogAttachmentStorage.GetDisplayFileName(stored, WorklogAttachmentStorage.FileFilePrefix, 7));
    }

    // The other half of the ambiguity: read as the record it actually names, the prefix IS this
    // record's own and is removed.
    [Fact]
    public void The_same_name_read_as_its_own_record_is_stripped()
    {
        Assert.Equal(
            "backup.pdf",
            WorklogAttachmentStorage.GetDisplayFileName("file_2_backup.pdf", WorklogAttachmentStorage.FileFilePrefix, 2));
    }

    // A stored name belonging to a DIFFERENT record is never stripped, so a mismatched id cannot
    // silently shorten another row's name.
    [Fact]
    public void Another_records_stored_name_is_left_alone()
    {
        Assert.Equal(
            "photo_3_board.png",
            WorklogAttachmentStorage.GetDisplayFileName("photo_3_board.png", WorklogAttachmentStorage.PhotoFilePrefix, 4));
    }

    // A prefix followed immediately by an underscore has no id, so it is not a name this class
    // built and nothing is stripped - and this must hold for BOTH forms: the new "photo_1_" one
    // and the older bare-id "1_" one that A_photo_stored_with_the_legacy_bare_id_prefix_still_has_it_stripped
    // covers, since "photo_board.png" happens to also start with the plain "photo_" text both
    // schemes share.
    [Fact]
    public void A_prefix_with_no_id_is_left_alone()
    {
        Assert.Equal("photo_board.png", WorklogAttachmentStorage.GetDisplayFileName("photo_board.png", WorklogAttachmentStorage.PhotoFilePrefix, 1));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void A_blank_stored_name_displays_as_blank(string? storedName)
    {
        Assert.Equal(string.Empty, WorklogAttachmentStorage.GetDisplayFileName(storedName, WorklogAttachmentStorage.PhotoFilePrefix, 1));
    }

    // Nothing left after the prefix means the prefix was the whole name; showing an empty link
    // would give the user nothing to click, so the stored name is kept.
    [Fact]
    public void A_name_that_is_only_a_prefix_is_kept_whole()
    {
        Assert.Equal("photo_3_", WorklogAttachmentStorage.GetDisplayFileName("photo_3_", WorklogAttachmentStorage.PhotoFilePrefix, 3));
    }

    // ------------------------------------------------------------- CopyAttachmentIntoFolder

    [Fact]
    public void Copying_puts_the_bytes_in_the_folder_under_the_stored_name()
    {
        using var workspace = new TempWorkspace();
        string source = workspace.Path_("source", "photo.png");
        File.WriteAllText(source, "image-bytes");
        string destinationFolder = workspace.Path_("entry-1-files", "placeholder");
        destinationFolder = Path.GetDirectoryName(destinationFolder)!;

        bool copied = WorklogAttachmentStorage.CopyAttachmentIntoFolder(source, destinationFolder, "stored.png");

        Assert.True(copied);
        Assert.Equal("image-bytes", File.ReadAllText(Path.Combine(destinationFolder, "stored.png")));
    }

    // The attachments folder does not exist until the first attachment is added.
    [Fact]
    public void Copying_creates_the_attachments_folder_when_it_is_missing()
    {
        using var workspace = new TempWorkspace();
        string source = workspace.Path_("photo.png");
        File.WriteAllText(source, "x");
        string destinationFolder = Path.Combine(workspace.Root, "entry-7-files");

        bool copied = WorklogAttachmentStorage.CopyAttachmentIntoFolder(source, destinationFolder, "stored.png");

        Assert.True(copied);
        Assert.True(File.Exists(Path.Combine(destinationFolder, "stored.png")));
    }

    // Reported rather than thrown, so the caller can refuse to add a row that would point at a
    // file which never landed. A row whose image can never load is worse than a refused add.
    [Fact]
    public void A_failed_copy_is_reported_rather_than_thrown()
    {
        using var workspace = new TempWorkspace();
        string destinationFolder = Path.Combine(workspace.Root, "entry-1-files");

        bool copied = WorklogAttachmentStorage.CopyAttachmentIntoFolder(
            Path.Combine(workspace.Root, "does-not-exist.png"),
            destinationFolder,
            "stored.png");

        Assert.False(copied);
        Assert.False(File.Exists(Path.Combine(destinationFolder, "stored.png")));
    }

    // Overwriting is intended now that stored names are built from the owning record's id:
    // replacing a photo's image reuses that photo's name on purpose, and the old bytes are meant
    // to go. There is no accidental-collision case left for a refusal to protect against.
    [Fact]
    public void Copying_over_an_existing_stored_name_replaces_it()
    {
        using var workspace = new TempWorkspace();
        string source = workspace.Path_("new.png");
        File.WriteAllText(source, "new-bytes");

        string destinationFolder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(destinationFolder);
        File.WriteAllText(Path.Combine(destinationFolder, "stored.png"), "original-bytes");

        bool copied = WorklogAttachmentStorage.CopyAttachmentIntoFolder(source, destinationFolder, "stored.png");

        Assert.True(copied);
        Assert.Equal("new-bytes", File.ReadAllText(Path.Combine(destinationFolder, "stored.png")));
    }

    // --------------------------------------------------------------- TryReplaceAttachmentFile

    // Builds an attachments folder holding one photo, and returns the folder plus the stored name.
    private static (string Folder, string StoredName) GivenAnExistingPhoto(TempWorkspace workspace, string sourceName, int photoId)
    {
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);

        string storedName = WorklogAttachmentStorage.BuildStoredFileName(sourceName, WorklogAttachmentStorage.PhotoFilePrefix, photoId);
        File.WriteAllText(Path.Combine(folder, storedName), "original-bytes");

        return (folder, storedName);
    }

    // Changing the extension changes the stored name, so the previous file would be orphaned -
    // invisible to the app but taking up space forever - if it were not removed.
    [Fact]
    public void Replacing_with_a_different_extension_deletes_the_original()
    {
        using var workspace = new TempWorkspace();
        var (folder, originalName) = GivenAnExistingPhoto(workspace, "board.jpg", photoId: 3);

        string source = workspace.Path_("replacement.png");
        File.WriteAllText(source, "new-bytes");

        bool replaced = WorklogAttachmentStorage.TryReplaceAttachmentFile(
            source,
            folder,
            originalName,
            WorklogAttachmentStorage.BuildStoredFileName(source, WorklogAttachmentStorage.PhotoFilePrefix, 3),
            out string storedName);

        Assert.True(replaced);
        Assert.False(File.Exists(Path.Combine(folder, originalName)));
        Assert.Equal("new-bytes", File.ReadAllText(Path.Combine(folder, storedName)));
        Assert.Single(Directory.GetFiles(folder));
    }

    // A differently-named source with the same extension also changes the stored name.
    [Fact]
    public void Replacing_with_a_different_file_name_deletes_the_original()
    {
        using var workspace = new TempWorkspace();
        var (folder, originalName) = GivenAnExistingPhoto(workspace, "board.png", photoId: 3);

        string source = workspace.Path_("something-else.png");
        File.WriteAllText(source, "new-bytes");

        WorklogAttachmentStorage.TryReplaceAttachmentFile(
            source,
            folder,
            originalName,
            WorklogAttachmentStorage.BuildStoredFileName(source, WorklogAttachmentStorage.PhotoFilePrefix, 3),
            out string storedName);

        Assert.NotEqual(originalName, storedName);
        Assert.False(File.Exists(Path.Combine(folder, originalName)));
        Assert.Single(Directory.GetFiles(folder));
    }

    // The dangerous case: re-picking a file with the same name gives the SAME stored name, so the
    // copy overwrote it in place. Deleting "the previous file" here would delete the replacement
    // that was just written, leaving the row pointing at nothing.
    [Fact]
    public void Replacing_with_the_same_stored_name_keeps_the_new_bytes()
    {
        using var workspace = new TempWorkspace();
        var (folder, originalName) = GivenAnExistingPhoto(workspace, "board.png", photoId: 3);

        string source = workspace.Path_("elsewhere", "board.png");
        File.WriteAllText(source, "new-bytes");

        bool replaced = WorklogAttachmentStorage.TryReplaceAttachmentFile(
            source,
            folder,
            originalName,
            WorklogAttachmentStorage.BuildStoredFileName(source, WorklogAttachmentStorage.PhotoFilePrefix, 3),
            out string storedName);

        Assert.True(replaced);
        Assert.Equal(originalName, storedName);
        Assert.True(File.Exists(Path.Combine(folder, storedName)));
        Assert.Equal("new-bytes", File.ReadAllText(Path.Combine(folder, storedName)));
        Assert.Single(Directory.GetFiles(folder));
    }

    // A failed copy must leave the original alone and keep the record pointing at it, rather than
    // deleting the only file the row still has.
    [Fact]
    public void A_failed_replacement_keeps_the_original_file_and_name()
    {
        using var workspace = new TempWorkspace();
        var (folder, originalName) = GivenAnExistingPhoto(workspace, "board.png", photoId: 3);

        bool replaced = WorklogAttachmentStorage.TryReplaceAttachmentFile(
            Path.Combine(workspace.Root, "does-not-exist.png"),
            folder,
            originalName,
            "3_does-not-exist.png",
            out string storedName);

        Assert.False(replaced);
        Assert.Equal(originalName, storedName);
        Assert.Equal("original-bytes", File.ReadAllText(Path.Combine(folder, originalName)));
        Assert.Single(Directory.GetFiles(folder));
    }

    // ------------------------------------------------------------------ BuildStoredFileName

    // The id prefix is what makes the name unique by construction, and what later lets a deleted
    // photo's file be identified with confidence.
    [Fact]
    public void A_stored_name_carries_its_owner_prefix_and_id()
    {
        string name = WorklogAttachmentStorage.BuildStoredFileName(
            Path.Combine("C:", "camera", "IMG_1234.jpg"),
            WorklogAttachmentStorage.PhotoFilePrefix,
            3);

        Assert.Equal("photo_3_IMG_1234.jpg", name);
    }

    // The prefix and id are separated by an underscore FROM EACH OTHER, not just from the source
    // name - "file_1_README.txt", not "file1_README.txt". Asserted for both prefixes since the
    // reported request named both photo and file naming explicitly.
    [Fact]
    public void A_photo_stored_name_has_its_own_underscore_between_prefix_and_id()
    {
        Assert.Equal(
            "photo_1_README.txt",
            WorklogAttachmentStorage.BuildStoredFileName("README.txt", WorklogAttachmentStorage.PhotoFilePrefix, 1));
    }

    [Fact]
    public void A_file_stored_name_has_its_own_underscore_between_prefix_and_id()
    {
        Assert.Equal(
            "file_2_Test.txt",
            WorklogAttachmentStorage.BuildStoredFileName("Test.txt", WorklogAttachmentStorage.FileFilePrefix, 2));
    }

    // The bug this prevents: two photos picked from different folders that are both "IMG_1234.jpg"
    // used to land on the same stored name, and the second overwrote the first one's bytes while
    // both rows sat in the list pointing at it.
    [Fact]
    public void The_same_source_name_under_two_ids_gives_two_different_stored_names()
    {
        string first = WorklogAttachmentStorage.BuildStoredFileName("IMG_1234.jpg", WorklogAttachmentStorage.PhotoFilePrefix, 1);
        string second = WorklogAttachmentStorage.BuildStoredFileName("IMG_1234.jpg", WorklogAttachmentStorage.PhotoFilePrefix, 2);

        Assert.NotEqual(first, second);
    }

    // Photos and Files number their ids independently, so photo #3 and file #3 both exist and would
    // collide in the folder they share without the differing prefix.
    [Fact]
    public void A_photo_and_a_file_with_the_same_id_do_not_collide()
    {
        string photo = WorklogAttachmentStorage.BuildStoredFileName("notes.png", WorklogAttachmentStorage.PhotoFilePrefix, 3);
        string file = WorklogAttachmentStorage.BuildStoredFileName("notes.png", WorklogAttachmentStorage.FileFilePrefix, 3);

        Assert.NotEqual(photo, file);
    }

    // Replacing a photo's image reuses its id, so the name is stable - that is what lets the copy
    // overwrite the old bytes in place instead of orphaning them.
    [Fact]
    public void The_same_id_and_source_name_always_give_the_same_stored_name()
    {
        string first = WorklogAttachmentStorage.BuildStoredFileName("board.png", WorklogAttachmentStorage.PhotoFilePrefix, 7);
        string second = WorklogAttachmentStorage.BuildStoredFileName("board.png", WorklogAttachmentStorage.PhotoFilePrefix, 7);

        Assert.Equal(first, second);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void A_path_with_no_file_name_falls_back_to_a_default(string? path)
    {
        string name = WorklogAttachmentStorage.BuildStoredFileName(path, WorklogAttachmentStorage.PhotoFilePrefix, 1);

        Assert.Equal("photo_1_attachment", name);
    }

    // A dropped file's name comes from wherever it was dragged from, so it is not assumed to be a
    // legal name here. A separator is the dangerous case: the stored name is combined into a path,
    // so leaving one in would write outside the attachments folder.
    [Fact]
    public void A_name_carrying_a_directory_separator_is_sanitized()
    {
        string name = WorklogAttachmentStorage.BuildStoredFileName("../../evil.png", WorklogAttachmentStorage.PhotoFilePrefix, 1);

        Assert.DoesNotContain("..", name);
        Assert.DoesNotContain("/", name);
        Assert.DoesNotContain("\\", name);
    }

    // ------------------------------------------------------------------ DeleteAttachmentFile

    [Fact]
    public void Deleting_removes_the_file_from_the_folder()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);
        string path = Path.Combine(folder, "1_board.png");
        File.WriteAllText(path, "bytes");

        bool deleted = WorklogAttachmentStorage.DeleteAttachmentFile(folder, "1_board.png");

        Assert.True(deleted);
        Assert.False(File.Exists(path));
    }

    // The caller's goal is that the file is gone; one already removed outside the app satisfies
    // that, and the row must still be deletable from the list.
    [Fact]
    public void Deleting_an_already_missing_file_counts_as_success()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);

        Assert.True(WorklogAttachmentStorage.DeleteAttachmentFile(folder, "never-existed.png"));
    }

    // A stored name that escapes its folder (hand-edited entries.json, or a record written before
    // names were sanitized) must not delete anything outside the attachments folder.
    [Fact]
    public void Deleting_refuses_a_name_that_escapes_the_attachments_folder()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);

        string outsidePath = Path.Combine(workspace.Root, "important.txt");
        File.WriteAllText(outsidePath, "must survive");

        bool deleted = WorklogAttachmentStorage.DeleteAttachmentFile(folder, Path.Combine("..", "important.txt"));

        Assert.False(deleted);
        Assert.True(File.Exists(outsidePath));
    }

    [Theory]
    [InlineData(null, "photo.png")]
    [InlineData("folder", null)]
    [InlineData("folder", "")]
    public void Deleting_with_a_missing_folder_or_name_is_refused(string? folder, string? fileName)
    {
        Assert.False(WorklogAttachmentStorage.DeleteAttachmentFile(folder, fileName));
    }

    // ------------------------------------------------------------------ DeleteAttachmentFileAndFolderIfEmpty

    // The behaviour this method adds over plain DeleteAttachmentFile: removing an entry's LAST
    // attachment must not leave an empty "entry-{id}-files" folder sitting on disk forever.
    [Fact]
    public void Deleting_the_last_attachment_in_a_folder_removes_the_folder_too()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "photo_1_board.png"), "bytes");

        bool deleted = WorklogAttachmentStorage.DeleteAttachmentFileAndFolderIfEmpty(folder, "photo_1_board.png");

        Assert.True(deleted);
        Assert.False(Directory.Exists(folder));
    }

    // Photos and Files share ONE folder per entry, so deleting a photo must NOT remove the folder
    // while a file (or another photo) is still in it - only the file that was just deleted is gone.
    [Fact]
    public void The_folder_survives_when_another_attachment_is_still_in_it()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "photo_1_board.png"), "bytes");
        File.WriteAllText(Path.Combine(folder, "file_1_manual.pdf"), "bytes");

        bool deleted = WorklogAttachmentStorage.DeleteAttachmentFileAndFolderIfEmpty(folder, "photo_1_board.png");

        Assert.True(deleted);
        Assert.True(Directory.Exists(folder));
        Assert.True(File.Exists(Path.Combine(folder, "file_1_manual.pdf")));
        Assert.False(File.Exists(Path.Combine(folder, "photo_1_board.png")));
    }

    // A file already missing (removed outside the app) still counts as deleted, and an otherwise
    // empty folder is still cleaned up - matching DeleteAttachmentFile's own "already gone is
    // success" rule.
    [Fact]
    public void An_already_missing_file_in_an_otherwise_empty_folder_still_removes_the_folder()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);

        bool deleted = WorklogAttachmentStorage.DeleteAttachmentFileAndFolderIfEmpty(folder, "never-existed.png");

        Assert.True(deleted);
        Assert.False(Directory.Exists(folder));
    }

    // A refused delete (the escape-the-folder guard) must not remove a folder that still holds the
    // file it refused to touch.
    [Fact]
    public void A_refused_delete_does_not_remove_the_folder()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "photo_1_board.png"), "bytes");

        bool deleted = WorklogAttachmentStorage.DeleteAttachmentFileAndFolderIfEmpty(
            folder, Path.Combine("..", "important.txt"));

        Assert.False(deleted);
        Assert.True(Directory.Exists(folder));
    }

    // A folder that never existed (or was already removed) is simply left alone - nothing to
    // delete and nothing to clean up.
    [Fact]
    public void A_nonexistent_folder_is_handled_without_throwing()
    {
        using var workspace = new TempWorkspace();
        string folder = Path.Combine(workspace.Root, "entry-1-files");

        bool deleted = WorklogAttachmentStorage.DeleteAttachmentFileAndFolderIfEmpty(folder, "board.png");

        Assert.True(deleted);
        Assert.False(Directory.Exists(folder));
    }

    // ------------------------------------------------------------------- ReorderAttachment

    private static List<WorklogAttachmentRecord> BuildAttachments(int count)
    {
        var attachments = new List<WorklogAttachmentRecord>();
        for (int i = 0; i < count; i++)
        {
            attachments.Add(new WorklogAttachmentRecord { Id = i + 1, FileName = $"photo{i + 1}.png", DisplayOrder = i });
        }

        return attachments;
    }

    private static string OrderOf(List<WorklogAttachmentRecord> attachments) =>
        string.Join(",", attachments.OrderBy(a => a.DisplayOrder).Select(a => a.Id));

    [Fact]
    public void Dragging_a_row_down_puts_it_at_the_dropped_position()
    {
        var attachments = BuildAttachments(4);

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 1, targetIndex: 2);

        Assert.Equal("2,3,1,4", OrderOf(attachments));
    }

    [Fact]
    public void Dragging_a_row_up_puts_it_at_the_dropped_position()
    {
        var attachments = BuildAttachments(4);

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 4, targetIndex: 1);

        Assert.Equal("1,4,2,3", OrderOf(attachments));
    }

    // A drag flung past the end means "put it last", which is what the user is asking for rather
    // than an error to discard.
    [Fact]
    public void A_target_past_the_end_lands_at_the_end()
    {
        var attachments = BuildAttachments(3);

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 1, targetIndex: 99);

        Assert.Equal("2,3,1", OrderOf(attachments));
    }

    [Fact]
    public void A_negative_target_lands_at_the_start()
    {
        var attachments = BuildAttachments(3);

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 3, targetIndex: -5);

        Assert.Equal("3,1,2", OrderOf(attachments));
    }

    // Dropping a row back where it started must not disturb the order.
    [Fact]
    public void Dropping_a_row_on_itself_changes_nothing()
    {
        var attachments = BuildAttachments(3);

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 2, targetIndex: 1);

        Assert.Equal("1,2,3", OrderOf(attachments));
    }

    // DisplayOrder is renumbered densely from 0, so an entry whose stored orders were sparse (or
    // duplicated by an older version) comes out clean rather than preserving the gaps.
    [Fact]
    public void Reordering_renumbers_display_order_densely_from_zero()
    {
        var attachments = new List<WorklogAttachmentRecord>
        {
            new() { Id = 1, DisplayOrder = 5 },
            new() { Id = 2, DisplayOrder = 40 },
            new() { Id = 3, DisplayOrder = 40 }
        };

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 1, targetIndex: 2);

        Assert.Equal(new[] { 0, 1, 2 }, attachments.OrderBy(a => a.DisplayOrder).Select(a => a.DisplayOrder).ToArray());
    }

    [Fact]
    public void Reordering_an_unknown_id_changes_nothing()
    {
        var attachments = BuildAttachments(3);

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 99, targetIndex: 0);

        Assert.Equal("1,2,3", OrderOf(attachments));
    }

    [Fact]
    public void Reordering_a_single_item_list_is_a_no_op()
    {
        var attachments = BuildAttachments(1);

        WorklogAttachmentStorage.ReorderAttachment(attachments, id: 1, targetIndex: 3);

        Assert.Equal("1", OrderOf(attachments));
    }

    // ---------------------------------------------------------------- NormalizeDisplayOrder

    // The state an older build could leave behind: its add path started DisplayOrder at 1 while its
    // reorder renumbered from 0, so a photo added after a reorder took an order another row already
    // held. Two rows sharing an order sort arbitrarily, so the list order could change per session.
    [Fact]
    public void Duplicate_display_orders_are_renumbered_densely()
    {
        var attachments = new List<WorklogAttachmentRecord>
        {
            new() { Id = 1, DisplayOrder = 0 },
            new() { Id = 2, DisplayOrder = 1 },
            new() { Id = 3, DisplayOrder = 1 }
        };

        bool changed = WorklogAttachmentStorage.NormalizeDisplayOrder(attachments);

        Assert.True(changed);
        Assert.Equal(new[] { 0, 1, 2 }, attachments.OrderBy(a => a.DisplayOrder).Select(a => a.DisplayOrder).ToArray());
    }

    [Fact]
    public void Gapped_display_orders_are_closed_up()
    {
        var attachments = new List<WorklogAttachmentRecord>
        {
            new() { Id = 1, DisplayOrder = 5 },
            new() { Id = 2, DisplayOrder = 40 },
            new() { Id = 3, DisplayOrder = 900 }
        };

        WorklogAttachmentStorage.NormalizeDisplayOrder(attachments);

        Assert.Equal("1,2,3", OrderOf(attachments));
        Assert.Equal(new[] { 0, 1, 2 }, attachments.OrderBy(a => a.DisplayOrder).Select(a => a.DisplayOrder).ToArray());
    }

    // Renumbering must not reshuffle rows that were already in a sensible order.
    [Fact]
    public void Normalizing_preserves_the_existing_relative_order()
    {
        var attachments = new List<WorklogAttachmentRecord>
        {
            new() { Id = 7, DisplayOrder = 2 },
            new() { Id = 4, DisplayOrder = 9 },
            new() { Id = 1, DisplayOrder = 0 }
        };

        WorklogAttachmentStorage.NormalizeDisplayOrder(attachments);

        Assert.Equal("1,7,4", OrderOf(attachments));
    }

    // Reported so the caller can skip a save it does not need.
    [Fact]
    public void An_already_dense_list_reports_no_change()
    {
        var attachments = BuildAttachments(4);

        Assert.False(WorklogAttachmentStorage.NormalizeDisplayOrder(attachments));
    }

    [Fact]
    public void Normalizing_an_empty_list_reports_no_change()
    {
        Assert.False(WorklogAttachmentStorage.NormalizeDisplayOrder(new List<WorklogAttachmentRecord>()));
    }

    // ------------------------------------------------------------------ AllocateAttachmentId

    private static List<WorklogAttachmentRecord> Records(params int[] ids) =>
        ids.Select(id => new WorklogAttachmentRecord { Id = id, FileName = $"photo_{id}_x.png" }).ToList();

    [Fact]
    public void The_first_attachment_takes_id_one()
    {
        Assert.Equal(1, WorklogAttachmentStorage.AllocateAttachmentId(
            new List<WorklogAttachmentRecord>(), WorklogAttachmentStorage.PhotoFilePrefix, Array.Empty<string>()));
    }

    [Fact]
    public void A_new_attachment_takes_one_past_the_highest_id()
    {
        Assert.Equal(4, WorklogAttachmentStorage.AllocateAttachmentId(
            Records(1, 3), WorklogAttachmentStorage.PhotoFilePrefix, Array.Empty<string>()));
    }

    // The reason this is not just Max(Id) + 1. Deleting the highest-numbered attachment frees its
    // id, so the next add reuses it - and the stored name is BUILT from the id. If the delete's
    // metadata save failed, the old bytes are still in the folder under exactly the name the new
    // attachment is about to claim, and the copy overwrites them silently.
    [Fact]
    public void An_id_whose_file_is_still_in_the_folder_is_skipped()
    {
        // Attachment #2 was removed from the list, but "photo_2_old.png" never made it off disk.
        var id = WorklogAttachmentStorage.AllocateAttachmentId(
            Records(1),
            WorklogAttachmentStorage.PhotoFilePrefix,
            new[] { "photo_1_a.png", "photo_2_old.png" });

        Assert.Equal(3, id);
    }

    // Several orphans in a row must all be stepped over, not just the first.
    [Fact]
    public void Consecutive_orphaned_ids_are_all_skipped()
    {
        var id = WorklogAttachmentStorage.AllocateAttachmentId(
            Records(1),
            WorklogAttachmentStorage.PhotoFilePrefix,
            new[] { "photo_2_a.png", "photo_3_b.png", "photo_4_c.png" });

        Assert.Equal(5, id);
    }

    // The two lists share a folder, so a photo's orphan must not push a file's id along - the
    // prefix is what keeps them independent.
    [Fact]
    public void An_orphan_of_the_other_kind_does_not_affect_the_id()
    {
        var id = WorklogAttachmentStorage.AllocateAttachmentId(
            Records(1),
            WorklogAttachmentStorage.FileFilePrefix,
            new[] { "photo_2_orphan.png" });

        Assert.Equal(2, id);
    }

    // An orphan stored in the OLDER "photo2_" form (before an underscore was added between the
    // prefix and the id) is just as much attachment #2's file as "photo_2_" is - GetDisplayFileName
    // reads it back as such. So it has to block id 2 too, or the overwrite this whole scan exists to
    // prevent comes straight back for every attachment stored before the naming change, AND the
    // orphan and the live row then render the SAME display name with nothing to tell them apart.
    [Fact]
    public void An_id_whose_file_is_in_the_older_prefix_form_is_skipped()
    {
        var id = WorklogAttachmentStorage.AllocateAttachmentId(
            Records(1),
            WorklogAttachmentStorage.PhotoFilePrefix,
            new[] { "photo_1_a.png", "photo2_old.png" });

        Assert.Equal(3, id);
    }

    // The same for the oldest form of all, from the build where PhotoFilePrefix was empty and photos
    // were stored bare as "2_board.png". GetDisplayFileName still strips those, so this scan must
    // still block them.
    [Fact]
    public void An_id_whose_file_is_in_the_bare_prefixless_form_is_skipped()
    {
        var id = WorklogAttachmentStorage.AllocateAttachmentId(
            Records(1),
            WorklogAttachmentStorage.PhotoFilePrefix,
            new[] { "photo_1_a.png", "2_old.png" });

        Assert.Equal(3, id);
    }

    // Every form GetDisplayFileName reads back must block, and it must keep stepping over them in a
    // run - a mixture of the three is exactly what a folder written across several app versions
    // holds.
    [Fact]
    public void A_mixture_of_old_and_new_stored_forms_is_all_skipped()
    {
        var id = WorklogAttachmentStorage.AllocateAttachmentId(
            Records(1),
            WorklogAttachmentStorage.PhotoFilePrefix,
            new[] { "photo2_a.png", "3_b.png", "photo_4_c.png" });

        Assert.Equal(5, id);
    }
}
