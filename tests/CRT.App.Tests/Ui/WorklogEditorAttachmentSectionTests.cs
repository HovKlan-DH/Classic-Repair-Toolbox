using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The editor's Photos and Files lists, which are now ONE implementation parameterised by an
// AttachmentSection rather than two near-identical copies.
//
// Photos and Files have the same storage shape - metadata in entries.json, bytes in the entry's
// "worklog_<id>" folder - and their add/edit/delete paths were written out twice, each encoding
// the same three ordering rules (undo the copy when the save fails; persist before swapping a
// file, because the swap deletes what it replaces; persist before deleting bytes). Every one of
// those rules was learned from a real fault, and holding two copies of them meant two places for
// one to be got wrong. They had already started to drift: the file-side copies carried "see the
// photo path" comments instead of the reasoning itself.
//
// Sharing the implementation trades that risk for a different one, and THAT is what these tests
// are for. A section pointed at the wrong list or the wrong filename prefix compiles perfectly and
// fails silently - Files rendering the Photos records, or two lists minting colliding stored names.
// So these assert the wiring itself, per section, rather than re-testing the fields of a record.
//
// The storage rules underneath (id allocation, stored-name shape, orphan cleanup) belong to
// WorklogAttachmentStorageTests and are not repeated here.
[Collection("HeadlessUi")]
public class WorklogEditorAttachmentSectionTests
{
    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static void WithEditor(Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow { Width = 1000, Height = 700 };

            using var bitmap = CreateBitmap();

            // A draft, exactly as WorklogEditorNewEntryTests does it: workbook 0 does not exist, so
            // nothing here touches the user's real Workbooks folder.
            window.InitializeForNewEntry(0, "Sheet 1", new Rect(10, 20, 30, 40), bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
                Dispatcher.UIThread.RunJobs();
            }
        });
    }

    // The two sections must not share a filename prefix. The stored name carries it (see
    // BuildStoredFileName), so if both sections minted "photo1_..." a file and a photo allocated
    // the same id would resolve to the same path on disk and one would overwrite the other.
    [Fact]
    public void The_two_attachment_sections_use_different_stored_name_prefixes()
    {
        WithEditor(window =>
        {
            var photos = window.DescribeAttachmentSectionForTests(photos: true);
            var files = window.DescribeAttachmentSectionForTests(photos: false);

            Assert.Equal("photo", photos.Prefix);
            Assert.Equal("file", files.Prefix);
            Assert.NotEqual(photos.Prefix, files.Prefix);
        });
    }

    // Each section must drive its OWN controls. Pointing both at one list's controls is the
    // copy/paste slip that unifying the implementation makes possible, and it would show up as the
    // Files count overwriting the Photos count rather than as any kind of error.
    [Fact]
    public void Each_attachment_section_drives_its_own_controls()
    {
        WithEditor(window =>
        {
            var photos = window.DescribeAttachmentSectionForTests(photos: true);
            var files = window.DescribeAttachmentSectionForTests(photos: false);

            Assert.NotSame(photos.List, files.List);
            Assert.NotSame(photos.EmptyText, files.EmptyText);
            Assert.NotSame(photos.CountText, files.CountText);
            Assert.NotEqual(photos.HeaderKey, files.HeaderKey);
        });
    }

    // Only the photo section loads thumbnails. That branch is the one real difference between the
    // two lists and the only reason the shared rebuild carries a conditional at all.
    [Fact]
    public void Only_the_photo_section_is_of_the_photo_kind()
    {
        WithEditor(window =>
        {
            Assert.True(window.DescribeAttachmentSectionForTests(photos: true).IsPhotoKind);
            Assert.False(window.DescribeAttachmentSectionForTests(photos: false).IsPhotoKind);
        });
    }

    // The sections must read SEPARATE record lists. Wiring the Files section to thisEntry.Photos
    // would make every photo appear in both lists, and a delete from one would empty the other.
    [Fact]
    public void Adding_to_one_attachment_section_does_not_change_the_other()
    {
        WithEditor(window =>
        {
            window.AddAttachmentRecordForTests(photos: true, 1, "photo1_board.png", "front");

            Assert.Equal(1, window.DescribeAttachmentSectionForTests(photos: true).RecordCount);
            Assert.Equal(0, window.DescribeAttachmentSectionForTests(photos: false).RecordCount);

            window.AddAttachmentRecordForTests(photos: false, 1, "file1_datasheet.pdf", "spec");

            Assert.Equal(1, window.DescribeAttachmentSectionForTests(photos: true).RecordCount);
            Assert.Equal(1, window.DescribeAttachmentSectionForTests(photos: false).RecordCount);
        });
    }

    // The shared rebuild has to serve both lists - that is the whole claim of the refactor. Each
    // section reads its own records, and the count text is worded for that section.
    [Fact]
    public void The_shared_rebuild_populates_each_section_from_its_own_records()
    {
        WithEditor(window =>
        {
            window.AddAttachmentRecordForTests(photos: true, 1, "photo1_board.png", "front");
            window.AddAttachmentRecordForTests(photos: true, 2, "photo2_back.png", "back");
            window.AddAttachmentRecordForTests(photos: false, 1, "file1_datasheet.pdf", "spec");

            window.RefreshAttachmentRowsForTests(photos: true);
            window.RefreshAttachmentRowsForTests(photos: false);

            var photos = window.DescribeAttachmentSectionForTests(photos: true);
            var files = window.DescribeAttachmentSectionForTests(photos: false);

            Assert.Equal(2, photos.RowCount);
            Assert.Equal(1, files.RowCount);

            // Singular/plural come from the section, so a shared rebuild cannot label a file list
            // "photos" - the kind of wording slip a single implementation invites.
            Assert.Equal("2 photos", photos.CountText.Text);
            Assert.Equal("1 file", files.CountText.Text);
        });
    }

    // Rebuilding is idempotent: it clears before it fills. Without the clear, the shared method
    // would append on every add/edit/delete/reorder and duplicate every row.
    [Fact]
    public void Rebuilding_a_section_replaces_its_rows_rather_than_appending()
    {
        WithEditor(window =>
        {
            window.AddAttachmentRecordForTests(photos: false, 1, "file1_datasheet.pdf", "spec");

            window.RefreshAttachmentRowsForTests(photos: false);
            window.RefreshAttachmentRowsForTests(photos: false);
            window.RefreshAttachmentRowsForTests(photos: false);

            Assert.Equal(1, window.DescribeAttachmentSectionForTests(photos: false).RowCount);
        });
    }

    // An empty section shows its "none added" text; a populated one does not. Both sections share
    // one line of code for this now, so it is asserted for each.
    [Fact]
    public void The_empty_state_text_follows_each_sections_own_record_count()
    {
        WithEditor(window =>
        {
            window.RefreshAttachmentRowsForTests(photos: true);
            window.RefreshAttachmentRowsForTests(photos: false);

            Assert.True(window.DescribeAttachmentSectionForTests(photos: true).EmptyText.IsVisible);
            Assert.True(window.DescribeAttachmentSectionForTests(photos: false).EmptyText.IsVisible);

            window.AddAttachmentRecordForTests(photos: true, 1, "photo1_board.png", string.Empty);
            window.RefreshAttachmentRowsForTests(photos: true);

            // The photo list has content now; the file list is untouched and still empty.
            Assert.False(window.DescribeAttachmentSectionForTests(photos: true).EmptyText.IsVisible);
            Assert.True(window.DescribeAttachmentSectionForTests(photos: false).EmptyText.IsVisible);
        });
    }
}
