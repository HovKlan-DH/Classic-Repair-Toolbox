using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The device photos' one reader - DevicePhotoThumbnails, shared by a device's page and its form. HeadlessUi: a photo
// is decoded as an Avalonia bitmap, and the reads it keeps are shared by every test of the process that shows one.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DevicePhotoThumbnailsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    // ###########################################################################################
    // *** NOT EVERY VERSION OF EVERY PHOTO FOR THE RUN (code review, 2026-11-08). *** A plain dictionary kept each read
    // for good - a photo replaced kept its older read beside the new one, and a deleted device's photos stayed held. A
    // photo replaced lets its older read go, a photo asked again shares its read, and at most Kept are held.
    // ###########################################################################################
    [Fact]
    public void A_photo_replaced_lets_its_older_read_go_and_no_more_than_the_most_kept_are_held()
    {
        UiTest.Run(() =>
        {
            string photo = this.thisWorkspace.WriteFile("front.jpg", "first");

            Task<Avalonia.Media.Imaging.Bitmap?> first = DevicePhotoThumbnails.Read(photo);
            int held = DevicePhotoThumbnails.HeldForTests;

            Assert.Same(first, DevicePhotoThumbnails.Read(photo));

            // Replaced - once its read has let go of the file, as Windows refuses to write one held open: another read,
            // and the older one let go of - the photo held once.
            Assert.True(first.Wait(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken), "The photo was not read in time");
            File.WriteAllText(photo, "replaced, longer");
            File.SetLastWriteTimeUtc(photo, File.GetLastWriteTimeUtc(photo).AddMinutes(1));

            Assert.NotSame(first, DevicePhotoThumbnails.Read(photo));
            Assert.Equal(held, DevicePhotoThumbnails.HeldForTests);

            for (int each = 0; each < DevicePhotoThumbnails.Kept + 20; each++)
            {
                DevicePhotoThumbnails.Read(this.thisWorkspace.WriteFile($"photo{each}.jpg", $"photo {each}"));
            }

            Assert.Equal(DevicePhotoThumbnails.Kept, DevicePhotoThumbnails.HeldForTests);
        });
    }

    // A photo that is not there reads as nothing, at once, and holds nothing.
    [Fact]
    public void A_photo_not_there_reads_as_nothing_at_once()
    {
        UiTest.Run(() =>
        {
            int held = DevicePhotoThumbnails.HeldForTests;

            Task<Avalonia.Media.Imaging.Bitmap?> reading = DevicePhotoThumbnails.Read(this.thisWorkspace.Path_("gone.jpg"));

            Assert.True(reading.IsCompletedSuccessfully);
            Assert.Null(reading.Result);
            Assert.Equal(held, DevicePhotoThumbnails.HeldForTests);
        });
    }
}
