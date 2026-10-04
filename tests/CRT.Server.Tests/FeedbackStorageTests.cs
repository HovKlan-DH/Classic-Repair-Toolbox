using CRT.Server.Handlers.Feedback;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // FeedbackStorage's remembered folder total (code review, 2026-10-04): each feedback used to
    // stat every saved file to total the folder. The walk is counted and the clock moved by hand.
    // ###########################################################################################
    public sealed class FeedbackStorageTests
    {
        private DateTimeOffset thisNow = new(2026, 10, 4, 12, 0, 0, TimeSpan.Zero);

        private int thisWalks;

        private long? thisOnDisk = 1_000;

        private FeedbackStorage Storage() =>
            new(() =>
            {
                this.thisWalks++;
                return this.thisOnDisk;
            },
            () => this.thisNow);

        [Fact]
        public void The_folder_is_walked_once_and_its_total_remembered()
        {
            FeedbackStorage storage = this.Storage();

            Assert.Equal(1_000L, storage.StoredBytes());
            this.thisNow += TimeSpan.FromMinutes(1);
            Assert.Equal(1_000L, storage.StoredBytes());

            Assert.Equal(1, this.thisWalks);
        }

        // What a feedback kept counts at once, without a walk.
        [Fact]
        public void A_feedback_kept_adds_to_the_remembered_total()
        {
            FeedbackStorage storage = this.Storage();
            storage.StoredBytes();

            storage.Saved(250);

            Assert.Equal(1_250L, storage.StoredBytes());
            Assert.Equal(1, this.thisWalks);
        }

        // A feedback whose mail failed has its files deleted again: they stop counting at once, and
        // the total never goes below nothing.
        [Fact]
        public void Files_deleted_again_are_taken_off_the_remembered_total()
        {
            FeedbackStorage storage = this.Storage();
            storage.StoredBytes();

            storage.Saved(250);
            storage.Removed(250);
            Assert.Equal(1_000L, storage.StoredBytes());

            storage.Removed(5_000);
            Assert.Equal(0L, storage.StoredBytes());
            Assert.Equal(1, this.thisWalks);
        }

        // ###########################################################################################
        // The project owner deletes old feedback through a network share, which the service never
        // hears about - so the folder is walked again after RescanAfter, and room made there counts.
        // ###########################################################################################
        [Fact]
        public void The_folder_is_walked_again_once_the_total_is_old_so_room_made_by_hand_counts()
        {
            FeedbackStorage storage = this.Storage();
            storage.StoredBytes();

            this.thisOnDisk = 100;
            this.thisNow += FeedbackStorage.RescanAfter;

            Assert.Equal(100L, storage.StoredBytes());
            Assert.Equal(2, this.thisWalks);
        }

        // An unknown total is not remembered - the next feedback walks again - and nothing is added
        // to it.
        [Fact]
        public void A_folder_that_cannot_be_read_is_walked_again_next_time()
        {
            this.thisOnDisk = null;
            FeedbackStorage storage = this.Storage();

            Assert.Null(storage.StoredBytes());
            storage.Saved(250);
            Assert.Null(storage.StoredBytes());

            Assert.Equal(2, this.thisWalks);
        }
    }
}
