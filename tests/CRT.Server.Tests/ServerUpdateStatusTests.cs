using CRT.Server.Handlers.Updates;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // What the updater did, as the service finds and says it (ServerUpdateStatus, 2026-11-09): beside the
    // build the service runs, or "not set up" without one - never guessed from anywhere else.
    // ###########################################################################################
    public sealed class ServerUpdateStatusTests : IDisposable
    {
        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-update-status-" + Guid.NewGuid().ToString("N"));

        public ServerUpdateStatusTests() => Directory.CreateDirectory(this.thisRoot);

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A temp folder left behind is harmless.
            }
        }

        [Fact]
        public void The_service_finds_the_status_beside_the_build_it_runs_or_through_its_app_link()
        {
            string releases = Path.Combine(this.thisRoot, "crt-server", ServerUpdater.ReleasesFolderName);
            Directory.CreateDirectory(Path.Combine(releases, "10.3.0"));
            Directory.CreateDirectory(Path.Combine(this.thisRoot, "crt-server", "app"));
            string status = Path.Combine(releases, ServerUpdateStatus.FileName);

            Assert.Null(ServerUpdateStatus.Locate(Path.Combine(releases, "10.3.0")));

            new ServerUpdateStatus { Channel = ServerUpdateOutcomes.ReleasesChannel }.Write(status);

            // Its folder as the kernel names it (releases/10.3.0), or as the unit does (crt-server/app).
            Assert.Equal(status, ServerUpdateStatus.Locate(Path.Combine(releases, "10.3.0")));
            Assert.Equal(status, ServerUpdateStatus.Locate(Path.Combine(this.thisRoot, "crt-server", "app") + Path.DirectorySeparatorChar));
        }

        [Fact]
        public void Without_a_status_the_updates_are_not_set_up()
        {
            Assert.Equal(new ServerUpdatesAnswer(SetUp: false), ServerUpdateStatus.AnswerFrom(null));
            Assert.Null(ServerUpdateStatus.Locate(Path.Combine(this.thisRoot, "app")));
        }

        // A status that cannot be read - broken, or far larger than one can be - is said as such, never as up to date.
        [Fact]
        public void A_status_that_cannot_be_read_is_said()
        {
            string broken = Path.Combine(this.thisRoot, "broken.json");
            File.WriteAllText(broken, "{ not json");

            string large = Path.Combine(this.thisRoot, "large.json");
            File.WriteAllText(large, "{\"message\":\"" + new string('x', ServerUpdateStatus.MaxBytes) + "\"}");

            foreach (string path in new[] { broken, large })
            {
                ServerUpdatesAnswer answer = ServerUpdateStatus.AnswerFrom(path);

                Assert.True(answer.SetUp);
                Assert.Equal(ServerUpdateOutcomes.Error, answer.Outcome);
                Assert.Contains("could not be read", answer.Message, StringComparison.Ordinal);
            }
        }

        // Written whole and read back - a field a later updater writes is kept by an earlier one.
        [Fact]
        public void A_status_is_written_whole_and_keeps_what_a_later_updater_wrote()
        {
            string path = Path.Combine(this.thisRoot, ServerUpdateStatus.FileName);
            File.WriteAllText(path, "{\"channel\":\"releases\",\"outcome\":\"installed\",\"examinedSha256\":\"ab12\",\"fromLater\":{\"a\":1}}");

            ServerUpdateStatus status = ServerUpdateStatus.Read(path)!;
            status.Message = "Up to date.";
            status.Write(path);

            ServerUpdateStatus again = ServerUpdateStatus.Read(path)!;
            Assert.Equal("Up to date.", again.Message);
            Assert.Equal("ab12", again.ExaminedSha256);
            Assert.True(again.Unknown!.ContainsKey("fromLater"));
            Assert.False(File.Exists(path + ".writing"));
        }
    }
}
