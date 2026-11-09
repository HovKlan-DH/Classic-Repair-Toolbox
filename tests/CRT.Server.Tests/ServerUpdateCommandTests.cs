using CRT.Server.Handlers.Updates;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // `CRT.Server update-server`'s command line, and the arguments it hands gh and reads back from gh and
    // systemd (ServerUpdateCommand, ServerUpdateHost; 2026-11-09). The command line is the unit's, which
    // only root may change - still, a value that could point the updater anywhere unexpected is refused.
    // ###########################################################################################
    public sealed class ServerUpdateCommandTests
    {
        private static readonly string Folder = Path.Combine(Path.GetTempPath(), "crt-server");

        private static List<string> Args(params string[] changes)
        {
            var args = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["--channel"] = "releases",
                ["--folder"] = ServerUpdateCommandTests.Folder,
                ["--service"] = "crt-server.service",
                ["--health"] = "http://127.0.0.1:5199/api/health",
            };

            foreach (string change in changes)
            {
                int equals = change.IndexOf('=');
                string name = equals < 0 ? change : change[..equals];

                if (equals < 0)
                    args.Remove(name);
                else
                    args[name] = change[(equals + 1)..];
            }

            return args.Select(pair => $"{pair.Key}={pair.Value}").ToList();
        }

        [Fact]
        public void The_units_command_line_is_read()
        {
            List<string> args = ServerUpdateCommandTests.Args();
            args.Add("--dry-run");

            ServerUpdateCommand.Options? options = ServerUpdateCommand.Parse(args, out string why);

            Assert.NotNull(options);
            Assert.Equal(string.Empty, why);
            Assert.Equal(ServerUpdateOutcomes.ReleasesChannel, options.Channel);
            Assert.Equal(ServerUpdateCommandTests.Folder, options.Folder);
            Assert.Equal("crt-server.service", options.Service);
            Assert.Equal(new Uri("http://127.0.0.1:5199/api/health"), options.Health);
            Assert.Equal(ServerUpdateCommand.DefaultGh, options.Gh);
            Assert.True(options.DryRun);
            Assert.False(options.RetryFailed);
        }

        [Theory]
        // The health address is the service's own on this machine - never anywhere the network reaches.
        [InlineData("--health=http://example.com/api/health")]
        [InlineData("--health=https://127.0.0.1:5199/api/health")]
        [InlineData("--health=http://127.0.0.1:5199/api/other")]
        [InlineData("--health=http://127.0.0.1:5199/api/health?x=1")]
        [InlineData("--channel=nightly")]
        [InlineData("--service=crt-server")]
        [InlineData("--service=../crt-server.service")]
        [InlineData("--service=crt server.service")]
        [InlineData("--folder=relative/crt-server")]
        [InlineData("--folder")]
        [InlineData("--channel")]
        [InlineData("--unknown=1")]
        public void A_value_that_could_point_the_updater_anywhere_unexpected_is_refused(string change)
        {
            Assert.Null(ServerUpdateCommand.Parse(ServerUpdateCommandTests.Args(change), out string why));
            Assert.NotEqual(string.Empty, why);
        }

        [Fact]
        public void A_folder_climbing_out_and_an_option_given_twice_are_refused()
        {
            Assert.Null(ServerUpdateCommand.Parse(ServerUpdateCommandTests.Args("--folder=" + Path.Combine(ServerUpdateCommandTests.Folder, "..", "etc")), out _));

            List<string> twice = ServerUpdateCommandTests.Args();
            twice.Add("--channel=main");
            Assert.Null(ServerUpdateCommand.Parse(twice, out string why));
            Assert.Contains("twice", why, StringComparison.Ordinal);
        }

        [Fact]
        public void A_wrong_command_line_exits_with_2_and_says_how()
        {
            using var output = new StringWriter();

            Assert.Equal(ServerUpdateCommand.WrongCommandLine, ServerUpdateCommand.Run(["update-server", "--channel=nightly"], output));
            Assert.Contains("Usage:", output.ToString(), StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** WHAT gh IS ASKED TO PROVE (owner decision, 2026-11-09: "GitHub attestations"). *** The build and
        // its bundle, held to this repository, the channel's workflow - the release workflow for the live
        // service, the push workflow for the test instance - main and a GitHub-hosted runner. Loosening any of
        // these fails here.
        // ###########################################################################################
        [Theory]
        [InlineData(ServerUpdateOutcomes.ReleasesChannel, "build-and-release.yml")]
        [InlineData(ServerUpdateOutcomes.MainChannel, "build-and-unittest.yml")]
        public void Gh_is_held_to_this_repository_the_channels_workflow_main_and_a_GitHub_hosted_runner(string channel, string workflow)
        {
            Assert.Equal(
                [
                    "attestation", "verify", "/r/build.tar.gz",
                    "--bundle", "/r/build.tar.gz.sigstore.json",
                    "--repo", "HovKlan-DH/Classic-Repair-Toolbox",
                    "--signer-workflow", "HovKlan-DH/Classic-Repair-Toolbox/.github/workflows/" + workflow,
                    "--source-ref", "refs/heads/main",
                    "--deny-self-hosted-runners",
                    "--format", "json",
                ],
                ServerUpdateHost.VerifyArguments("/r/build.tar.gz", "/r/build.tar.gz.sigstore.json", channel));
        }

        // gh's answer, as `gh attestation verify --format json` writes it (measured 2026-11-09 on gh 2.98.0).
        [Fact]
        public void The_commit_is_read_from_the_certificate_gh_verified()
        {
            const string verified =
                "[{\"verificationResult\":{\"signature\":{\"certificate\":{\"sourceRepositoryDigest\":\"a255baf71d13fe5947a4eb7ad521ffd412d64cee\"," +
                "\"sourceRepositoryRef\":\"refs/heads/main\",\"runnerEnvironment\":\"github-hosted\"}}}}]";

            Assert.Equal("a255baf71d13fe5947a4eb7ad521ffd412d64cee", ServerUpdateHost.CommitOf(verified));
            Assert.Null(ServerUpdateHost.CommitOf("[]"));
            Assert.Null(ServerUpdateHost.CommitOf("not json"));
        }

        [Theory]
        [InlineData("ActiveState=active\nResult=success\nExecMainStatus=0\n", "active")]
        [InlineData("ActiveState=failed\nResult=exit-code\nExecMainStatus=78\n", "failed (exit 78)")]
        [InlineData("ActiveState=activating\nResult=success\nExecMainStatus=69\n", "activating (exit 69)")]
        [InlineData("", "unknown")]
        public void Systemds_answer_is_said_in_words(string show, string words) =>
            Assert.Equal(words, ServerUpdateHost.StateOf(show));
    }
}
