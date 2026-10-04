using CRT.Server.Configuration;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the rule that decides whether the service may start at all.
    //
    // This is the most consequential pure logic in Phase 3: it is the in-code half of the
    // interlock that keeps the service off the production data tree, and it is what turns a
    // missing setting into a refusal to start rather than a silent guess.
    //
    // WHAT THESE TESTS DO NOT CLAIM. Passing validation does not make a production write
    // impossible - only the filesystem permissions (DEPLOYMENT.md step 3) and systemd's
    // ProtectSystem=strict do that. These tests pin the early, loud failure for honest
    // misconfiguration. Read ServerOptionsValidator's header before assuming a check here is the
    // security control.
    //
    // Paths in these tests are built with Path.Combine and Path.GetTempPath rather than written as
    // literals, so they pass on the Linux CI runner as well as on Windows.
    // ###########################################################################################
    public class ServerOptionsValidatorTests
    {
        // A configuration with everything set correctly, which individual tests then break in one
        // specific way. Building valid-by-default and breaking one thing is what makes each test
        // name true: a failure then names the single cause.
        private static ServerOptions ValidOptions(string betaRoot, string productionRoot)
        {
            return new ServerOptions
            {
                DataTreeRoot = betaRoot,
                ProductionTreeRoot = productionRoot,
                PublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data-BETA/Data",
                ManifestPath = Path.Combine(betaRoot, "dataChecksums.json"),
                ConnectionString = "Server=localhost;Database=crt_review;Uid=crt_review;Pwd=x;",
                PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
                MailFromAddress = "noreply@classic-repair-toolbox.dk",
                BlobStoreRoot = Path.Combine(Path.GetTempPath(), "crt-blobs"),
                FeedbackRoot = Path.Combine(Path.GetTempPath(), "crt-test", "user-feedback"),
                FeedbackToAddress = "dennis@classic-repair-toolbox.dk"
            };
        }

        private static IReadOnlyList<string> Validate(
            ServerOptions options,
            bool exists = true,
            bool writable = true)
        {
            return ServerOptionsValidator.Validate(options, _ => exists, _ => writable);
        }

        private static string BetaPath() =>
            Path.Combine(Path.GetTempPath(), "crt-test", "app-data-BETA", "Data");

        private static string ProductionPath() =>
            Path.Combine(Path.GetTempPath(), "crt-test", "app-data", "Data");

        // -----------------------------------------------------------------------------------
        // The happy path.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_fully_configured_service_passes_validation()
        {
            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(
                ServerOptionsValidatorTests.ValidOptions(
                    ServerOptionsValidatorTests.BetaPath(),
                    ServerOptionsValidatorTests.ProductionPath()));

            Assert.Empty(failures);
        }

        // -----------------------------------------------------------------------------------
        // The no-default rule. Each of these must REFUSE, not fall back to anything.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void A_missing_data_tree_root_refuses_to_start(string? value)
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.DataTreeRoot = value;

            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(options);

            Assert.Contains(failures, f => f.Contains("DataTreeRoot", StringComparison.Ordinal));
        }

        // The message is what appears in "systemctl status" when the unit fails, and DEPLOYMENT.md
        // promises the project owner that a refusal names the setting. If that stops being true, the
        // runbook is lying.
        [Fact]
        public void A_refusal_names_the_setting_that_caused_it()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.DataTreeRoot = null;

            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(options);

            Assert.Contains(failures, f =>
                f.Contains($"{ServerOptions.SectionName}:DataTreeRoot", StringComparison.Ordinal));
        }

        [Fact]
        public void A_missing_connection_string_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.ConnectionString = null;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("ConnectionString", StringComparison.Ordinal));
        }

        [Fact]
        public void A_missing_production_tree_root_refuses_to_start()
        {
            // Leaving this blank removes a safety check rather than disabling a feature, so it is
            // a failure rather than an accepted omission.
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.ProductionTreeRoot = null;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("ProductionTreeRoot", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // The production interlock - the reason this class exists.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Pointing_the_data_tree_at_production_refuses_to_start()
        {
            string production = ServerOptionsValidatorTests.ProductionPath();
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(production, production);
            // Keep the other markers valid so the ONLY failure can be the production collision.
            options.RequiredTreeMarker = string.Empty;
            options.PublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data/Data";
            options.ManifestPath = Path.Combine(production, "dataChecksums.json");

            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(options);

            Assert.Contains(failures, f => f.Contains("same directory", StringComparison.Ordinal));
        }

        [Fact]
        public void A_data_tree_inside_the_production_tree_refuses_to_start()
        {
            string production = Path.Combine(Path.GetTempPath(), "crt-test", "app-data");
            string inside = Path.Combine(production, "sub", "Data");

            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(inside, production);
            options.RequiredTreeMarker = string.Empty;
            options.PublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data/Data";
            options.ManifestPath = Path.Combine(inside, "dataChecksums.json");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("is inside", StringComparison.Ordinal));
        }

        [Fact]
        public void A_production_tree_inside_the_data_tree_refuses_to_start()
        {
            // The reverse containment matters too: if production sits under the data root, a write
            // anywhere below that root could reach it.
            string dataRoot = Path.Combine(Path.GetTempPath(), "crt-test", "trees-BETA");
            string production = Path.Combine(dataRoot, "app-data");

            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(dataRoot, production);
            options.ManifestPath = Path.Combine(dataRoot, "dataChecksums.json");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("is inside", StringComparison.Ordinal));
        }

        // The classic prefix bug: "/srv/app-data-BETA" starts with "/srv/app-data" as a STRING but
        // is not inside it as a DIRECTORY. Comparing without a separator boundary would refuse this
        // perfectly good configuration - and, with the arguments reversed, would accept a bad one.
        [Fact]
        public void A_sibling_tree_whose_name_merely_starts_the_same_is_accepted()
        {
            string production = Path.Combine(Path.GetTempPath(), "crt-test", "app-data");
            string beta = Path.Combine(Path.GetTempPath(), "crt-test", "app-data-BETA");

            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(beta, production);
            options.ManifestPath = Path.Combine(beta, "dataChecksums.json");

            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(options);

            Assert.DoesNotContain(failures, f => f.Contains("is inside", StringComparison.Ordinal));
            Assert.Empty(failures);
        }

        // Two spellings of one directory must compare equal, or the interlock is bypassed by
        // writing the same path with a trailing separator or a "..".
        [Fact]
        public void Different_spellings_of_the_production_path_are_still_refused()
        {
            string production = ServerOptionsValidatorTests.ProductionPath();
            string sameWithDotDot = Path.Combine(production, "..", "Data");

            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(sameWithDotDot, production);
            options.RequiredTreeMarker = string.Empty;
            options.PublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data/Data";
            options.ManifestPath = Path.Combine(production, "dataChecksums.json");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("same directory", StringComparison.Ordinal));
        }

        [Fact]
        public void A_trailing_separator_does_not_hide_the_production_path()
        {
            string production = ServerOptionsValidatorTests.ProductionPath();

            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                production + Path.DirectorySeparatorChar, production);
            options.RequiredTreeMarker = string.Empty;
            options.PublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data/Data";
            options.ManifestPath = Path.Combine(production, "dataChecksums.json");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("same directory", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // The marker guard - belt and braces on top of the path comparison.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_data_tree_without_the_required_marker_refuses_to_start()
        {
            string noMarker = Path.Combine(Path.GetTempPath(), "crt-test", "somewhere", "Data");

            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                noMarker, ServerOptionsValidatorTests.ProductionPath());
            options.ManifestPath = Path.Combine(noMarker, "dataChecksums.json");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("required marker", StringComparison.Ordinal));
        }

        // Clearing the marker is how the project owner deliberately targets a tree without "-BETA" in
        // its name. It must be possible, or the service could never be pointed anywhere else - but
        // it has to be an explicit act, which is why the default is not empty.
        [Fact]
        public void Clearing_the_marker_deliberately_allows_a_tree_without_it()
        {
            string noMarker = Path.Combine(Path.GetTempPath(), "crt-test", "somewhere", "Data");

            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                noMarker, ServerOptionsValidatorTests.ProductionPath());
            options.RequiredTreeMarker = string.Empty;
            options.PublicDataBaseUrl = "https://classic-repair-toolbox.dk/somewhere/Data";
            options.ManifestPath = Path.Combine(noMarker, "dataChecksums.json");

            Assert.Empty(ServerOptionsValidatorTests.Validate(options));
        }

        // -----------------------------------------------------------------------------------
        // Filesystem state.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_data_tree_that_does_not_exist_refuses_to_start()
        {
            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(
                ServerOptionsValidatorTests.ValidOptions(
                    ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath()),
                exists: false);

            Assert.Contains(failures, f => f.Contains("does not exist", StringComparison.Ordinal));
        }

        [Fact]
        public void A_data_tree_the_service_cannot_write_refuses_to_start()
        {
            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(
                ServerOptionsValidatorTests.ValidOptions(
                    ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath()),
                writable: false);

            Assert.Contains(failures, f => f.Contains("not writable", StringComparison.Ordinal));
        }

        [Fact]
        public void A_relative_data_tree_path_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.DataTreeRoot = Path.Combine("relative", "app-data-BETA", "Data");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("absolute path", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // The manifest settings. Nothing reads these until Phase 5; they are validated now
        // because the failure they prevent is invisible - a well-formed manifest pointing every
        // client at the wrong tree.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_public_data_url_for_the_wrong_tree_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            // The BETA data root paired with the PRODUCTION base URL - the exact mismatch that
            // would publish a valid manifest sending every client to production content.
            options.PublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data/Data";

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("PublicDataBaseUrl", StringComparison.Ordinal));
        }

        [Fact]
        public void A_non_https_public_data_url_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.PublicDataBaseUrl = "http://classic-repair-toolbox.dk/app-data-BETA/Data";

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("https", StringComparison.Ordinal));
        }

        // Same defect as the data tree: Path.GetFullPath silently resolves a relative path against
        // the working directory, so a check made after normalisation can never fire. Both checks
        // therefore test the raw value.
        [Fact]
        public void A_relative_manifest_path_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.ManifestPath = Path.Combine("app-data-BETA", "dataChecksums.json");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("ManifestPath", StringComparison.Ordinal) &&
                     f.Contains("absolute path", StringComparison.Ordinal));
        }

        [Fact]
        public void A_missing_manifest_path_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.ManifestPath = null;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("ManifestPath", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // Mail and API links.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_non_https_api_base_url_refuses_to_start()
        {
            // Password-reset links are built from this. Over http they are an account takeover for
            // anyone on the same network.
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.PublicApiBaseUrl = "http://classic-repair-toolbox.dk/api";

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("PublicApiBaseUrl", StringComparison.Ordinal));
        }

        [Fact]
        public void A_missing_api_base_url_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.PublicApiBaseUrl = null;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("PublicApiBaseUrl", StringComparison.Ordinal));
        }

        [Fact]
        public void A_missing_blob_store_root_refuses_to_start()
        {
            // No default, because this directory accumulates contributor-supplied bytes: a guessed
            // location could fill a partition nobody is watching.
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.BlobStoreRoot = null;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("BlobStoreRoot", StringComparison.Ordinal));
        }

        [Fact]
        public void A_relative_blob_store_root_refuses_to_start()
        {
            // A relative path resolves against the service's working directory - which is where
            // the binaries live, not where submissions belong.
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.BlobStoreRoot = "blobs";

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("BlobStoreRoot", StringComparison.Ordinal));
        }

        [Fact]
        public void A_missing_mail_from_address_refuses_to_start()
        {
            // Required from the accounts step onward. Without this check a missing From address
            // passes validation and then fails when somebody registers - inside a background
            // send, where the failure is a log line nobody is watching rather than a refusal to
            // start.
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.MailFromAddress = null;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("MailFromAddress", StringComparison.Ordinal));
        }

        [Fact]
        public void An_smtp_port_outside_the_valid_range_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.SmtpPort = 0;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("SmtpPort", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // The Argon2 cost floor. This is where the PRODUCTION hashing parameters are asserted -
        // the hasher's own tests run at tiny parameters to keep the suite fast, so without this
        // the real cost would be pinned by nothing.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_default_hashing_parameters_meet_the_floor()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());

            // Untouched: whatever ServerOptions defaults to must itself be acceptable.
            Assert.Empty(ServerOptionsValidatorTests.Validate(options));
        }

        [Fact]
        public void Lowering_the_hashing_memory_below_the_floor_refuses_to_start()
        {
            // Memory is what makes an offline attack on a stolen hash expensive. Quietly lowering
            // it is the single change that most undoes password hashing without looking wrong.
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.Argon2MemoryKib = 1024;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("Argon2MemoryKib", StringComparison.Ordinal));
        }

        [Fact]
        public void Lowering_the_hashing_iterations_below_the_floor_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.Argon2Iterations = 1;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("Argon2Iterations", StringComparison.Ordinal));
        }

        [Fact]
        public void A_parallelism_below_one_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.Argon2Parallelism = 0;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("Argon2Parallelism", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // Token lifetimes.
        // -----------------------------------------------------------------------------------

        // There is no AccessTokenMinutes any more - it was validated and read by nothing (security
        // review, 2026-09-25). The one token is the session token, and this is its lifetime.
        [Fact]
        public void A_zero_refresh_token_lifetime_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.RefreshTokenDays = 0;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("RefreshTokenDays", StringComparison.Ordinal));
        }

        // -----------------------------------------------------------------------------------
        // The disk reserve (security review, 2026-09-25).
        // -----------------------------------------------------------------------------------

        // A negative reserve reads as "always room" - the guard silently off while the setting looks
        // configured. Refused rather than treated as zero.
        [Fact]
        public void A_negative_disk_reserve_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.MinimumFreeDiskBytes = -1;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("MinimumFreeDiskBytes", StringComparison.Ordinal));
        }

        // The feedback folder's total (code review, 2026-10-04): negative would read as "always
        // full", refusing every attachment while the setting looked configured.
        [Fact]
        public void A_negative_feedback_total_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.FeedbackMaxStoredBytes = -1;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("FeedbackMaxStoredBytes", StringComparison.Ordinal));
        }

        // Zero is the deliberate way to switch the reserve off, and must not be mistaken for a fault.
        [Fact]
        public void A_zero_disk_reserve_is_allowed()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.MinimumFreeDiskBytes = 0;

            Assert.DoesNotContain(
                ServerOptionsValidatorTests.Validate(options),
                f => f.Contains("MinimumFreeDiskBytes", StringComparison.Ordinal));
        }

        // The reserve has a DEFAULT, unlike the data paths: forgetting it must leave the protection
        // on, not off.
        [Fact]
        public void The_disk_reserve_is_on_by_default()
        {
            Assert.True(new ServerOptions().MinimumFreeDiskBytes > 0);
        }

        // -----------------------------------------------------------------------------------
        // Reporting behaviour.
        // -----------------------------------------------------------------------------------

        // The project owner editing a config file over SSH should see every problem at once. Fixing one,
        // restarting, and discovering the next is a slow loop, so validation must not stop at the
        // first failure.
        // -----------------------------------------------------------------------------------
        // Publishing to production (2026-09-25): all three settings or none.
        // -----------------------------------------------------------------------------------

        private static ServerOptions WithProduction(ServerOptions options)
        {
            string production = ServerOptionsValidatorTests.ProductionPath();

            options.ProductionDataTreeRoot = production;
            options.ProductionManifestPath = Path.Combine(Path.GetDirectoryName(production)!, "dataChecksums.json");
            options.ProductionPublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data/Data";

            return options;
        }

        private static ServerOptions ValidWithProduction()
        {
            string production = ServerOptionsValidatorTests.ProductionPath();

            // ProductionTreeRoot is the tree ("app-data"); the data root sits inside it.
            return ServerOptionsValidatorTests.WithProduction(ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), Path.GetDirectoryName(production)!));
        }

        [Fact]
        public void With_NO_production_settings_the_service_starts_and_production_publishing_is_off()
        {
            // A server configured before this existed keeps starting, and stays unable to write
            // production.
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());

            Assert.Empty(ServerOptionsValidatorTests.Validate(options));
            Assert.False(options.IsProductionPublishingConfigured);
        }

        [Fact]
        public void With_ALL_THREE_production_settings_the_service_starts_and_production_publishing_is_on()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidWithProduction();

            Assert.Empty(ServerOptionsValidatorTests.Validate(options));
            Assert.True(options.IsProductionPublishingConfigured);
        }

        [Fact]
        public void Only_SOME_of_the_production_settings_refuses_to_start_and_names_the_missing_ones()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidWithProduction();
            options.ProductionManifestPath = null;

            string failure = Assert.Single(ServerOptionsValidatorTests.Validate(options));
            Assert.Contains("ProductionManifestPath", failure, StringComparison.Ordinal);
        }

        [Fact]
        public void A_production_setting_carrying_the_BETA_marker_refuses_to_start()
        {
            // The mistake this catches: pasting the BETA manifest path or URL into the production
            // setting, which would advertise BETA data to every user.
            ServerOptions options = ServerOptionsValidatorTests.ValidWithProduction();
            options.ProductionPublicDataBaseUrl = "https://classic-repair-toolbox.dk/app-data-BETA/Data";

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                failure => failure.Contains("ProductionPublicDataBaseUrl", StringComparison.Ordinal) &&
                           failure.Contains("-BETA", StringComparison.Ordinal));
        }

        [Fact]
        public void A_production_data_root_that_OVERLAPS_the_BETA_tree_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidWithProduction();
            options.RequiredTreeMarker = string.Empty;
            options.ProductionDataTreeRoot = options.DataTreeRoot;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                failure => failure.Contains("overlaps DataTreeRoot", StringComparison.Ordinal));
        }

        [Fact]
        public void A_production_data_root_OUTSIDE_ProductionTreeRoot_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidWithProduction();
            options.ProductionDataTreeRoot = Path.Combine(Path.GetTempPath(), "crt-test", "elsewhere", "Data");

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                failure => failure.Contains("is not inside", StringComparison.Ordinal));
        }

        [Fact]
        public void A_production_data_root_the_service_cannot_WRITE_refuses_to_start_rather_than_failing_half_way()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidWithProduction();
            string production = ServerOptionsValidatorTests.ProductionPath();

            IReadOnlyList<string> failures = ServerOptionsValidator.Validate(
                options, _ => true, path => !path.StartsWith(production, StringComparison.Ordinal));

            Assert.Contains(failures, failure => failure.Contains("ProductionDataTreeRoot is not writable", StringComparison.Ordinal));
        }

        [Fact]
        public void The_same_manifest_file_for_both_trees_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidWithProduction();
            options.RequiredTreeMarker = string.Empty;
            options.ProductionManifestPath = options.ManifestPath;

            Assert.Contains(
                ServerOptionsValidatorTests.Validate(options),
                failure => failure.Contains("same file as ManifestPath", StringComparison.Ordinal));
        }

        [Fact]
        public void Every_problem_is_reported_at_once_rather_than_only_the_first()
        {
            var options = new ServerOptions
            {
                DataTreeRoot = null,
                ProductionTreeRoot = null,
                ConnectionString = null,
                PublicApiBaseUrl = null,
                PublicDataBaseUrl = null,
                ManifestPath = null
            };

            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(options);

            Assert.True(
                failures.Count >= 6,
                $"Expected every missing setting to be reported, but got {failures.Count}: " +
                string.Join(" | ", failures));
        }

        [Fact]
        public void Validation_refuses_null_arguments()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());

            Assert.Throws<ArgumentNullException>(() =>
                ServerOptionsValidator.Validate(null!, _ => true, _ => true));
            Assert.Throws<ArgumentNullException>(() =>
                ServerOptionsValidator.Validate(options, null!, _ => true));
            Assert.Throws<ArgumentNullException>(() =>
                ServerOptionsValidator.Validate(options, _ => true, null!));
        }

        // -----------------------------------------------------------------------------------
        // Feedback (2026-10-03): the folder its files are saved in, and who reads it. No default
        // for either, and the folder kept apart from everything published or submitted.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public void A_missing_feedback_folder_or_address_refuses_to_start(string? value)
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.FeedbackRoot = value;
            options.FeedbackToAddress = value;

            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(options);

            Assert.Contains(failures, failure => failure.Contains("FeedbackRoot is not set", StringComparison.Ordinal));
            Assert.Contains(failures, failure => failure.Contains("FeedbackToAddress is not set", StringComparison.Ordinal));
        }

        [Fact]
        public void A_relative_feedback_folder_or_an_address_without_an_at_refuses_to_start()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());
            options.FeedbackRoot = "user-feedback";
            options.FeedbackToAddress = "dennis";

            IReadOnlyList<string> failures = ServerOptionsValidatorTests.Validate(options);

            Assert.Contains(failures, failure => failure.Contains("FeedbackRoot must be an absolute path", StringComparison.Ordinal));
            Assert.Contains(failures, failure => failure.Contains("FeedbackToAddress is not an email address", StringComparison.Ordinal));
        }

        // A stranger's attached files must never be published, nor sit among the blobs.
        [Fact]
        public void A_feedback_folder_inside_the_BETA_tree_or_holding_the_blob_store_refuses_to_start()
        {
            string beta = ServerOptionsValidatorTests.BetaPath();

            ServerOptions inBeta = ServerOptionsValidatorTests.ValidOptions(beta, ServerOptionsValidatorTests.ProductionPath());
            inBeta.FeedbackRoot = Path.Combine(beta, "user-feedback");

            ServerOptions aboveBlobs = ServerOptionsValidatorTests.ValidOptions(beta, ServerOptionsValidatorTests.ProductionPath());
            aboveBlobs.FeedbackRoot = Path.GetTempPath();

            Assert.Contains(ServerOptionsValidatorTests.Validate(inBeta), failure => failure.Contains("FeedbackRoot", StringComparison.Ordinal) && failure.Contains("overlaps DataTreeRoot", StringComparison.Ordinal));
            Assert.Contains(ServerOptionsValidatorTests.Validate(aboveBlobs), failure => failure.Contains("overlaps BlobStoreRoot", StringComparison.Ordinal));
        }

        [Fact]
        public void A_feedback_folder_the_service_cannot_write_refuses_to_start_and_names_the_fix()
        {
            ServerOptions options = ServerOptionsValidatorTests.ValidOptions(
                ServerOptionsValidatorTests.BetaPath(), ServerOptionsValidatorTests.ProductionPath());

            IReadOnlyList<string> failures = ServerOptionsValidator.Validate(
                options,
                _ => true,
                path => !path.EndsWith("user-feedback", StringComparison.Ordinal));

            string failure = Assert.Single(failures);
            Assert.Contains("FeedbackRoot is not writable", failure, StringComparison.Ordinal);
            Assert.Contains("Feedback from CRT", failure, StringComparison.Ordinal);
        }
    }
}
