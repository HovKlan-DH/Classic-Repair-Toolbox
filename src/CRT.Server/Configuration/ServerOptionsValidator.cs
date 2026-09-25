using System.Globalization;

namespace CRT.Server.Configuration
{
    // ###########################################################################################
    // Decides whether the service may start. PURE: it takes an options object and a pair of
    // predicates for the two questions only the filesystem can answer, and returns a list of
    // human-readable failures. It opens nothing, writes nothing and logs nothing, so every branch
    // below is a unit test with no server, no disk and no database.
    //
    // WHY A LIST RATHER THAN THE FIRST FAILURE. A maintainer editing a config file by hand on a
    // server, through the deployment runbook, wants every problem at once - fixing one, restarting,
    // and discovering the next is a slow loop over an SSH session. Every check therefore runs and
    // the messages accumulate.
    //
    // WHY THE MESSAGES NAME THE SETTING. These strings are what appears in "systemctl status" and
    // the journal when the unit fails to start. DEPLOYMENT.md tells the maintainer that a refusal
    // names the setting, so each message must carry the key as written in the JSON.
    //
    // WHAT THIS IS NOT. Passing here does NOT make writing to Production impossible - only the
    // filesystem permissions in DEPLOYMENT.md step 3 and systemd's ProtectSystem=strict do that.
    // These checks catch honest misconfiguration early and loudly; they are not the control.
    // ###########################################################################################
    public static class ServerOptionsValidator
    {
        // ###########################################################################################
        // Validates the options. The two delegates exist so the pure logic can be tested without
        // touching a disk: production passes real filesystem probes, tests pass canned answers.
        //
        // directoryExists:  does this absolute path name an existing directory?
        // directoryWritable: can this process actually create a file in it? Asking the permission
        //                    bits is not enough - group membership, ACLs, a read-only mount and
        //                    systemd's own ProtectSystem can each make a "writable-looking"
        //                    directory refuse a write. The only honest test is to try.
        // ###########################################################################################
        public static IReadOnlyList<string> Validate(
            ServerOptions options,
            Func<string, bool> directoryExists,
            Func<string, bool> directoryWritable)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(directoryExists);
            ArgumentNullException.ThrowIfNull(directoryWritable);

            var failures = new List<string>();

            string? dataTreeRoot = ServerOptionsValidator.ValidateDataTree(
                options, directoryExists, directoryWritable, failures);

            ServerOptionsValidator.ValidateProductionSeparation(options, dataTreeRoot, failures);
            ServerOptionsValidator.ValidateProductionPublishing(options, dataTreeRoot, directoryExists, directoryWritable, failures);
            ServerOptionsValidator.ValidateManifestSettings(options, failures);
            ServerOptionsValidator.ValidateRequiredValues(options, failures);
            ServerOptionsValidator.ValidateHashingParameters(options, failures);
            ServerOptionsValidator.ValidateTokenLifetimes(options, failures);

            return failures;
        }

        // ###########################################################################################
        // The data tree itself: present, absolute, real, writable, and carrying the safety marker.
        // Returns the normalised path, or null when it could not be established - callers use that
        // to skip comparisons that would be meaningless.
        // ###########################################################################################
        private static string? ValidateDataTree(
            ServerOptions options,
            Func<string, bool> directoryExists,
            Func<string, bool> directoryWritable,
            List<string> failures)
        {
            if (string.IsNullOrWhiteSpace(options.DataTreeRoot))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:DataTreeRoot is not set. It has no default on " +
                    "purpose - the service will not guess where to write data. Set it to the BETA " +
                    "data tree, e.g. /.../app-data-BETA/Data");
                return null;
            }

            // The rootedness check must be made against the RAW value, before normalisation:
            // Path.GetFullPath resolves a relative path against the current working directory and
            // hands back an absolute one, so testing the normalised result would never fail and
            // the check would be dead code. Caught by its own test, which passed validation on a
            // deliberately relative path.
            if (!Path.IsPathRooted(options.DataTreeRoot.Trim()))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:DataTreeRoot must be an absolute path, but is " +
                    $"[{options.DataTreeRoot}]. A relative path would resolve against the service's " +
                    "own working directory, which is not where the data lives.");
                return null;
            }

            if (!ServerOptionsValidator.TryNormalisePath(options.DataTreeRoot, out string normalised))
            {
                failures.Add($"{ServerOptions.SectionName}:DataTreeRoot is not a usable path: [{options.DataTreeRoot}]");
                return null;
            }

            if (!ServerOptionsValidator.HasRequiredMarker(normalised, options.RequiredTreeMarker))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:DataTreeRoot does not contain the required marker " +
                    $"[{options.RequiredTreeMarker}]: [{options.DataTreeRoot}]. This is the guard " +
                    "against pointing the service at the production tree. If you genuinely intend " +
                    $"to target a tree without that marker, clear {ServerOptions.SectionName}:RequiredTreeMarker " +
                    "deliberately.");
            }

            if (!directoryExists(normalised))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:DataTreeRoot does not exist or is not a directory: " +
                    $"[{normalised}]");
                return normalised;
            }

            if (!directoryWritable(normalised))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:DataTreeRoot is not writable by the service user: " +
                    $"[{normalised}]. Check the group ownership and the setgid bit - see " +
                    "DEPLOYMENT.md step 3.");
            }

            return normalised;
        }

        // ###########################################################################################
        // The Production refusal. Rejects a data tree that IS production, sits inside it, or
        // contains it - all three of which would let a write reach the tree every user syncs from.
        // ###########################################################################################
        private static void ValidateProductionSeparation(
            ServerOptions options,
            string? dataTreeRoot,
            List<string> failures)
        {
            if (string.IsNullOrWhiteSpace(options.ProductionTreeRoot))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:ProductionTreeRoot is not set. It names the " +
                    "Production tree (the folder holding Production's Data folder and its " +
                    "dataChecksums.json) so BETA publishing can be kept out of it, and it is " +
                    "required even while publishing to production is off. It is NOT one of the " +
                    "three settings that switch that on - ProductionDataTreeRoot, " +
                    "ProductionManifestPath and ProductionPublicDataBaseUrl may stay empty; this " +
                    "one may not.");
                return;
            }

            if (dataTreeRoot is null)
                return;

            if (!ServerOptionsValidator.TryNormalisePath(options.ProductionTreeRoot, out string production))
            {
                failures.Add($"{ServerOptions.SectionName}:ProductionTreeRoot is not a usable path: [{options.ProductionTreeRoot}]");
                return;
            }

            if (ServerOptionsValidator.PathsAreEqual(dataTreeRoot, production))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:DataTreeRoot and ProductionTreeRoot are the same " +
                    $"directory: [{dataTreeRoot}]. The service must never write the production tree.");
                return;
            }

            if (ServerOptionsValidator.PathContains(production, dataTreeRoot))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:DataTreeRoot [{dataTreeRoot}] is inside " +
                    $"ProductionTreeRoot [{production}]. Writing there would reach the tree every " +
                    "user syncs from.");
            }
            else if (ServerOptionsValidator.PathContains(dataTreeRoot, production))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:ProductionTreeRoot [{production}] is inside " +
                    $"DataTreeRoot [{dataTreeRoot}]. Writing anywhere under the data root could " +
                    "then reach production.");
            }
        }

        // ###########################################################################################
        // Publishing to PRODUCTION (maintainer request, 2026-09-25): all three settings or none.
        //
        // None is the feature switched off, and valid. Some-but-not-all is refused, because a
        // half-configured promotion either cannot run or - worse - writes one tree and advertises
        // it in another's manifest.
        //
        // When configured, each Production value must be the PRODUCTION twin: no BETA marker in
        // it, not equal to its BETA counterpart, and the data root inside ProductionTreeRoot and
        // clear of DataTreeRoot in both directions. The data root must also be writable, checked
        // here so a missing permission is a refusal to start rather than a promotion that stops
        // half way through a board.
        // ###########################################################################################
        private static void ValidateProductionPublishing(
            ServerOptions options,
            string? dataTreeRoot,
            Func<string, bool> directoryExists,
            Func<string, bool> directoryWritable,
            List<string> failures)
        {
            string prefix = ServerOptions.SectionName;

            var set = new (string Name, string? Value)[]
            {
                ("ProductionDataTreeRoot", options.ProductionDataTreeRoot),
                ("ProductionManifestPath", options.ProductionManifestPath),
                ("ProductionPublicDataBaseUrl", options.ProductionPublicDataBaseUrl)
            };

            int configured = set.Count(setting => !string.IsNullOrWhiteSpace(setting.Value));

            if (configured == 0)
                return;

            if (configured < set.Length)
            {
                string missing = string.Join(", ", set
                    .Where(setting => string.IsNullOrWhiteSpace(setting.Value))
                    .Select(setting => $"{prefix}:{setting.Name}"));

                failures.Add(
                    $"Publishing to production needs all three Production settings, and {missing} " +
                    "is not set. Set all three to switch it on, or none to leave it off.");
                return;
            }

            string marker = options.RequiredTreeMarker;
            bool hasMarker = !string.IsNullOrEmpty(marker);

            // ---- The data root ---------------------------------------------------------------
            string rawRoot = options.ProductionDataTreeRoot!.Trim();

            if (!Path.IsPathRooted(rawRoot))
            {
                failures.Add($"{prefix}:ProductionDataTreeRoot must be an absolute path, but is [{options.ProductionDataTreeRoot}]");
            }
            else if (!ServerOptionsValidator.TryNormalisePath(rawRoot, out string productionData))
            {
                failures.Add($"{prefix}:ProductionDataTreeRoot is not a usable path: [{options.ProductionDataTreeRoot}]");
            }
            else
            {
                if (hasMarker && ServerOptionsValidator.HasRequiredMarker(productionData, marker))
                {
                    failures.Add(
                        $"{prefix}:ProductionDataTreeRoot contains the BETA marker [{marker}]: " +
                        $"[{options.ProductionDataTreeRoot}]. It must name the PRODUCTION data tree.");
                }

                if (dataTreeRoot is not null &&
                    (ServerOptionsValidator.PathsAreEqual(dataTreeRoot, productionData) ||
                     ServerOptionsValidator.PathContains(dataTreeRoot, productionData) ||
                     ServerOptionsValidator.PathContains(productionData, dataTreeRoot)))
                {
                    failures.Add(
                        $"{prefix}:ProductionDataTreeRoot [{productionData}] overlaps DataTreeRoot " +
                        $"[{dataTreeRoot}]. BETA and production must be two separate trees.");
                }

                if (ServerOptionsValidator.TryNormalisePath(options.ProductionTreeRoot ?? string.Empty, out string productionTree) &&
                    !ServerOptionsValidator.PathsAreEqual(productionTree, productionData) &&
                    !ServerOptionsValidator.PathContains(productionTree, productionData))
                {
                    failures.Add(
                        $"{prefix}:ProductionDataTreeRoot [{productionData}] is not inside " +
                        $"ProductionTreeRoot [{productionTree}]. The two must name the same tree.");
                }

                if (!directoryExists(productionData))
                {
                    failures.Add($"{prefix}:ProductionDataTreeRoot does not exist or is not a directory: [{productionData}]");
                }
                else if (!directoryWritable(productionData))
                {
                    failures.Add(
                        $"{prefix}:ProductionDataTreeRoot is not writable by the service user: " +
                        $"[{productionData}]. Publishing to production needs the permissions in " +
                        "DEPLOYMENT.md step 3 and the tree in the unit's ReadWritePaths.");
                }
            }

            // ---- The manifest ----------------------------------------------------------------
            string rawManifest = options.ProductionManifestPath!.Trim();

            if (!Path.IsPathRooted(rawManifest))
            {
                failures.Add($"{prefix}:ProductionManifestPath must be an absolute path, but is [{options.ProductionManifestPath}]");
            }
            else if (ServerOptionsValidator.TryNormalisePath(rawManifest, out string productionManifest))
            {
                if (hasMarker && ServerOptionsValidator.HasRequiredMarker(productionManifest, marker))
                {
                    failures.Add(
                        $"{prefix}:ProductionManifestPath contains the BETA marker [{marker}]: " +
                        $"[{options.ProductionManifestPath}]");
                }

                if (ServerOptionsValidator.TryNormalisePath(options.ManifestPath ?? string.Empty, out string betaManifest) &&
                    ServerOptionsValidator.PathsAreEqual(betaManifest, productionManifest))
                {
                    failures.Add(
                        $"{prefix}:ProductionManifestPath is the same file as ManifestPath: " +
                        $"[{productionManifest}]. Each tree has its own dataChecksums.json.");
                }
            }
            else
            {
                failures.Add($"{prefix}:ProductionManifestPath is not a usable path: [{options.ProductionManifestPath}]");
            }

            // ---- The public URL --------------------------------------------------------------
            string url = options.ProductionPublicDataBaseUrl!.Trim();

            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri? productionUri) ||
                !string.Equals(productionUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                failures.Add(
                    $"{prefix}:ProductionPublicDataBaseUrl must be an absolute https:// URL, but is " +
                    $"[{options.ProductionPublicDataBaseUrl}].");
            }
            else
            {
                if (hasMarker && ServerOptionsValidator.HasRequiredMarker(url, marker))
                {
                    failures.Add(
                        $"{prefix}:ProductionPublicDataBaseUrl contains the BETA marker [{marker}]: " +
                        $"[{options.ProductionPublicDataBaseUrl}]. That would advertise BETA data to every user.");
                }

                if (string.Equals(
                        url.TrimEnd('/'),
                        (options.PublicDataBaseUrl ?? string.Empty).Trim().TrimEnd('/'),
                        StringComparison.OrdinalIgnoreCase))
                {
                    failures.Add(
                        $"{prefix}:ProductionPublicDataBaseUrl is the same as PublicDataBaseUrl: [{url}].");
                }
            }
        }

        // ###########################################################################################
        // The manifest settings. Nothing reads these until Phase 5, but a wrong pairing here is
        // invisible in every way that matters - the manifest is well-formed, correctly hashed and
        // correctly sorted, and simply points every client at the wrong tree. Validating now costs
        // nothing and cannot be forgotten later.
        // ###########################################################################################
        private static void ValidateManifestSettings(ServerOptions options, List<string> failures)
        {
            if (string.IsNullOrWhiteSpace(options.PublicDataBaseUrl))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:PublicDataBaseUrl is not set. It has no default: " +
                    "it becomes the url of every entry in dataChecksums.json, and pointing it at " +
                    "the wrong tree produces a perfectly valid manifest that sends every client to " +
                    "the wrong data.");
            }
            else if (!Uri.TryCreate(options.PublicDataBaseUrl, UriKind.Absolute, out Uri? dataUri) ||
                     !string.Equals(dataUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:PublicDataBaseUrl must be an absolute https:// URL, " +
                    $"but is [{options.PublicDataBaseUrl}]. The application refuses a manifest whose " +
                    "entries are not https.");
            }
            else if (!ServerOptionsValidator.HasRequiredMarker(options.PublicDataBaseUrl, options.RequiredTreeMarker))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:PublicDataBaseUrl does not contain the required " +
                    $"marker [{options.RequiredTreeMarker}]: [{options.PublicDataBaseUrl}]. That is " +
                    "how a BETA data tree ends up publishing production URLs.");
            }

            if (string.IsNullOrWhiteSpace(options.ManifestPath))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:ManifestPath is not set. It has no default: on the " +
                    "server the manifest sits BESIDE the data root rather than inside it, so it " +
                    "cannot be derived from DataTreeRoot.");
            }
            // Rootedness against the RAW value, for the same reason as DataTreeRoot above:
            // GetFullPath would silently make a relative path absolute.
            else if (!Path.IsPathRooted(options.ManifestPath.Trim()))
            {
                failures.Add($"{ServerOptions.SectionName}:ManifestPath must be an absolute path, but is [{options.ManifestPath}]");
            }
            else if (!ServerOptionsValidator.TryNormalisePath(options.ManifestPath, out string manifest))
            {
                failures.Add($"{ServerOptions.SectionName}:ManifestPath is not a usable path: [{options.ManifestPath}]");
            }
            else if (!ServerOptionsValidator.HasRequiredMarker(manifest, options.RequiredTreeMarker))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:ManifestPath does not contain the required marker " +
                    $"[{options.RequiredTreeMarker}]: [{options.ManifestPath}]");
            }
        }

        // ###########################################################################################
        // The remaining values that have no safe default.
        // ###########################################################################################
        private static void ValidateRequiredValues(ServerOptions options, List<string> failures)
        {
            if (string.IsNullOrWhiteSpace(options.ConnectionString))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:ConnectionString is not set. It has no default - " +
                    "a guessed database is worse than none.");
            }

            if (string.IsNullOrWhiteSpace(options.PublicApiBaseUrl))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:PublicApiBaseUrl is not set. It has no default: it " +
                    "builds the links in verification and password-reset mail, and a wrong value " +
                    "emails people a link that does not work.");
            }
            else if (!Uri.TryCreate(options.PublicApiBaseUrl, UriKind.Absolute, out Uri? apiUri) ||
                     !string.Equals(apiUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:PublicApiBaseUrl must be an absolute https:// URL, " +
                    $"but is [{options.PublicApiBaseUrl}]. A password-reset link sent over http is " +
                    "an account takeover waiting for someone on the same network.");
            }

            // The blob store, from the submissions step onward. No default, and required, because
            // this directory accumulates contributor-supplied bytes: a guessed location could fill
            // a partition nobody is watching, and one inside the document root would make the
            // blobs of a submission still under review downloadable by anyone who guessed a hash.
            if (string.IsNullOrWhiteSpace(options.BlobStoreRoot))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:BlobStoreRoot is not set. It has no default: it is " +
                    "where submitted files are stored, and it must sit outside the web document root.");
            }
            else if (!Path.IsPathRooted(options.BlobStoreRoot))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:BlobStoreRoot must be an absolute path, but is " +
                    $"[{options.BlobStoreRoot}]. A relative path would resolve against the service's " +
                    "own working directory, which is not where submissions belong.");
            }

            // Required rather than merely validated-if-present, from the accounts step onward:
            // every mail this service sends needs a From address, so a missing one would fail at
            // the moment someone registers - a runtime NullReferenceException in a background
            // send - rather than at startup. Same bargain as every other no-default setting: an
            // outage that is noticed beats a silent failure that is not.
            if (string.IsNullOrWhiteSpace(options.MailFromAddress))
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:MailFromAddress is not set. It has no default: it " +
                    "is the From address on every verification and password-reset mail, and mail " +
                    "sent without one is rejected by the receiving server.");
            }
            else if (!options.MailFromAddress.Contains('@', StringComparison.Ordinal))
            {
                failures.Add($"{ServerOptions.SectionName}:MailFromAddress is not an email address: [{options.MailFromAddress}]");
            }

            if (options.SmtpPort is < 1 or > 65535)
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:SmtpPort must be between 1 and 65535, but is " +
                    $"{options.SmtpPort.ToString(CultureInfo.InvariantCulture)}");
            }
        }

        // ###########################################################################################
        // Argon2id cost floor. This is where the production parameters are ASSERTED without ever
        // being executed: the hasher's own tests run at deliberately tiny parameters so the suite
        // stays fast, which would otherwise leave the real cost pinned by nothing at all.
        // ###########################################################################################
        private static void ValidateHashingParameters(ServerOptions options, List<string> failures)
        {
            const int minimumMemoryKib = 19 * 1024;   // RFC 9106's low-memory profile floor.
            const int minimumIterations = 2;

            if (options.Argon2MemoryKib < minimumMemoryKib)
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:Argon2MemoryKib is {options.Argon2MemoryKib.ToString(CultureInfo.InvariantCulture)}, " +
                    $"below the floor of {minimumMemoryKib.ToString(CultureInfo.InvariantCulture)} KiB. Memory is what makes " +
                    "an offline attack on a stolen hash expensive; lowering it is the one parameter " +
                    "that quietly undoes password hashing.");
            }

            if (options.Argon2Iterations < minimumIterations)
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:Argon2Iterations is {options.Argon2Iterations.ToString(CultureInfo.InvariantCulture)}, " +
                    $"below the floor of {minimumIterations.ToString(CultureInfo.InvariantCulture)}.");
            }

            if (options.Argon2Parallelism < 1)
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:Argon2Parallelism must be at least 1, but is " +
                    $"{options.Argon2Parallelism.ToString(CultureInfo.InvariantCulture)}");
            }
        }

        private static void ValidateTokenLifetimes(ServerOptions options, List<string> failures)
        {
            if (options.RefreshTokenDays < 1)
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:RefreshTokenDays must be at least 1, but is " +
                    $"{options.RefreshTokenDays.ToString(CultureInfo.InvariantCulture)}");
            }

            // A negative reserve would read as "always room", silently switching the disk guard off
            // while the setting looked configured. Zero is the deliberate way to turn it off.
            if (options.MinimumFreeDiskBytes < 0)
            {
                failures.Add(
                    $"{ServerOptions.SectionName}:MinimumFreeDiskBytes must be 0 or more, but is " +
                    $"{options.MinimumFreeDiskBytes.ToString(CultureInfo.InvariantCulture)}");
            }
        }

        // ###########################################################################################
        // Path helpers.
        //
        // Comparison is ORDINAL and case-SENSITIVE, because the server runs on Linux where two
        // paths differing only in case are two different directories. Comparing case-insensitively
        // here would let "/srv/App-Data" and "/srv/app-data" look like the same tree when they are
        // not - and the whole point of these checks is to tell two trees apart.
        // ###########################################################################################
        internal static bool TryNormalisePath(string? value, out string normalised)
        {
            normalised = string.Empty;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            try
            {
                // GetFullPath resolves "..", duplicate separators and a trailing separator, so two
                // spellings of one directory compare equal. It does NOT resolve symlinks - a
                // symlinked path could still point somewhere unexpected, which is exactly why the
                // filesystem permissions rather than this check are the real control.
                normalised = Path.GetFullPath(value.Trim()).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

                // Preserve the root itself ("/" trims to empty above).
                if (normalised.Length == 0)
                    normalised = Path.GetFullPath(value.Trim());

                return true;
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
            {
                return false;
            }
        }

        internal static bool PathsAreEqual(string left, string right) =>
            string.Equals(left, right, StringComparison.Ordinal);

        // ###########################################################################################
        // True when "candidate" sits inside "container". The separator is appended before comparing
        // so that "/srv/app-data-BETA" is not treated as living inside "/srv/app-data" - a plain
        // StartsWith would say it does, and would then refuse a perfectly good configuration (or,
        // with the arguments the other way round, accept a dangerous one).
        // ###########################################################################################
        internal static bool PathContains(string container, string candidate)
        {
            string prefix = container.EndsWith(Path.DirectorySeparatorChar)
                ? container
                : container + Path.DirectorySeparatorChar;

            return candidate.StartsWith(prefix, StringComparison.Ordinal);
        }

        private static bool HasRequiredMarker(string value, string? marker)
        {
            // An empty marker means the maintainer deliberately turned the check off.
            if (string.IsNullOrEmpty(marker))
                return true;

            return value.Contains(marker, StringComparison.Ordinal);
        }
    }
}
