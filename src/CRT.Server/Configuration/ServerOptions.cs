namespace CRT.Server.Configuration
{
    // ###########################################################################################
    // Everything the service reads from appsettings.Production.json, which sits in the service's
    // own directory beside the binaries. See DEPLOYMENT.md for the deployed layout.
    //
    // THE DEFAULTS ARE THE POINT. Every value that decides WHERE DATA GOES is null here, with no
    // fallback anywhere, so a service that was not told explicitly refuses to start rather than
    // quietly choosing somewhere. That is the strategy document's own worked example of secure by
    // design: "the production path has no default, so a misconfigured service refuses to start".
    // The failure mode of forgetting is an outage - noticed at once, harms nobody - instead of an
    // exposure.
    //
    // Values that merely TUNE behaviour (token lifetimes, hashing cost) do carry defaults, because
    // a missing one there has a safe answer and refusing to start would be obstruction rather than
    // protection. The line between the two groups is "could getting this wrong write data
    // somewhere it should not, or weaken a control?" - if yes, no default.
    //
    // Nothing in this class may be logged wholesale: ConnectionString is a credential. Log
    // individual non-secret fields where useful, never the object.
    // ###########################################################################################
    public sealed class ServerOptions
    {
        // The configuration section name in appsettings.json.
        public const string SectionName = "CrtServer";

        // -----------------------------------------------------------------------------------
        // Data locations. All three are deliberately null by default - see the header.
        // -----------------------------------------------------------------------------------

        // The BETA data tree the service may write. Phase 3 only validates it; Phase 4 onward
        // writes through it.
        public string? DataTreeRoot { get; set; }

        // The Production tree, named here ONLY so it can be refused. The service never writes it:
        // promotion from BETA to Production is a manual copy the maintainer performs. Naming it
        // lets the validator reject a DataTreeRoot that equals, contains, or sits inside it.
        //
        // This check is the in-code half of the interlock and is NOT what makes it safe. The real
        // guarantee is that the service user holds no write permission on that tree, so the kernel
        // refuses the write regardless of what this program believes - plus systemd's
        // ProtectSystem=strict. See DEPLOYMENT.md step 3.
        public string? ProductionTreeRoot { get; set; }

        // Where the published data is served from, e.g. https://classic-repair-toolbox.dk/app-data-BETA/Data.
        // Unread until Phase 5, which regenerates dataChecksums.json - but validated from Phase 3,
        // because the failure it prevents is invisible: the data root, the manifest path and this
        // URL are three independent values, so pairing the BETA tree with the Production base URL
        // produces a perfectly valid, correctly sorted manifest in which every url points at the
        // wrong tree. Clients would then sync Production content believing it is BETA, or 404 on
        // everything. Cheap to require now; impossible to notice later.
        public string? PublicDataBaseUrl { get; set; }

        // Where dataChecksums.json is written. On BETA it sits BESIDE the Data/ folder, not inside
        // it, so this is genuinely independent of DataTreeRoot and must be configured separately.
        // Unread until Phase 5, validated from Phase 3, same reasoning as above.
        public string? ManifestPath { get; set; }

        // Where submitted blobs are stored while a submission is in flight and after it is queued
        // (Phase 4). Content-addressed - see BlobStorePaths.
        //
        // NO DEFAULT, for the same reason as the data tree: this directory accumulates
        // contributor-supplied bytes, and a service that guessed at its location could fill a
        // partition nobody was watching. It must also sit OUTSIDE the document root, or the blobs
        // of a submission still under review would be downloadable by anyone who guessed a hash.
        public string? BlobStoreRoot { get; set; }

        // -----------------------------------------------------------------------------------
        // Safety marker.
        // -----------------------------------------------------------------------------------

        // Every configured data path and the public base URL must contain this marker. It is the
        // belt-and-braces half of the Production refusal: even if DataTreeRoot and
        // ProductionTreeRoot were somehow both wrong, a path with no "-BETA" in it does not start
        // the service.
        //
        // The maintainer sets this to empty ONLY when deliberately pointing the service at
        // Production, which is a decision the strategy document says must be explicit. Defaulted
        // rather than null because the safe value is knowable.
        public string RequiredTreeMarker { get; set; } = "-BETA";

        // -----------------------------------------------------------------------------------
        // Database. No default: a connection string cannot be guessed, and a service that
        // silently pointed at the wrong database would be worse than one that will not start.
        // -----------------------------------------------------------------------------------
        public string? ConnectionString { get; set; }

        // -----------------------------------------------------------------------------------
        // Public URL of the API itself, used to build links in verification and password-reset
        // mail. No default: a wrong value here emails people a dead or hostile link.
        // -----------------------------------------------------------------------------------
        public string? PublicApiBaseUrl { get; set; }

        // -----------------------------------------------------------------------------------
        // Mail. Defaults match the deployment documented in DEPLOYMENT.md (local postfix), and
        // are safe: the worst case of a wrong port here is mail that fails to send, which is
        // visible and harms nothing.
        // -----------------------------------------------------------------------------------
        public string SmtpHost { get; set; } = "127.0.0.1";
        public int SmtpPort { get; set; } = 25;
        public string? MailFromAddress { get; set; }
        public string MailFromDisplayName { get; set; } = "Classic Repair Toolbox";

        // -----------------------------------------------------------------------------------
        // Token lifetimes. Defaults are sane, so these are tuning rather than safety.
        //
        // Access is short because it is presented on every request; refresh is long because the
        // audience opens CRT every few weeks and forcing a monthly password re-entry trains people
        // into weaker passwords. Rotation with reuse detection is what keeps the long refresh
        // lifetime acceptable - see the sessions design.
        // -----------------------------------------------------------------------------------
        public int AccessTokenMinutes { get; set; } = 30;
        public int RefreshTokenDays { get; set; } = 30;

        // -----------------------------------------------------------------------------------
        // Argon2id parameters. Defaults are RFC 9106's second recommended profile, chosen for a
        // small box: 64 MiB at parallelism 2 is about 128 MiB of transient allocation per
        // concurrent login, which is why the rate limiter must sit IN FRONT of the hasher rather
        // than behind it - otherwise unlimited login attempts are a memory-exhaustion vector.
        //
        // These may be raised without invalidating existing passwords: the encoded hash stores the
        // parameters it was made with, so verification uses those and the login path rehashes when
        // they are below the current configuration.
        // -----------------------------------------------------------------------------------
        public int Argon2MemoryKib { get; set; } = 65536;
        public int Argon2Iterations { get; set; } = 3;
        public int Argon2Parallelism { get; set; } = 2;
    }
}
