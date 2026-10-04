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

        // The Production tree, named so the BETA settings can be kept AWAY from it: the validator
        // rejects a DataTreeRoot that equals, contains, or sits inside it, so no publish - which
        // writes BETA - can ever land in Production.
        //
        // *** PRODUCTION IS WRITTEN SINCE 2026-09-25, BUT ONLY BY THE PROMOTION. *** The
        // project owner asked for a two-stage publish - BETA first, Production after a maintainer has
        // checked it there - so the service may now write ProductionDataTreeRoot below, and only
        // through ProductionPromoter, which copies bytes that are already in BETA and nothing
        // else. Until the three Production* settings below are set, that is switched OFF and the
        // service is as unable to write Production as it always was. See DEPLOYMENT.md step 3.
        public string? ProductionTreeRoot { get; set; }

        // -----------------------------------------------------------------------------------
        // Publishing to PRODUCTION (owner request, 2026-09-25). All three, or none.
        //
        // NONE means the feature is off: the Maintainer tab says so, and nothing can write
        // Production. That is the safe answer for a service configured before this existed, which
        // is why these three have no default and are not required.
        //
        // ALL THREE are the Production twins of DataTreeRoot, ManifestPath and PublicDataBaseUrl.
        // Each is a separate value for the reason those are: the data root, the manifest beside it
        // and the URL clients fetch from cannot be derived from one another, and pairing the wrong
        // two produces a manifest that sends every client to the wrong tree with nothing failing.
        // The validator refuses any of them carrying the BETA marker, and any that equals its BETA
        // twin.
        // -----------------------------------------------------------------------------------
        public string? ProductionDataTreeRoot { get; set; }

        public string? ProductionManifestPath { get; set; }

        public string? ProductionPublicDataBaseUrl { get; set; }

        public bool IsProductionPublishingConfigured =>
            !string.IsNullOrWhiteSpace(this.ProductionDataTreeRoot) &&
            !string.IsNullOrWhiteSpace(this.ProductionManifestPath) &&
            !string.IsNullOrWhiteSpace(this.ProductionPublicDataBaseUrl);

        // ###########################################################################################
        // The tree the Systems screen READS as the stable source: the promotion's own root, else the
        // older ProductionTreeRoot - or null when neither is set. Reading needs no publishing switched
        // on. ONE rule for every reader (code review, 2026-10-04): the overview said a system was
        // in the stable source from this root, which offers CRT's "Data: Stable" switch, while the
        // stable table and files asked for publishing to be configured - so every pick of the switch
        // answered 404 on a server that only had ProductionTreeRoot.
        // ###########################################################################################
        public string? StableSourceRoot =>
            !string.IsNullOrWhiteSpace(this.ProductionDataTreeRoot)
                ? this.ProductionDataTreeRoot
                : string.IsNullOrWhiteSpace(this.ProductionTreeRoot) ? null : this.ProductionTreeRoot;

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
        // Feedback from CRT's Feedback tab (owner request, 2026-10-03 - the old PHP page's job).
        //
        // FeedbackRoot is where each feedback's attached files are saved, one "feedback-<random>"
        // folder per feedback - the PHP page's /mydir/http/classic-repair-toolbox.dk/user-feedback,
        // which the project owner opens from a network share. FeedbackToAddress is who the mail
        // goes to. NO DEFAULT for either, by this class's rule: one decides where strangers' files
        // land, the other who reads their logs. The folder must sit outside the web document root
        // and away from the data trees and the blob store (the validator checks the latter).
        // -----------------------------------------------------------------------------------
        public string? FeedbackRoot { get; set; }

        public string? FeedbackToAddress { get; set; }

        // The most the saved feedback under FeedbackRoot may take all together (code review,
        // 2026-10-04). Anybody may send feedback, so without a total, files could fill the disk
        // down to MinimumFreeDiskBytes - the reserve the blob store shares - and pause every
        // contribution. Full, a feedback's text is still mailed, its files are not saved, and the
        // mail says so. Tuning, so it has a default: 20 GiB. 0 turns the total off.
        public long FeedbackMaxStoredBytes { get; set; } = 20L * 1024 * 1024 * 1024;

        // -----------------------------------------------------------------------------------
        // Safety marker.
        // -----------------------------------------------------------------------------------

        // Every configured data path and the public base URL must contain this marker. It is the
        // belt-and-braces half of the Production refusal: even if DataTreeRoot and
        // ProductionTreeRoot were somehow both wrong, a path with no "-BETA" in it does not start
        // the service.
        //
        // The project owner sets this to empty ONLY when deliberately pointing the service at
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
        // Session lifetime, in days since the session was LAST USED. A default is sane, so this is
        // tuning rather than safety.
        //
        // *** THERE IS ONE TOKEN, AND THIS IS ITS LIFETIME. *** Login answers a `refreshToken`, and
        // that very value is the bearer token every request presents; AuthenticateAsync slides its
        // expiry forward as it is used (SessionExtensionRules). There is no separate short-lived
        // access token. An `AccessTokenMinutes` setting used to sit here describing one - it was
        // validated at startup and read by nothing, so it promised a protection that did not exist
        // (security review, 2026-09-25). It was removed rather than implemented, because a desktop
        // client holding a rotating token in a file is exactly what SessionExtensionRules' header
        // explains the project owner decided against. An old appsettings file that still carries the
        // key is harmless: an unknown key binds to nothing.
        // -----------------------------------------------------------------------------------
        public int RefreshTokenDays { get; set; } = 30;

        // -----------------------------------------------------------------------------------
        // The free space the blob store's disk must keep (security review, 2026-09-25).
        //
        // New submissions and upload chunks are refused while the disk holding BlobStoreRoot has
        // less than this free, so an anonymous sender with many addresses can pause contributions
        // but cannot fill the disk the web site and the database share. Tuning, so it has a
        // default: 5 GiB is far more than any honest day of contributions and small beside any
        // disk this service runs on. 0 turns the reserve off.
        // -----------------------------------------------------------------------------------
        public long MinimumFreeDiskBytes { get; set; } = 5L * 1024 * 1024 * 1024;

        // -----------------------------------------------------------------------------------
        // Whether board views sent from the server's OWN network are stored (owner request,
        // 2026-09-27: "For now I would like my own home usage also to count, as we then can check
        // the numbers and see how it works and looks like").
        //
        // A request arriving from a local-network address is the project owner's own machines -
        // see SenderAddress. ON, their views are stored like anybody's, with the country of the
        // server's own public address, and each row is marked `fromLocalNetwork = 1` so they can be
        // told apart or deleted later. OFF, such a batch is accepted and nothing of it is stored -
        // the launch check-in's own rule. Tuning, not safety, so it has a default: on, "for now".
        // -----------------------------------------------------------------------------------
        public bool CountLocalNetworkBoardViews { get; set; } = true;

        // -----------------------------------------------------------------------------------
        // Whether Account > "Reset contribution data" may delete (owner request, 2026-10-04: a clean
        // start at go-live) - see DataResetFlow. OFF by default and meant to stay off: the project
        // owner sets it true in appsettings.Production.json, restarts the service, resets, and sets
        // it back. So a click alone - or a stolen administrator session - can never wipe the
        // database; switching it on needs a shell on the box. The Account screen still shows the
        // counts while it is off, and says how to switch it on.
        // -----------------------------------------------------------------------------------
        public bool AllowDataReset { get; set; }

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
