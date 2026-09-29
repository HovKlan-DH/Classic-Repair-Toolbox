using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The wire contract between CRT and the server for submitting a system. ONE definition,
    // referenced by both sides - which is the whole reason CRT.Data was extracted in Phase 1.
    //
    // The old ComponentContributionPayload had the app and the PHP each carrying their own idea of
    // the shape, kept in step by hand and by a note in CLAUDE.md telling whoever changed one to
    // remember the other. That is the defect this replaces: here, a field added on one side does
    // not compile on the other until it is handled.
    //
    // WHAT IS SUBMITTED, SEMANTICALLY: the complete system as it should be after the change.
    // WHAT IS UPLOADED, PHYSICALLY: only the blobs the server does not already hold.
    //
    // The three-step transport (NewContributeStrategy.md, "The transport"):
    //
    //   1. Client POSTs a SubmissionManifest - every file path with its SHA-256, plus the board
    //      rows. Kilobytes of JSON even for a 76 MB system.
    //   2. Server answers with a HashNegotiationResponse naming the hashes it lacks.
    //   3. Client uploads only those blobs, then finalises.
    //
    // A typo fix therefore uploads a manifest and nothing else. A shared image already present
    // under another board is never re-sent, because THE HASH IS THE IDENTITY - the same idea
    // dataChecksums.json already proves at scale.
    //
    // ROW IDENTITY IS NOT IN THE DATA. There is no UuidV4 anywhere in this contract. The server
    // diffs two states it fully knows - the base revision and the submitted state - and pairs rows
    // on the natural keys in BoardDraftNaturalKeys, the same keys the draft layer already uses.
    // The one case natural keys cannot see is a RENAME, which reads as delete-plus-add; that is
    // why Renames is carried explicitly, since the client is the only party that knows.
    // ###########################################################################################
    public static class SubmissionFormat
    {
        // ###########################################################################################
        // Bumped whenever this contract changes in a way an older client would get wrong.
        //
        // Version 1 is the first whole-system submission format. It is deliberately NOT a
        // continuation of ComponentContributionPayload's numbering: that contract described one
        // component, this describes an entire system, and pretending they are the same sequence
        // would let an old client's "PayloadFormat 2" be mistaken for something this understands.
        //
        // The server rejects a submission whose version it does not know, with an update-required
        // message - the same approach $minimumContributionVersion takes in the PHP, and for the
        // same reason: silently accepting a payload you cannot fully interpret is how half-applied
        // data gets published.
        // ###########################################################################################
        public const int CurrentVersion = 1;

        // Guards a manifest against being absurd before any of it is trusted. Not security - the
        // server re-validates everything - but a client that has gone wrong should be told so
        // rather than uploading for an hour first.
        public const int MaximumFilesPerSubmission = 5000;

        // A single blob larger than this is refused. The largest legitimate file in the shipped
        // tree is a board scan of a few tens of megabytes; 256 MB leaves generous headroom while
        // still bounding what one submission can cost.
        public const long MaximumBlobBytes = 256L * 1024 * 1024;

        // SHA-256 as lowercase hex, which is what dataChecksums.json already uses.
        public const int HashLength = 64;

        // ###########################################################################################
        // How many bytes CRT sends per upload request. 4 MB is large enough that per-request
        // overhead is irrelevant and small enough that a dropped connection costs seconds.
        //
        // *** IN THE CONTRACT, NOT IN THE CLIENT (security review, 2026-09-25). *** The server now
        // caps each chunk request's body (RequestBodyLimits), and derives that cap from THIS value.
        // Kept as a private constant in the client, a larger chunk there would compile, ship, and
        // have every upload refused with 413 by a server that never heard of the change.
        // ###########################################################################################
        public const int UploadChunkBytes = 4 * 1024 * 1024;

        // ###########################################################################################
        // The longest a few free-text fields may be. Each equals the database column it lands in
        // (submissions.summary VARCHAR(500), submissions.base_revision and
        // systems.current_revision VARCHAR(64)).
        //
        // *** REFUSED, NOT TRUNCATED, AND REFUSED EARLY. *** An over-long value used to reach the
        // INSERT and fail there under strict mode, answering a 500 - and the revision date only
        // reaches its column AFTER a publish has already written the tree, so an over-long one
        // turned a successful, irreversible publish into an error the maintainer would retry.
        // ###########################################################################################
        public const int MaximumSummaryLength = 500;

        public const int MaximumRevisionLength = 64;
    }

    // ###########################################################################################
    // Step 1: what the client says the system should look like after the change.
    //
    // BaseRevision is the published revision the contributor started from. It is what makes this
    // a DELTA without the client having to compute one: the server knows that revision's state and
    // the submitted state, so it can diff them itself. A submission whose base is no longer current
    // needs re-basing rather than blind merging, and recording it here is what makes that
    // detectable rather than silent.
    //
    // It is EMPTY for a brand-new system - there is no published revision to have started from.
    // The server decides new-versus-update by whether SystemId already exists - one lookup -
    // rather than by trusting a flag in the payload, which is the "server figures it out"
    // behaviour the project owner asked for.
    // ###########################################################################################
    public sealed class SubmissionManifest
    {
        public int FormatVersion { get; set; } = SubmissionFormat.CurrentVersion;

        // ###########################################################################################
        // "Manufacturer/Hardware/Board" - the `systems` table's primary key, and the same key the
        // data tree, the sync manifest and every draft folder already use.
        //
        // ALWAYS SET, including for a system the contributor invented five minutes ago. That is
        // the point of a path-shaped id: the client can compute it from what was typed, so there
        // is no chicken-and-egg where a new system has no identity until the server grants one.
        // What makes a submission NEW is that no `systems` row carries this id yet, which is a
        // lookup the server does - never a flag the client sets.
        //
        // See SystemDescriptorRules.BuildSystemId for why this is the data tree's own key rather
        // than a random surrogate.
        // ###########################################################################################
        public string SystemId { get; set; } = string.Empty;

        public string Manufacturer { get; set; } = string.Empty;
        public string Hardware { get; set; } = string.Empty;
        public string Board { get; set; } = string.Empty;

        // The published revision this was drafted from. Empty for a new system.
        public string BaseRevision { get; set; } = string.Empty;

        // Free text from the contributor saying what they changed and why. This is what a maintainer
        // reads first, so it is part of the contract rather than an afterthought.
        public string Summary { get; set; } = string.Empty;

        // ###########################################################################################
        // Where to tell the contributor whether their work was accepted.
        //
        // *** THIS IS NOT A CREDENTIAL AND NOT AN IDENTITY. *** Contributing requires no account
        // (NewContributeStrategy.md, "CONTRIBUTING NEEDS NO ACCOUNT"): nothing signs in with this
        // address, there is no password beside it, and two submissions carrying the same address
        // are two unrelated submissions rather than one contributor's history.
        //
        // It is required because a contribution nobody can reply to can only be accepted or
        // dropped - never improved. "Your highlight coordinates look wrong, did you mean X?" is
        // the message this exists to make possible.
        //
        // A signed-in maintainer submitting to their own system may leave it empty; their account
        // already carries an address.
        // ###########################################################################################
        public string ContactEmail { get; set; } = string.Empty;

        // Which CRT produced this, for diagnosing a client-specific problem later. Not used for
        // any decision - the format version is what gates acceptance.
        public string ApplicationVersion { get; set; } = string.Empty;

        public DateTimeOffset CreatedUtc { get; set; }

        // Every file the system should contain afterwards, each with its hash. A file present in
        // the base revision but ABSENT here is a deletion - the manifest is the complete intended
        // state, not a list of changes.
        public List<SubmissionFile> Files { get; set; } = new();

        // The board rows, as the system should read afterwards. Same "complete state" rule.
        public SubmissionRows Rows { get; set; } = new();

        // Explicit rename intent, because natural-key pairing cannot see a rename - U8 becoming U9
        // is indistinguishable from deleting U8 and adding U9 unless somebody says otherwise, and
        // the client is the only party that knows. Without this, a rename silently discards the
        // row's history and any per-row data keyed to the old name.
        public List<SubmissionRename> Renames { get; set; } = new();
    }

    // ###########################################################################################
    // One file in the submitted system.
    //
    // Path is RELATIVE to the system's own folder and is UNTRUSTED INPUT - it arrives over the
    // network and names a location the server will write to. SubmissionPathRules is the one place
    // that is checked; see its header for what it rejects and why a plain "does it start with .."
    // test is not enough.
    //
    // *** PATH CASE IS SIGNIFICANT AND MUST BE PRESERVED EXACTLY. *** From Phase 3 the data tree
    // lives on Linux, where CRT.Data's exact-case File.Exists means a row naming "foo.pdf" when the
    // file is "foo.PDF" works on the contributor's Windows machine and fails only on the server.
    // Mixed case is real in the shipped tree (27 .PDF alongside 85 .pdf), so this cannot be
    // normalised away - and normalising it would hide contributed data that is genuinely wrong
    // while leaving Windows clients disagreeing with the server about which file is meant.
    // ###########################################################################################
    public sealed class SubmissionFile
    {
        public string Path { get; set; } = string.Empty;

        // Lowercase hex SHA-256 of the file's bytes. This is the blob's identity: two systems
        // carrying the same image reference the same hash and it is stored once.
        public string Sha256 { get; set; } = string.Empty;

        public long SizeBytes { get; set; }
    }

    // ###########################################################################################
    // The board rows as the system should read after the change, one list per BoardData section.
    //
    // These are the SAME row types the app and the reader already use, not parallel DTOs. A
    // separate set would be a second definition of the schema to keep in step - exactly what this
    // contract exists to eliminate.
    // ###########################################################################################
    public sealed class SubmissionRows
    {
        // ###########################################################################################
        // The board's own revision date - the human-authored text in the workbook's
        // "# Revision date:" marker, not a machine version.
        //
        // *** IT IS ALSO WHAT `systems.current_revision` BECOMES, and that is a contract with the
        // CLIENT. *** DraftBaseRevision stamps a draft's BaseRevision from the official board's
        // revision date, and the submission sends that value back to be diffed against. So the
        // published revision has to be the same thing, or every contributor would be re-basing
        // against a value their drafts never carry and the drift warning would fire forever.
        //
        // *** ADDED 2026-09-22, and OPTIONAL on purpose. *** It was missing, which meant two
        // things: a maintainer could not see a revision-date change (ReviewEndpoints had to use the
        // published date for both sides), and a contributor could not publish one at all. An older
        // client simply omits it, and PublishMerge then KEEPS the published date rather than
        // blanking it - so this needed no format-version bump.
        // ###########################################################################################
        public string RevisionDate { get; set; } = string.Empty;

        public List<BoardSchematicEntry> Schematics { get; set; } = new();
        public List<ComponentEntry> Components { get; set; } = new();
        public List<ComponentImageEntry> ComponentImages { get; set; } = new();
        public List<ComponentHighlightEntry> ComponentHighlights { get; set; } = new();

        // Component-scoped and board-scoped files and links are SEPARATE sections in BoardData,
        // and stay separate here. They carry different natural keys and a maintainer needs to see
        // which is which - a datasheet attached to U8 is a different claim from one attached to
        // the board as a whole.
        public List<ComponentLocalFileEntry> ComponentLocalFiles { get; set; } = new();
        public List<ComponentLinkEntry> ComponentLinks { get; set; } = new();
        public List<BoardLocalFileEntry> BoardLocalFiles { get; set; } = new();
        public List<BoardLinkEntry> BoardLinks { get; set; } = new();

        public List<CreditEntry> Credits { get; set; } = new();
        public List<KiCadImportantSignalEntry> KiCadImportantSignals { get; set; } = new();

        // The KiCad trace calibration, which is per-schematic rather than per-component. Included
        // because a contributor who calibrates a board's traces has done substantial work that
        // would otherwise not travel with their submission.
        public List<KiCadCalibrationEntry> KiCadCalibrations { get; set; } = new();
    }

    // ###########################################################################################
    // "This row is the one that used to be called that."
    //
    // Section names the BoardData section the rename applies to, and From/To are natural keys as
    // BoardDraftNaturalKeys builds them - not display text - so the server can pair rows without
    // knowing which fields make up a key for that section.
    // ###########################################################################################
    public sealed class SubmissionRename
    {
        public string Section { get; set; } = string.Empty;
        public string From { get; set; } = string.Empty;
        public string To { get; set; } = string.Empty;
    }

    // ###########################################################################################
    // Step 2: the server's answer - which blobs it does not already hold.
    //
    // SubmissionId identifies this submission for the upload and finalise steps that follow. It is
    // issued by the server rather than chosen by the client, so a client cannot address another
    // contributor's in-flight submission by guessing.
    //
    // MissingHashes is usually EMPTY: most submissions are edits to systems whose files the server
    // already has, which is what makes a typo fix cost a single round trip.
    // ###########################################################################################
    public sealed class HashNegotiationResponse
    {
        public long SubmissionId { get; set; }

        // ###########################################################################################
        // The capability token for this submission. RETURNED ONCE, HERE, AND NEVER AGAIN.
        //
        // Contributing requires no account, so there is no identity to authorise against - and the
        // submission id is a small consecutive integer anyone could guess. This token is what every
        // later call for this submission must present. The client must keep it for the duration of
        // the upload; losing it means starting the submission again.
        //
        // The server stores only its hash, so it cannot be recovered - by the contributor or by
        // anyone reading the database.
        // ###########################################################################################
        public string UploadToken { get; set; } = string.Empty;

        public List<string> MissingHashes { get; set; } = new();

        // What the server already holds, so the client can show "uploading 3 of 240 files" rather
        // than a bare count of what is left.
        public int AlreadyHeldCount { get; set; }

        public long TotalBytesToUpload { get; set; }
    }

    // ###########################################################################################
    // Step 3: the outcome of finalising.
    //
    // Accepted means QUEUED FOR REVIEW, never published. Nothing in this pipeline writes to the
    // Production tree - promotion stays a manual act by the project owner (Phase 3 step 3's interlock
    // makes it impossible for the service to do otherwise).
    //
    // A submission that fails automated validation is rejected HERE, before any human sees it, with
    // the reasons attached. That is the highest-leverage part of the whole plan: it turns most bad
    // submissions into a fast, polite, automatic answer rather than maintainer time.
    // ###########################################################################################
    public sealed class SubmissionResult
    {
        public long SubmissionId { get; set; }

        public bool IsAccepted { get; set; }

        // Present when IsAccepted is false. Written to be read by the contributor, naming the row
        // or file at fault - a rejection that does not say what to fix produces a resubmission of
        // the same thing.
        public List<ValidationFinding> Findings { get; set; } = new();

        public string State { get; set; } = string.Empty;
    }

    // ###########################################################################################
    // What the server says about one already-sent submission - the reply to GET
    // /api/submissions/{id}, which the "my submissions" view (Phase 4, task 6) asks for.
    //
    // AUTHORISED BY THE CAPABILITY TOKEN, not by an account: contributing needs no sign-in, so
    // holding the token returned at creation is the whole proof of ownership. A request without it
    // is answered 404 rather than 403, so the small sequential id space cannot be walked to learn
    // which submissions exist.
    //
    // MaintainerComment is carried now and filled by Phase 5. The Maintainer tab does not exist
    // yet, so the server has no field for it and it arrives empty - defined here so that adding it
    // server-side later needs no contract version bump and no change on any contributor's disk.
    // ###########################################################################################
    public sealed class SubmissionStatus
    {
        public long Id { get; set; }

        public string SystemId { get; set; } = string.Empty;

        // The server's own vocabulary ("uploading", "pending", "rejected", "abandoned"). NOT shown
        // to the user as-is - SubmissionReceiptPresenter.DescribeState turns it into words that
        // mean the same thing to somebody who is not holding the database schema.
        public string State { get; set; } = string.Empty;

        public string Summary { get; set; } = string.Empty;

        public DateTimeOffset CreatedUtc { get; set; }

        public DateTimeOffset? DecidedUtc { get; set; }

        public string MaintainerComment { get; set; } = string.Empty;

        // A maintainer changed some of the submission's rows in the Maintainer tab before
        // deciding it (2026-09-25) - what is published is then not exactly what was sent, and the
        // contributor is told so in "My submissions".
        public bool AmendedByMaintainer { get; set; }

        public List<ValidationFinding> Findings { get; set; } = new();
    }

    // ###########################################################################################
    // One thing wrong with a submission.
    //
    // Severity decides the outcome: any Error rejects the submission automatically and it is never
    // queued. A Warning is attached for the maintainer to weigh and does not block.
    //
    // Subject names the offending row or file - the board label, the path - so a contributor with
    // 400 components can find the one at fault. A finding without a subject is nearly useless to
    // whoever has to act on it.
    // ###########################################################################################
    public sealed class ValidationFinding
    {
        public ValidationSeverity Severity { get; set; }

        // A short stable code (e.g. "file.missing", "highlight.degenerate") so the client can
        // group or translate findings without parsing prose.
        public string Code { get; set; } = string.Empty;

        public string Subject { get; set; } = string.Empty;

        // Written for the contributor, not for a log: it says what is wrong and what to do.
        public string Message { get; set; } = string.Empty;
    }

    public enum ValidationSeverity
    {
        Warning,
        Error
    }

    // ###########################################################################################
    // system.json, written per system in the published tree (Phase 4 task 7).
    //
    // SYSTEMID IS DERIVED ONCE AT CREATION AND NEVER CHANGES. That is the whole point: the folder
    // path carries Manufacturer/Hardware/Board, so renaming a board would otherwise lose its
    // history and its maintainer list. With a stable id, a rename is a metadata change and
    // everything keyed to the system survives it.
    //
    // ContentHash covers the published state, so a client can tell whether it holds the current
    // revision without comparing every file.
    // ###########################################################################################
    public sealed class SystemDescriptor
    {
        public string SystemId { get; set; } = string.Empty;
        public string Manufacturer { get; set; } = string.Empty;
        public string Hardware { get; set; } = string.Empty;
        public string Board { get; set; } = string.Empty;
        public string Revision { get; set; } = string.Empty;
        public DateTimeOffset PublishedUtc { get; set; }

        // Display names of the system's maintainers, for showing in the app. Authority itself is
        // decided by the maintainers table on the server, never by this file - a published file is
        // something a contributor could edit locally, so it must not be able to grant anything.
        public List<string> Maintainers { get; set; } = new();

        // "shipped" for a system that came with CRT, "contributed" for one that arrived through
        // this pipeline and was vetted.
        public string Origin { get; set; } = string.Empty;

        public string ContentHash { get; set; } = string.Empty;
    }
}
