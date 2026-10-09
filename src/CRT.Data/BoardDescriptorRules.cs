using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The derivations behind system.json (NewContributeStrategy.md Phase 4, task 7): a board's
    // identity, and the content hash that says which published state a client holds.
    //
    // BOTH ARE PURE AND BOTH ARE LOAD-BEARING, which is why they live here rather than inside
    // whatever writes the file. A BoardId that disagrees with the database is a submission that
    // cannot be stored; a ContentHash that disagrees between two machines makes every client
    // believe it is permanently out of date.
    //
    // WHAT THIS FILE DOES NOT DO: it does not write system.json, and nothing in the running server
    // may. The Production tree is denied to the service AT THE FILESYSTEM (Phase 3 step 0's
    // interlock), so publishing stays a deliberate act by the project owner. These helpers produce
    // the values; the publishing tool the project owner runs is what puts them on disk.
    // ###########################################################################################
    public static class BoardDescriptorRules
    {
        // ###########################################################################################
        // *** THE BOARD ID IS "Manufacturer/Hardware/Board", AND THAT IS A DECISION ALREADY TAKEN. ***
        //
        // It is written down in 0001_initial.sql, on the `boards` table, in these words: "board_id
        // is the STRING key the data tree already uses (manufacturer/hardware/board), not a
        // surrogate number, because it is what the desktop app, the file tree and every submission
        // already name."
        //
        // A RANDOM SURROGATE ID WAS BUILT HERE FIRST AND WAS WRONG. The argument for it was that a
        // random id survives a rename while a path-shaped one does not - which is true, and
        // irrelevant: NOTHING IN CRT CAN RENAME A BOARD. A board's folder path is its identity in
        // the sync manifest, in BoardDataReader's cache key, in DraftManager.GetBoardFolder and in
        // the worklog board key. SubmissionManifest.Renames covers ROWS INSIDE a board (components,
        // schematics), never the board itself. So the rename that the surrogate id protected
        // against cannot happen, and paying for it meant a second identity for every board, a
        // minting step, and a database whose primary key the client could not compute.
        //
        // If a board ever does need renaming, that is a migration with a redirect - not a reason
        // to carry a second identity forever on the chance.
        //
        // The three parts are joined with "/" because that is what ExcelDataFile, the sync manifest
        // and the published tree all already use (see NewBoardIdentity), so the id reads as the
        // path it is. NO trailing file name: the .xlsx names a file, and for a draft-only board it
        // names one that does not exist - neither belongs in an identity.
        // ###########################################################################################
        public static string BuildBoardId(string? manufacturer, string? hardware, string? board)
        {
            string cleanManufacturer = NewBoardIdentity.SanitizePathSegment(manufacturer);
            string cleanHardware = NewBoardIdentity.SanitizePathSegment(hardware);
            string cleanBoard = NewBoardIdentity.SanitizePathSegment(board);

            if (cleanManufacturer.Length == 0 || cleanHardware.Length == 0 || cleanBoard.Length == 0)
                return string.Empty;

            return $"{cleanManufacturer}/{cleanHardware}/{cleanBoard}";
        }

        // ###########################################################################################
        // The board id for a board identified by its ExcelDataFile - the form the app carries
        // everywhere ("Commodore/C64/250407/Data C64 250407.xlsx").
        //
        // Drops the file name and keeps the folder segments, which is exactly the mapping
        // DraftManager.GetBoardFolder already makes from the same string. Having this here means
        // the app never has to take the id apart by hand at a call site, which is where the two
        // forms would drift.
        // ###########################################################################################
        public static string BoardIdFromExcelDataFile(string? excelDataFile)
        {
            if (string.IsNullOrWhiteSpace(excelDataFile))
                return string.Empty;

            string[] segments = excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);

            // Three folders plus a file name is the shape NewBoardIdentity builds and the shape
            // the data tree uses. Anything shorter cannot name a board.
            if (segments.Length < 4)
                return string.Empty;

            return string.Join('/', segments.Take(segments.Length - 1));
        }

        // ###########################################################################################
        // Whether a string is a well-formed board id.
        //
        // Checked wherever one arrives from outside (a manifest, a system.json on disk), because
        // this value is a DATABASE PRIMARY KEY and a folder lookup. The rules are the data tree's
        // own: exactly three "/"-separated segments, each one a name that could really be a folder.
        //
        // Reusing NewBoardIdentity.IsValidPathSegment is the point rather than a convenience: it
        // already refuses traversal, reserved device names, trailing dots, control characters and
        // every character ANY platform reserves. A separate rule here would be a second opinion
        // about what a folder name may be, and the two would eventually disagree - at which point
        // a board could be created locally and refused by the server, or worse, the reverse.
        // ###########################################################################################
        public static bool IsValidBoardId(string? boardId)
        {
            if (string.IsNullOrWhiteSpace(boardId))
                return false;

            // The database column is VARCHAR(255). Refusing here rather than letting MariaDB
            // truncate silently, which would map two different boards onto one key.
            if (boardId.Length > MaximumBoardIdLength)
                return false;

            // Split WITHOUT RemoveEmptyEntries, so "Commodore//250407" is caught as the malformed
            // thing it is rather than quietly collapsing into a valid two-segment id.
            string[] segments = boardId.Split('/');

            if (segments.Length != 3)
                return false;

            return segments.All(segment => NewBoardIdentity.IsValidPathSegment(segment, out _));
        }

        // Three segments of at most NewBoardIdentity.MaxSegmentLength plus two separators, which
        // is 194 - comfortably inside the column's 255 and derived from the segment rule rather
        // than guessed, so raising one raises the other.
        public const int MaximumBoardIdLength = (NewBoardIdentity.MaxSegmentLength * 3) + 2;

        // ###########################################################################################
        // The hash of a published board's content, so a client can tell whether it holds the
        // current revision without comparing every file.
        //
        // *** THE INPUT IS THE FILE LIST, ORDERED AND DELIMITED - NOT A CONCATENATION. ***
        //
        // Three decisions here, each of which produces a silently wrong hash if taken the other
        // way:
        //
        //   - ORDERED, because a directory walk returns files in whatever order the filesystem
        //     feels like. Two machines hashing the same published tree must agree, so the paths
        //     are sorted ORDINALLY (never culture-aware - a Turkish or Swedish collation reorders
        //     them and the hash changes with the server's locale).
        //
        //   - DELIMITED with a character that cannot occur in the input. Concatenating
        //     "a.png" + hash + "b.png" + hash is ambiguous: a file named "a.pngb" would produce
        //     the same byte stream as two differently named files, which is a collision an
        //     attacker can construct deliberately. "\n" is the delimiter and a path containing one
        //     is refused outright rather than escaped, because SubmissionPathRules already refuses
        //     control characters - so such a path cannot reach a published tree in the first place,
        //     and accepting one here would mean the rules disagree.
        //
        //   - CASE-SENSITIVE throughout, like every other path comparison in this project. The
        //     Linux server is the filesystem that matters, and "U8.png" and "u8.PNG" are two files
        //     there.
        //
        // The revision is folded in too, so a metadata-only change (a corrected revision date with
        // identical files) still produces a new hash and clients pick it up.
        // ###########################################################################################
        public static string ComputeContentHash(string revision, IEnumerable<BoardContentEntry> files)
        {
            ArgumentNullException.ThrowIfNull(files);

            var builder = new StringBuilder();

            // The revision first, on its own line, so it can never be confused with a path.
            builder.Append("revision\n");
            builder.Append(revision ?? string.Empty);
            builder.Append('\n');

            List<BoardContentEntry> ordered = files
                .OrderBy(file => file.Path, StringComparer.Ordinal)
                .ToList();

            foreach (BoardContentEntry file in ordered)
            {
                if (file.Path.Contains('\n') || file.Sha256.Contains('\n'))
                {
                    throw new ArgumentException(
                        $"A published path or hash contains a newline, which cannot be hashed unambiguously: [{file.Path}]",
                        nameof(files));
                }

                builder.Append(file.Path);
                builder.Append('\n');
                builder.Append(file.Sha256);
                builder.Append('\n');
            }

            byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString()));
            return Convert.ToHexStringLower(hash);
        }

        // ###########################################################################################
        // Builds the descriptor for a board about to be published.
        //
        // The id is DERIVED from the three name parts rather than passed in, because with a
        // path-shaped id there is nothing else it could be - and deriving it here means the
        // descriptor's id and its Manufacturer/Hardware/Board can never disagree, which a
        // hand-assembled one could.
        // ###########################################################################################
        public static BoardDescriptor Build(
            string manufacturer,
            string hardware,
            string board,
            string revision,
            DateTimeOffset publishedUtc,
            IEnumerable<string>? maintainers,
            string origin,
            IEnumerable<BoardContentEntry> files)
        {
            return new BoardDescriptor
            {
                BoardId = BoardDescriptorRules.BuildBoardId(manufacturer, hardware, board),
                Manufacturer = NewBoardIdentity.SanitizePathSegment(manufacturer),
                Hardware = NewBoardIdentity.SanitizePathSegment(hardware),
                Board = NewBoardIdentity.SanitizePathSegment(board),
                Revision = revision ?? string.Empty,
                PublishedUtc = publishedUtc,

                // De-duplicated and ordered, so republishing an unchanged board does not produce
                // a file that differs only in the order a query happened to return rows.
                Maintainers = (maintainers ?? [])
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Select(name => name.Trim())
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(name => name, StringComparer.Ordinal)
                    .ToList(),

                Origin = origin ?? string.Empty,
                ContentHash = BoardDescriptorRules.ComputeContentHash(revision ?? string.Empty, files)
            };
        }

        // ###########################################################################################
        // Where a board came from. These are the two values the `boards` table's own `origin`
        // column was created to hold ("shipped with CRT, or contributed and vetted").
        //
        // Strings rather than an enum because this value is written to a JSON file read by older
        // and newer builds alike, and an unknown enum member deserialises to whatever happens to
        // be zero.
        //
        // IT IS SET ONCE, WHEN THE BOARD ROW IS FIRST CREATED, and never recomputed: a board
        // that arrived through the contribution pipeline stays "contributed" however many times it
        // is later revised, including revisions made by the project owner. It records where the
        // board CAME FROM, not who touched it last.
        // ###########################################################################################
        public static class BoardOrigin
        {
            // Came with CRT - the boards in Assets/Data that ship with the application.
            public const string Shipped = "shipped";

            // Arrived through the contribution pipeline and was vetted by a maintainer.
            public const string Contributed = "contributed";

            // Whether a value is one this build understands. An unrecognised origin is not an
            // error - a newer server may know more of them - but a caller choosing an icon or a
            // label needs to know when it is looking at something it cannot name.
            public static bool IsKnown(string? origin) =>
                origin is BoardOrigin.Shipped or BoardOrigin.Contributed;
        }
    }

    // ###########################################################################################
    // One file in a published board, for content hashing: its path relative to the board folder
    // ("/"-separated, as everything path-shaped in this project is) and the SHA-256 of its bytes.
    //
    // The hash is taken as given rather than computed here, because the caller has already hashed
    // these files to publish them and re-reading a 76 MB board to hash it a second time would be
    // pure waste.
    // ###########################################################################################
    public readonly record struct BoardContentEntry(string Path, string Sha256);
}
