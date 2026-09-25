using System.Globalization;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // Reads and writes the PHC string that a password hash is stored as:
    //
    //     $argon2id$v=19$m=65536,t=3,p=2$<salt-b64>$<hash-b64>
    //
    // PURE STRING WORK, no hashing here at all - Argon2PasswordHasher does that and calls this to
    // encode and decode. The split exists because parsing is where the subtle bugs live and it is
    // fully testable, while the hashing itself is a library call with nothing to get wrong.
    //
    // WHY THE PARAMETERS TRAVEL WITH THE HASH. This is the single most important property of the
    // format. Verification reads m, t and p from the STORED string, not from configuration, so
    // raising the cost in appsettings does not invalidate a single existing password: old hashes
    // keep verifying with the parameters they were made with, and NeedsRehash then reports that
    // they are below the current policy so the login path can upgrade them silently. A format that
    // did not carry its parameters would make every cost increase a mass password reset.
    //
    // WHY BASE64 WITHOUT PADDING. The PHC standard omits the '=' padding, and interoperability is
    // the point: a hash written here should be readable by any other Argon2 implementation, so a
    // future move off Konscious - or a migration to another language entirely - does not strand
    // every stored password. That is also why the format is not something of our own invention.
    //
    // THE PARSER IS DELIBERATELY STRICT. Anything it does not fully understand is rejected rather
    // than half-interpreted, because a partially-parsed hash is a hash that might verify against
    // the wrong parameters. A malformed stored hash must fail closed: the login fails, the
    // maintainer sees it, nobody is let in.
    // ###########################################################################################
    public static class PasswordHashEncoding
    {
        // The only algorithm this service writes. Argon2i and Argon2d are deliberately not
        // accepted on the way in either: accepting a variant we never produce would mean
        // supporting a verification path no test here exercises.
        public const string AlgorithmName = "argon2id";

        // Argon2 version 0x13 (19). The earlier 0x10 has a known weakness and is not accepted.
        public const int Version = 19;

        // ###########################################################################################
        // Builds the stored string. Salt and hash are raw bytes; everything else is the cost.
        // ###########################################################################################
        public static string Encode(
            int memoryKib,
            int iterations,
            int parallelism,
            byte[] salt,
            byte[] hash)
        {
            ArgumentNullException.ThrowIfNull(salt);
            ArgumentNullException.ThrowIfNull(hash);

            string saltPart = PasswordHashEncoding.ToUnpaddedBase64(salt);
            string hashPart = PasswordHashEncoding.ToUnpaddedBase64(hash);

            return string.Create(
                CultureInfo.InvariantCulture,
                $"${PasswordHashEncoding.AlgorithmName}$v={PasswordHashEncoding.Version}$" +
                $"m={memoryKib},t={iterations},p={parallelism}${saltPart}${hashPart}");
        }

        // ###########################################################################################
        // Reads a stored string back. Returns false for anything malformed rather than throwing:
        // the caller is a login path, and the correct response to an unreadable stored hash is to
        // fail the login, not to crash the request.
        //
        // NOTE the deliberate absence of a "best effort" branch. Every field must be present,
        // in order, and parse cleanly.
        // ###########################################################################################
        public static bool TryDecode(string? encoded, out DecodedPasswordHash decoded)
        {
            decoded = default;

            if (string.IsNullOrWhiteSpace(encoded))
                return false;

            // Leading '$' means Split produces an empty first element - that is expected, and its
            // absence means this is not a PHC string at all.
            string[] parts = encoded.Split('$');

            if (parts.Length != 6 || parts[0].Length != 0)
                return false;

            if (!string.Equals(parts[1], PasswordHashEncoding.AlgorithmName, StringComparison.Ordinal))
                return false;

            if (!PasswordHashEncoding.TryParseLabelled(parts[2], "v", out int version) ||
                version != PasswordHashEncoding.Version)
            {
                return false;
            }

            string[] costParts = parts[3].Split(',');

            if (costParts.Length != 3)
                return false;

            if (!PasswordHashEncoding.TryParseLabelled(costParts[0], "m", out int memoryKib) ||
                !PasswordHashEncoding.TryParseLabelled(costParts[1], "t", out int iterations) ||
                !PasswordHashEncoding.TryParseLabelled(costParts[2], "p", out int parallelism))
            {
                return false;
            }

            // A zero or negative cost parameter would make the hash trivially cheap. Such a string
            // should never exist, and if one does it must not be honoured.
            if (memoryKib <= 0 || iterations <= 0 || parallelism <= 0)
                return false;

            if (!PasswordHashEncoding.TryFromUnpaddedBase64(parts[4], out byte[]? salt) ||
                !PasswordHashEncoding.TryFromUnpaddedBase64(parts[5], out byte[]? hash))
            {
                return false;
            }

            if (salt.Length == 0 || hash.Length == 0)
                return false;

            decoded = new DecodedPasswordHash(memoryKib, iterations, parallelism, salt, hash);
            return true;
        }

        // ###########################################################################################
        // Is this stored hash weaker than what we would produce today?
        //
        // Called on every successful login so a password can be upgraded in place, at the one
        // moment the plaintext is legitimately available. A hash is NOT rehashed merely for being
        // different - only for being BELOW policy - because a hash made with stronger parameters
        // than the current configuration is fine, and downgrading it would be a regression.
        // ###########################################################################################
        public static bool NeedsRehash(
            DecodedPasswordHash decoded,
            int currentMemoryKib,
            int currentIterations,
            int currentParallelism)
        {
            return decoded.MemoryKib < currentMemoryKib
                || decoded.Iterations < currentIterations
                || decoded.Parallelism < currentParallelism;
        }

        // -------------------------------------------------------------------------------------
        // Helpers.
        // -------------------------------------------------------------------------------------

        // Parses "m=65536" given the label "m". The label must match exactly - a value arriving in
        // the wrong position (t where m was expected) is a malformed string, not something to
        // reinterpret.
        private static bool TryParseLabelled(string part, string label, out int value)
        {
            value = 0;

            if (part.Length < label.Length + 2)
                return false;

            if (!part.AsSpan(0, label.Length).SequenceEqual(label) || part[label.Length] != '=')
                return false;

            return int.TryParse(
                part.AsSpan(label.Length + 1),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out value);
        }

        private static string ToUnpaddedBase64(byte[] value)
        {
            return Convert.ToBase64String(value).TrimEnd('=');
        }

        private static bool TryFromUnpaddedBase64(string value, out byte[] bytes)
        {
            bytes = Array.Empty<byte>();

            if (value.Length == 0)
                return false;

            // Base64 decodes in 4-character groups, so restore whatever padding was stripped.
            // A remainder of 1 is impossible in valid base64 and is rejected by TryFromBase64String.
            int padding = (4 - (value.Length % 4)) % 4;
            string padded = value + new string('=', padding);

            byte[] buffer = new byte[((padded.Length / 4) * 3)];

            if (!Convert.TryFromBase64String(padded, buffer, out int written))
                return false;

            bytes = buffer[..written];
            return true;
        }
    }

    // ###########################################################################################
    // A stored hash, taken apart. The cost parameters here are the ones the hash was MADE with,
    // which is what verification must use - never the current configuration.
    // ###########################################################################################
    public readonly record struct DecodedPasswordHash(
        int MemoryKib,
        int Iterations,
        int Parallelism,
        byte[] Salt,
        byte[] Hash);
}
