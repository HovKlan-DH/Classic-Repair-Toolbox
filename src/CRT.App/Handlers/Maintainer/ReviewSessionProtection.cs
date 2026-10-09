using System;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ENCRYPTS THE STORED SESSION TOKEN so a file copied off the machine is worthless
    // (owner request, 2026-09-22 - "no risk in case of a malicious user").
    //
    // *** WHY THIS EXISTS AT ALL. *** The token it protects authorises reading every queued
    // contribution and, for an administrator, PUBLISHING to every user of a board. ReviewSession's
    // header originally forbade writing it to disk outright, and relaxed to "do not add remember-me
    // without deciding where that file lives and who can read it". This class is that decision.
    //
    // *** DPAPI, CurrentUser SCOPE. *** Windows derives the key from the logged-in user's own
    // credentials, so the ciphertext cannot be read by another account on the same machine, and
    // cannot be read AT ALL on a different machine - copying the file to another PC yields nothing.
    // That is the exact property wanted: the convenience stays on the one machine that earned it.
    //
    // *** ON NON-WINDOWS THERE IS NO EQUIVALENT, AND THE ANSWER IS TO REFUSE RATHER THAN PRETEND.
    // *** .NET's ProtectedData throws PlatformNotSupportedException off Windows. The tempting move
    // is to fall back to plaintext, or to "encryption" with a key sitting beside the file - both
    // give a file that LOOKS protected and is not, which is worse than an honest refusal because
    // nobody then thinks to check. So Protect returns null off Windows, the caller stores nothing,
    // and the maintainer signs in each launch exactly as before. A Linux maintainer loses a
    // convenience; nobody gains a false one.
    //
    // *** EVERY FAILURE IS SOFT. *** Decryption fails for ordinary reasons - the file was copied
    // from another machine, the Windows profile was rebuilt, the bytes were truncated by a bad
    // shutdown. None may crash the app or block sign-in; all degrade to "no stored session".
    // ###########################################################################################
    internal static class ReviewSessionProtection
    {
        // Distinguishes the maintainer session's ciphertext from anything else stored the same way. DPAPI binds
        // it into the encryption, so bytes protected with one entropy cannot be unprotected with
        // another - a stored CRT value could never be decrypted here even if the two ever shared
        // a file.
        //
        // *** IT STILL SAYS "Review", ON PURPOSE - DO NOT RENAME IT. *** The app became CRT
        // Maintainer on 2026-09-25 and CRT's Maintainer tab on 2026-09-29, but every session
        // remembered before then was encrypted with this value; changing it would make each of them undecryptable, silently signing everyone
        // out. ReviewSessionStore.AdoptLegacyFile carries those files over on that promise.
        private static readonly byte[] Entropy =
            Encoding.UTF8.GetBytes("Classic-Repair-Toolbox.Review.Session.v1");

        // ###########################################################################################
        // Whether a session can be stored safely on this machine at all.
        //
        // The UI reads this to decide whether to OFFER to stay signed in, so a Linux maintainer is
        // told the truth rather than being given a tick box that silently does nothing.
        // ###########################################################################################
        public static bool IsSupported => OperatingSystem.IsWindows();

        // ###########################################################################################
        // Encrypts, or returns NULL when this machine cannot do it safely.
        //
        // Null is a real, expected answer and the caller must handle it by storing nothing - never
        // by writing the plaintext it was handed.
        // ###########################################################################################
        public static byte[]? Protect(string plaintext)
        {
            // *** THE GUARD IS WRITTEN INLINE, NOT AS `!IsSupported`. *** CA1416 only recognises
            // OperatingSystem.IsWindows() itself as a platform guard; routing it through a
            // property leaves the analyser unable to prove the call is safe, and Release (which
            // treats warnings as errors) fails to build. Keep the call, not the abstraction.
            if (string.IsNullOrEmpty(plaintext) || !OperatingSystem.IsWindows())
                return null;

            try
            {
                return ReviewSessionProtection.ProtectWindows(Encoding.UTF8.GetBytes(plaintext));
            }
            catch (Exception)
            {
                // A machine that cannot protect the value must not store it unprotected.
                return null;
            }
        }

        // ###########################################################################################
        // Decrypts, or returns NULL for anything that is not our own ciphertext from this profile.
        //
        // The ordinary reasons to land here are all benign - a copied file, a rebuilt profile,
        // truncated bytes - and all mean the same thing to the caller: there is no usable stored
        // session, so ask for a password.
        // ###########################################################################################
        public static string? Unprotect(byte[]? ciphertext)
        {
            // Inline for the same reason as in Protect above - see the note there.
            if (ciphertext is null || ciphertext.Length == 0 || !OperatingSystem.IsWindows())
                return null;

            try
            {
                return Encoding.UTF8.GetString(ReviewSessionProtection.UnprotectWindows(ciphertext));
            }
            catch (Exception)
            {
                return null;
            }
        }

        // ###########################################################################################
        // The two Windows-only calls, isolated behind [SupportedOSPlatform] so the compiler proves
        // the guard above is real.
        //
        // *** WITHOUT THIS SPLIT THE RELEASE BUILD DOES NOT COMPILE. *** CRT.App treats warnings
        // as errors in Release, and calling ProtectedData without the platform attribute raises
        // CA1416 ("only supported on Windows"). Keeping these separate means the OperatingSystem
        // check is verified by the compiler rather than trusted.
        // ###########################################################################################
        [SupportedOSPlatform("windows")]
        private static byte[] ProtectWindows(byte[] plaintext) =>
            ProtectedData.Protect(plaintext, ReviewSessionProtection.Entropy, DataProtectionScope.CurrentUser);

        [SupportedOSPlatform("windows")]
        private static byte[] UnprotectWindows(byte[] ciphertext) =>
            ProtectedData.Unprotect(ciphertext, ReviewSessionProtection.Entropy, DataProtectionScope.CurrentUser);
    }
}
