using System;
using System.IO;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // Resolves the REAL, symlink-free filesystem path a path string ultimately points at, so a
    // containment check ("is this inside the data root") cannot be defeated by a symlinked
    // directory sitting somewhere along the path.
    //
    // WHY THIS EXISTS. Path.GetFullPath is purely LEXICAL: it collapses "..", resolves relative
    // segments and fixes separators, but it never touches the filesystem and so never notices
    // that a directory along the way is a symlink (or, on Windows, a junction) pointing somewhere
    // else entirely. ExternalTargetLauncher's containment check - and the same shape of check on
    // the contribution webserver - both compare a Path.GetFullPath()'d string against the data
    // root with StartsWith. A symlinked directory INSIDE the root pointing OUTSIDE it therefore
    // passes that check while the file it actually opens is not inside the root at all - proven
    // against a real Windows junction (mklink /J), which .NET's own Directory.ResolveLinkTarget
    // recognizes and follows exactly like a symlink.
    //
    // This is currently UNREACHABLE in this application: nothing writes a symlink into the data
    // root (OnlineServices' sync writes plain bytes byte-for-byte, and nothing extracts an
    // archive there), so the gap costs nothing today. It exists so the boundary is still correct
    // the day something does - a future contribution feature that unpacks a zip into the data
    // root, say, would otherwise reopen exactly this hole.
    // ###########################################################################################
    public static class RealPathResolver
    {
        // ###########################################################################################
        // Walks fullyQualifiedPath one directory component at a time, following any symlink or
        // junction found at each step to its FINAL target (ResolveLinkTarget already collapses a
        // whole chain of links in one call), and continuing the walk from there. On success
        // realPath is the symlink-free path - which for the overwhelmingly common case of nothing
        // along the path being linked is the input, unchanged.
        //
        // Returns FALSE when resolution could not be carried out at all: a blank or non-rooted
        // input, a malformed path, or a failure partway through the walk (permissions, a race with
        // something deleting the path, a malformed reparse point). realPath is then the plain
        // lexically-normalized input, which is a REASONABLE value but NOT a resolved one.
        //
        // WHY A BOOL RATHER THAN JUST THE STRING. The single caller resolves TWO paths and compares
        // them for containment. If one side resolves and the other silently falls back, the
        // comparison is made between a resolved and an unresolved path and reaches a verdict about
        // neither - it refused a legitimate file whenever the data root itself sat behind a link
        // and the root's own resolve happened to throw. A caller that cannot tell "nothing was
        // linked" from "I could not look" has no way to fail closed, so this reports which it was
        // and ExternalTargetLauncher refuses the open outright when resolution did not happen.
        //
        // Takes an ALREADY-ROOTED, already-Path.GetFullPath'd string. It does not itself call
        // GetFullPath, because a caller comparing two resolved paths needs both run through the
        // exact same normalization it already trusts, and doing that twice here would risk the two
        // disagreeing about it.
        //
        // Tolerates a path that does not exist (yet, or at all) and still reports success: nothing
        // that does not exist can be a symlink, so the walk stops resolving at the first missing
        // component and appends whatever remains of the input unresolved, rather than losing it.
        // ###########################################################################################
        public static bool TryResolveRealPath(string fullyQualifiedPath, out string realPath)
        {
            realPath = fullyQualifiedPath;

            if (string.IsNullOrEmpty(fullyQualifiedPath))
            {
                return false;
            }

            string? root;
            try
            {
                root = Path.GetPathRoot(fullyQualifiedPath);
            }
            catch
            {
                // A malformed path (invalid characters, etc.) - not this method's job to validate,
                // but equally not something it can claim to have resolved.
                return false;
            }

            if (string.IsNullOrEmpty(root))
            {
                // Not an absolute path. Callers are expected to have already normalized with
                // Path.GetFullPath, so this is a defensive fallback rather than the normal path.
                return false;
            }

            // Split on BOTH separators. Path.GetFullPath does not convert backslashes on Linux or
            // macOS, so a stored Windows-style path arriving here would otherwise be walked as one
            // giant component, no link along it would ever be resolved, and the containment check
            // would silently degrade to the purely lexical one this class exists to fix.
            string[] segments = fullyQualifiedPath
                .Substring(root.Length)
                .Split(
                    new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                    StringSplitOptions.RemoveEmptyEntries);

            string accumulated = root;

            try
            {
                for (int i = 0; i < segments.Length; i++)
                {
                    accumulated = Path.Combine(accumulated, segments[i]);

                    FileSystemInfo? linkTarget;

                    if (Directory.Exists(accumulated))
                    {
                        linkTarget = Directory.ResolveLinkTarget(accumulated, returnFinalTarget: true);
                    }
                    else if (File.Exists(accumulated))
                    {
                        linkTarget = File.ResolveLinkTarget(accumulated, returnFinalTarget: true);
                    }
                    else
                    {
                        // Nothing exists at this point in the path - not a not-yet-created
                        // intermediate directory, and not a symlink either, since neither can be
                        // true of something that is not there. Append whatever remains of the
                        // input UNRESOLVED and stop: there is nothing real left to walk.
                        //
                        // Joined rather than Path.Combine'd: Combine(params string[]) RESTARTS from
                        // the last rooted element, so a single remaining segment that
                        // Path.IsPathRooted considers rooted would throw the accumulated prefix
                        // away and hand back a path unrelated to the input.
                        accumulated = string.Join(
                            Path.DirectorySeparatorChar,
                            new[] { accumulated }.Concat(segments.Skip(i + 1)));
                        break;
                    }

                    if (linkTarget != null)
                    {
                        accumulated = linkTarget.FullName;
                    }
                }
            }
            catch
            {
                // A failure partway through resolution is not a reason to hand back a half-resolved
                // path that LOOKS fully resolved - that is worse than not resolving at all, since a
                // caller comparing it against the data root would trust it. Report the failure and
                // leave realPath as the plain lexically-normalized input.
                realPath = fullyQualifiedPath;
                return false;
            }

            realPath = accumulated;
            return true;
        }
    }
}
