using System;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // Decides which release versions a user has actually opted into, from the two Configuration
    // checkboxes ("Allow notification for ALPHA versions" / "... BETA versions").
    //
    // This exists because GitHub's own "prerelease" flag has no concept of stage - it either
    // considers every pre-release or none - so the stage has to be read back off the version string
    // itself. The convention is the release pipeline's own stage-label step in
    // .github/workflows/build-and-release.yml, and this class deliberately mirrors ITS default: a
    // pre-release whose suffix matches neither "-alpha." nor "-beta." is labelled ALPHA there, so it
    // is treated as ALPHA here. Reading an unrecognised suffix (-rc.1, -preview.1, a bare -rc1) as
    // "no stage at all" is what let such a build slip past the check entirely and be offered to a
    // user who had only ticked BETA.
    //
    // Pure string work so it can be unit tested; UpdateService applies it through
    // StageFilteredUpdateSource, which filters the feed BEFORE Velopack ranks it.
    // ###########################################################################################
    public static class UpdateChannelFilter
    {
        public enum ReleaseStage
        {
            Stable,
            Beta,
            Alpha,
        }

        // ###########################################################################################
        // The stage a version string belongs to. A version with no pre-release suffix at all (a bare
        // "2.5.0") is Stable; "-beta." is Beta; anything else carrying a "-" is Alpha, matching the
        // pipeline's own catch-all arm.
        //
        // Velopack renders a version's build metadata after a "+", which is not a pre-release marker,
        // so it is cut off before the "-" is looked for - "2.5.0+abc-def" is a stable release.
        // ###########################################################################################
        public static ReleaseStage GetStage(string? version)
        {
            if (string.IsNullOrWhiteSpace(version))
            {
                return ReleaseStage.Stable;
            }

            string withoutBuildMetadata = version;

            int plusIndex = withoutBuildMetadata.IndexOf('+');
            if (plusIndex >= 0)
            {
                withoutBuildMetadata = withoutBuildMetadata.Substring(0, plusIndex);
            }

            int dashIndex = withoutBuildMetadata.IndexOf('-');
            if (dashIndex < 0)
            {
                return ReleaseStage.Stable;
            }

            string suffix = withoutBuildMetadata.Substring(dashIndex);

            if (suffix.StartsWith("-beta.", StringComparison.OrdinalIgnoreCase) ||
                suffix.Equals("-beta", StringComparison.OrdinalIgnoreCase))
            {
                return ReleaseStage.Beta;
            }

            return ReleaseStage.Alpha;
        }

        // ###########################################################################################
        // Whether a version may be offered to a user with these two checkboxes. A stable release is
        // always allowed - the boxes only ever ADD pre-release channels, they never opt out of the
        // ordinary one.
        // ###########################################################################################
        public static bool IsVersionAllowed(string? version, bool allowAlpha, bool allowBeta)
        {
            return GetStage(version) switch
            {
                ReleaseStage.Stable => true,
                ReleaseStage.Beta => allowBeta,
                ReleaseStage.Alpha => allowAlpha,
                _ => true,
            };
        }
    }
}
