using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHO IS ASKING, and which systems they review (Phase 6 roles, 2026-09-25).
    //
    // The account alone no longer answers the authority question: a maintainer is somebody in one
    // or more systems' pools, and that is a set of rows, not a flag. Loading the set ONCE per
    // request, here, and handing it to every rule means no route can forget to look it up - and
    // it is looked up from the database on every request, so removing a maintainer takes effect on
    // their very next call (Phase 6's definition of done).
    //
    // An administrator's set is whatever rows happen to exist and is never consulted - the
    // administrator is in every pool by definition (see accounts.is_administrator in
    // 0001_initial.sql), and ReviewAuthority checks the flag first.
    // ###########################################################################################
    public sealed record ReviewAccess(AccountRecord Account, IReadOnlySet<string> MaintainerOf)
    {
        public static ReviewAccess For(AccountRecord account, IEnumerable<string>? maintainerOf = null)
        {
            ArgumentNullException.ThrowIfNull(account);

            return new ReviewAccess(account, new HashSet<string>(maintainerOf ?? [], StringComparer.Ordinal));
        }
    }
}
