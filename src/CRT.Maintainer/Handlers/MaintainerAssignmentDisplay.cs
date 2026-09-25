using System;
using System.Globalization;
using System.Linq;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // How the administrator's "Maintainers" screen reads (Phase 6 roles, 2026-09-25).
    //
    // Pure, so the wording is tested rather than eyeballed - the same rule ReviewQueueDisplay
    // follows. Two things here are decisions rather than formatting: an account that CANNOT be
    // granted says why in the list itself (so the refusal is read before the button is pressed,
    // not after), and a system with nobody assigned says so in words, because an empty bracket
    // reads as a rendering fault and an unassigned board is exactly the thing the administrator
    // opens this screen to find.
    // ###########################################################################################
    public static class MaintainerAssignmentDisplay
    {
        // ###########################################################################################
        // "Commodore / C64 / 250407  -  2 maintainers". The three parts with spaces, because a
        // system id with slashes reads as a path rather than as a board.
        // ###########################################################################################
        public static string SystemLine(ReviewSystemRow system)
        {
            ArgumentNullException.ThrowIfNull(system);

            string name = string.Join(
                " / ",
                new[] { system.Manufacturer, system.Hardware, system.Board }
                    .Where(part => !string.IsNullOrWhiteSpace(part)));

            if (name.Length == 0)
                name = string.IsNullOrWhiteSpace(system.SystemId) ? "(unknown system)" : system.SystemId;

            return $"{name}  -  {MaintainerAssignmentDisplay.CountPhrase(system.Maintainers.Count)}";
        }

        // ###########################################################################################
        // "Anna (anna@example.com)". The address is there because two people can share a display
        // name, and the address is what the administrator actually knows them by.
        // ###########################################################################################
        public static string MaintainerLine(MaintainerRow maintainer)
        {
            ArgumentNullException.ThrowIfNull(maintainer);

            return MaintainerAssignmentDisplay.NameAndAddress(maintainer.DisplayName, maintainer.Email);
        }

        // ###########################################################################################
        // One choice in the "add a maintainer" list. An account the server would refuse is still
        // LISTED - hiding it would leave the administrator wondering where somebody went - but
        // it says why, in the list, before anything is pressed.
        // ###########################################################################################
        public static string AccountChoice(ReviewAccountRow account)
        {
            ArgumentNullException.ThrowIfNull(account);

            string line = MaintainerAssignmentDisplay.NameAndAddress(account.DisplayName, account.Email);

            string? why = MaintainerAssignmentDisplay.WhyNotGrantable(account);

            return why is null ? line : $"{line}  -  {why}";
        }

        // ###########################################################################################
        // Mirrors MaintainerAssignmentRules on the server, so the list can say what the server will
        // say. The server still decides; this only saves a round trip that would end in a refusal.
        // ###########################################################################################
        public static string? WhyNotGrantable(ReviewAccountRow account)
        {
            ArgumentNullException.ThrowIfNull(account);

            if (account.IsAdministrator)
                return "administrator, reviews every system already";

            if (!account.IsVerified)
                return "address not verified yet";

            if (account.IsLocked)
                return "locked";

            return null;
        }

        public static string CountPhrase(int count) =>
            count switch
            {
                0 => "nobody assigned",
                1 => "1 maintainer",
                _ => $"{count.ToString(CultureInfo.InvariantCulture)} maintainers"
            };

        private static string NameAndAddress(string? displayName, string? email)
        {
            string name = string.IsNullOrWhiteSpace(displayName) ? "(no name)" : displayName.Trim();

            return string.IsNullOrWhiteSpace(email) ? name : $"{name} ({email.Trim()})";
        }
    }
}
