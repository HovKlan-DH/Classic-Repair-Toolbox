using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHY APPROVE IS OFF for a submission this account could otherwise approve (2026-09-27) - the two
    // rules the server enforces at the approval, said before the maintainer has read the whole
    // table rather than after they press the button:
    //
    //   - a NEW system with no place in CRT's drop-down lists yet (owner request: "Before a new
    //     system can move from queue to Beta, then the maintainer must have saved the system
    //     position and the names/notes") - SystemPlacementDisplay.ReviewLine;
    //   - a system with an earlier submission still in BETA, not yet in production (owner decision:
    //     one submission in BETA per system) - CRT.Data's OneSubmissionInBeta, whose sentence the
    //     server's refusal uses too.
    //
    // Read from the two lists the Maintainer tab already holds - the drop-down listing and the
    // "Beta > Prod" list, whose systems are exactly the ones waiting in BETA. Explanation only: the
    // server refuses regardless, so a list a minute out of date costs a refusal, never a publish.
    // ###########################################################################################
    public static class ApprovalGate
    {
        public static string? Blocked(
            string? systemId,
            SystemListingAnswer? listing,
            IEnumerable<ProductionSystemRow>? waitingInBeta)
        {
            if (string.IsNullOrWhiteSpace(systemId))
                return null;

            if (waitingInBeta?.Any(row => string.Equals(row.SystemId, systemId, StringComparison.Ordinal)) == true)
                return OneSubmissionInBeta.BusyMessage(systemId);

            return SystemPlacementDisplay.ReviewLine(listing, systemId);
        }
    }
}
