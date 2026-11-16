using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW LONG A CREDIT'S SHORT FIELDS MAY BE (owner request, 2026-11-15: "limit the text length for both the "What" input
    // field ... as it was not the intention to state everything in there, as this should be a one-liner only ... The same for
    // the "Which part of it""; asked: "What" 60, "Which part of it" 60, "Name or handle" 50, "Contact" 100 - eight published
    // contacts are web addresses of 57 to 60 characters, and the longest published "What" and "Which part of it" are 33).
    //
    // Held on a credit the board has not published as it stands - the main contribution's rule (BoardRowRequired,
    // PublishedCredits): what is published stays as it is. An ERROR in both scopes (BoardDataChecks' "credit.too_long"), so the
    // table marks the cell, the draft's lists count it, Submit sends nothing while it is there and the server refuses it; the
    // component pane's boxes take no more (MaxLength). Counted as the value is kept: trimmed, in UTF-16 code units.
    //
    // *** AND THE MAIN CONTRIBUTION AT MOST 500 CHARACTERS (owner decision, 2026-11-16: "500 characters? Shouldn't this be
    // enough to state the most important things? ... let us start with this") - counted as its WORDS, the text a person reads
    // (RichNote.PlainOf): its formatting marks are no characters of it, as the pane's box counts them.
    // ###########################################################################################
    public static class CreditLimits
    {
        public const int Category = 60;
        public const int SubCategory = 60;
        public const int NameOrHandle = 50;
        public const int Contact = 100;
        public const int MainContribution = 500;

        // The columns of `row` longer than they may be - each with its limit and its length. None unless `judged`: the row does
        // not stand as published, which only its list can say (PublishedCredits.Judged - each published credit standing for one
        // row; none at the finalise, whose rows were judged as the submission was created).
        public static IReadOnlyList<(string Column, int Limit, int Length)> TooLong(CreditEntry row, bool judged)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (!judged)
            {
                return [];
            }

            var found = new List<(string Column, int Limit, int Length)>();

            void Check(string? value, string column, int limit)
            {
                int length = value?.Trim().Length ?? 0;

                if (length > limit)
                {
                    found.Add((column, limit, length));
                }
            }

            Check(row.Category, BoardWorkbookSchema.ColCategory, CreditLimits.Category);
            Check(row.SubCategory, BoardWorkbookSchema.ColSubCategory, CreditLimits.SubCategory);
            Check(row.NameOrHandle, BoardWorkbookSchema.ColNameOrHandle, CreditLimits.NameOrHandle);
            Check(row.Contact, BoardWorkbookSchema.ColContact, CreditLimits.Contact);
            Check(RichNote.PlainOf(row.MainContribution), BoardWorkbookSchema.ColMainContribution, CreditLimits.MainContribution);

            return found;
        }
    }
}
