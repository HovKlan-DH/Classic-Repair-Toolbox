using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Orders component labels the way a person reads them: C1, C2, C10, C105 - not the C1, C10,
    // C105, C2 that plain string ordering gives.
    //
    // A run of digits is compared as a NUMBER (so 2 < 10), a run of anything else as text,
    // case-insensitively. Leading zeros do not change a number's value, but when two runs are
    // otherwise equal the shorter one sorts first (C2 before C02), so the order is still total and
    // repeatable.
    //
    // Lifted out of BoardComponentHighlightStorage, where it was private, so the KiCad match
    // report could use the SAME ordering rather than a second copy of it. Every list of component
    // labels a person has to scan should go through this.
    // ###########################################################################################
    internal static class NaturalLabelComparer
    {
        public static readonly IComparer<string> Instance = Comparer<string>.Create(Compare);

        public static int Compare(string? left, string? right)
        {
            string leftValue = left?.Trim() ?? string.Empty;
            string rightValue = right?.Trim() ?? string.Empty;

            int leftIndex = 0;
            int rightIndex = 0;

            while (leftIndex < leftValue.Length && rightIndex < rightValue.Length)
            {
                bool leftIsDigit = char.IsDigit(leftValue[leftIndex]);
                bool rightIsDigit = char.IsDigit(rightValue[rightIndex]);

                if (leftIsDigit && rightIsDigit)
                {
                    int leftDigitStart = leftIndex;
                    int rightDigitStart = rightIndex;

                    while (leftIndex < leftValue.Length && char.IsDigit(leftValue[leftIndex]))
                    {
                        leftIndex++;
                    }

                    while (rightIndex < rightValue.Length && char.IsDigit(rightValue[rightIndex]))
                    {
                        rightIndex++;
                    }

                    string leftDigits = leftValue[leftDigitStart..leftIndex];
                    string rightDigits = rightValue[rightDigitStart..rightIndex];

                    string leftTrimmedDigits = leftDigits.TrimStart('0');
                    string rightTrimmedDigits = rightDigits.TrimStart('0');

                    if (leftTrimmedDigits.Length == 0)
                    {
                        leftTrimmedDigits = "0";
                    }

                    if (rightTrimmedDigits.Length == 0)
                    {
                        rightTrimmedDigits = "0";
                    }

                    if (leftTrimmedDigits.Length != rightTrimmedDigits.Length)
                    {
                        return leftTrimmedDigits.Length.CompareTo(rightTrimmedDigits.Length);
                    }

                    int digitCompare = string.Compare(leftTrimmedDigits, rightTrimmedDigits, StringComparison.Ordinal);
                    if (digitCompare != 0)
                    {
                        return digitCompare;
                    }

                    int originalDigitLengthCompare = leftDigits.Length.CompareTo(rightDigits.Length);
                    if (originalDigitLengthCompare != 0)
                    {
                        return originalDigitLengthCompare;
                    }

                    continue;
                }

                if (leftIsDigit != rightIsDigit)
                {
                    return leftIsDigit ? -1 : 1;
                }

                int leftTextStart = leftIndex;
                int rightTextStart = rightIndex;

                while (leftIndex < leftValue.Length && !char.IsDigit(leftValue[leftIndex]))
                {
                    leftIndex++;
                }

                while (rightIndex < rightValue.Length && !char.IsDigit(rightValue[rightIndex]))
                {
                    rightIndex++;
                }

                string leftText = leftValue[leftTextStart..leftIndex];
                string rightText = rightValue[rightTextStart..rightIndex];

                int textCompare = string.Compare(leftText, rightText, StringComparison.OrdinalIgnoreCase);
                if (textCompare != 0)
                {
                    return textCompare;
                }
            }

            return leftValue.Length.CompareTo(rightValue.Length);
        }
    }
}
