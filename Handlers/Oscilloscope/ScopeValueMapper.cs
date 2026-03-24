using Handlers.DataHandling;
using System;
using System.Globalization;
using System.Linq;

namespace Handlers.Oscilloscope
{
    public sealed class ScopeMappedValue
    {
        public string RawValue { get; init; } = string.Empty;
        public string MatchedDisplayValue { get; init; } = string.Empty;
        public double NumericValue { get; init; }
        public string ScpiValue { get; init; } = string.Empty;
    }

    public static class ScopeValueMapper
    {
        private delegate bool ValueParser(string text, out double value);

        // ###########################################################################################
        // Maps a board Excel T/DIV value to a supported oscilloscope T/DIV value and SCPI number.
        // ###########################################################################################
        public static bool TryMapTimeDiv(
            ComponentImageEntry componentImageEntry,
            OscilloscopeEntry oscilloscopeEntry,
            out ScopeMappedValue mappedValue)
        {
            mappedValue = new ScopeMappedValue();

            if (string.IsNullOrWhiteSpace(componentImageEntry.TimeDiv))
            {
                return false;
            }

            if (!TryParseTimeValue(componentImageEntry.TimeDiv, out double seconds))
            {
                return false;
            }

            if (!TryMatchSupportedValue(
                componentImageEntry.TimeDiv,
                seconds,
                oscilloscopeEntry.TimeDivList,
                TryParseTimeValue,
                out mappedValue))
            {
                return false;
            }

            return true;
        }

        // ###########################################################################################
        // Maps a board Excel V/DIV value to a supported oscilloscope V/DIV value and SCPI number.
        // ###########################################################################################
        public static bool TryMapVoltsDiv(
            ComponentImageEntry componentImageEntry,
            OscilloscopeEntry oscilloscopeEntry,
            out ScopeMappedValue mappedValue)
        {
            mappedValue = new ScopeMappedValue();

            if (string.IsNullOrWhiteSpace(componentImageEntry.VoltsDiv))
            {
                return false;
            }

            if (!TryParseVoltageValue(componentImageEntry.VoltsDiv, out double volts))
            {
                return false;
            }

            if (!TryMatchSupportedValue(
                componentImageEntry.VoltsDiv,
                volts,
                oscilloscopeEntry.VoltsDivList,
                TryParseVoltageValue,
                out mappedValue))
            {
                return false;
            }

            return true;
        }

        // ###########################################################################################
        // Maps a board Excel trigger level value to a SCPI numeric voltage value.
        // ###########################################################################################
        public static bool TryMapTriggerLevel(
            ComponentImageEntry componentImageEntry,
            out ScopeMappedValue mappedValue)
        {
            mappedValue = new ScopeMappedValue();

            if (string.IsNullOrWhiteSpace(componentImageEntry.TriggerLevelVolts))
            {
                return false;
            }

            if (!TryParseVoltageValue(componentImageEntry.TriggerLevelVolts, out double volts))
            {
                return false;
            }

            mappedValue = new ScopeMappedValue
            {
                RawValue = componentImageEntry.TriggerLevelVolts,
                MatchedDisplayValue = componentImageEntry.TriggerLevelVolts.Trim(),
                NumericValue = volts,
                ScpiValue = FormatScpiNumber(volts)
            };

            return true;
        }

        // ###########################################################################################
        // Parses a time text such as 5ms, 100us, 5ns or 1s into seconds.
        // ###########################################################################################
        public static bool TryParseTimeValue(string text, out double seconds)
        {
            seconds = 0;

            if (!TrySplitEngineeringValue(text, out double value, out string unit))
            {
                return false;
            }

            switch (NormalizeUnit(unit))
            {
                case "s":
                    seconds = value;
                    return true;

                case "ms":
                    seconds = value / 1000.0;
                    return true;

                case "us":
                    seconds = value / 1000000.0;
                    return true;

                case "ns":
                    seconds = value / 1000000000.0;
                    return true;

                default:
                    return false;
            }
        }

        // ###########################################################################################
        // Parses a voltage text such as 1V, 500mV, 0.5V or 2mV into volts.
        // ###########################################################################################
        public static bool TryParseVoltageValue(string text, out double volts)
        {
            volts = 0;

            if (!TrySplitEngineeringValue(text, out double value, out string unit))
            {
                return false;
            }

            switch (NormalizeUnit(unit))
            {
                case "v":
                    volts = value;
                    return true;

                case "mv":
                    volts = value / 1000.0;
                    return true;

                case "uv":
                    volts = value / 1000000.0;
                    return true;

                default:
                    return false;
            }
        }

        // ###########################################################################################
        // Matches a parsed board value against the supported values list from the main Excel file.
        // ###########################################################################################
        private static bool TryMatchSupportedValue(
            string rawValue,
            double numericValue,
            string supportedValuesCsv,
            ValueParser parser,
            out ScopeMappedValue mappedValue)
        {
            mappedValue = new ScopeMappedValue();

            if (string.IsNullOrWhiteSpace(supportedValuesCsv))
            {
                return false;
            }

            var supportedValues = supportedValuesCsv
                .Split(',')
                .Select(value => value.Trim())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList();

            foreach (string supportedValue in supportedValues)
            {
                if (!parser(supportedValue, out double supportedNumericValue))
                {
                    continue;
                }

                if (AreEquivalent(numericValue, supportedNumericValue))
                {
                    mappedValue = new ScopeMappedValue
                    {
                        RawValue = rawValue.Trim(),
                        MatchedDisplayValue = supportedValue,
                        NumericValue = supportedNumericValue,
                        ScpiValue = FormatScpiNumber(supportedNumericValue)
                    };
                    return true;
                }
            }

            return false;
        }

        // ###########################################################################################
        // Splits an engineering value into its numeric and unit parts using invariant parsing.
        // ###########################################################################################
        private static bool TrySplitEngineeringValue(string text, out double value, out string unit)
        {
            value = 0;
            unit = string.Empty;

            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            string trimmed = text.Trim();
            int unitStart = 0;

            while (unitStart < trimmed.Length &&
                   (char.IsDigit(trimmed[unitStart]) ||
                    trimmed[unitStart] == '.' ||
                    trimmed[unitStart] == '-' ||
                    trimmed[unitStart] == '+'))
            {
                unitStart++;
            }

            if (unitStart == 0 || unitStart >= trimmed.Length)
            {
                return false;
            }

            string numericPart = trimmed.Substring(0, unitStart).Trim();
            unit = trimmed.Substring(unitStart).Trim();

            return double.TryParse(
                numericPart,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out value);
        }

        // ###########################################################################################
        // Normalizes engineering unit text so Excel variations like uS, us and µs compare equally.
        // ###########################################################################################
        private static string NormalizeUnit(string unit)
        {
            return unit
                .Trim()
                .Replace("µ", "u", StringComparison.Ordinal)
                .ToLowerInvariant();
        }

        // ###########################################################################################
        // Compares two numeric values with a small tolerance for engineering notation matching.
        // ###########################################################################################
        private static bool AreEquivalent(double left, double right)
        {
            double delta = Math.Abs(left - right);
            double scale = Math.Max(1.0, Math.Max(Math.Abs(left), Math.Abs(right)));
            return delta <= scale * 1e-9;
        }

        // ###########################################################################################
        // Formats a numeric SCPI value using invariant culture and lower-case exponent notation.
        // ###########################################################################################
        private static string FormatScpiNumber(double value)
        {
            return value.ToString("G15", CultureInfo.InvariantCulture)
                .Replace("E", "e", StringComparison.Ordinal);
        }
    }
}