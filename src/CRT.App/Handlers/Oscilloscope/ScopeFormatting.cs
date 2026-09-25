using System;
using System.Globalization;
using System.IO;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // Number, voltage and time formatting for the oscilloscope panel, plus the small string
    // helpers around window titles, captured image file names and serial masking.
    //
    // Extracted from TabOscilloscope: none of this touches a control, so it belongs here where it
    // can be tested. The engineering-suffix rules are deliberately quirky (see the tests) - time
    // uses capital "S" for seconds while voltage uses capital "V", and both fall back to plain
    // "u" rather than a micro sign so the SCPI output panel stays ASCII.
    // ###########################################################################################
    public static class ScopeFormatting
    {
        // ###########################################################################################
        // Formats a double for use inside an SCPI command. "G15" keeps enough significant digits for
        // scope resolution, and the exponent is lower-cased because some firmware rejects "E".
        // ###########################################################################################
        public static string FormatScpiNumber(double value)
        {
            return value.ToString("G15", CultureInfo.InvariantCulture).Replace("E", "e", StringComparison.Ordinal);
        }

        // ###########################################################################################
        // Formats a voltage value into a compact engineering string.
        // ###########################################################################################
        public static string FormatVoltage(double volts)
        {
            double absoluteValue = Math.Abs(volts);

            if (absoluteValue >= 1.0)
            {
                return $"{volts.ToString("0.###", CultureInfo.InvariantCulture)}V";
            }

            if (absoluteValue >= 0.001)
            {
                return $"{(volts * 1000.0).ToString("0.###", CultureInfo.InvariantCulture)}mV";
            }

            return $"{(volts * 1000000.0).ToString("0.###", CultureInfo.InvariantCulture)}uV";
        }

        // ###########################################################################################
        // Formats a time value into a compact engineering string.
        // ###########################################################################################
        public static string FormatTime(double seconds)
        {
            double absoluteValue = Math.Abs(seconds);

            if (absoluteValue >= 1.0)
            {
                return $"{seconds.ToString("0.###", CultureInfo.InvariantCulture)}S";
            }

            if (absoluteValue >= 0.001)
            {
                return $"{(seconds * 1000.0).ToString("0.###", CultureInfo.InvariantCulture)}mS";
            }

            if (absoluteValue >= 0.000001)
            {
                return $"{(seconds * 1000000.0).ToString("0.###", CultureInfo.InvariantCulture)}uS";
            }

            return $"{(seconds * 1000000000.0).ToString("0.###", CultureInfo.InvariantCulture)}nS";
        }

        // ###########################################################################################
        // Returns the next trigger level when stepping up or down by a fixed 0.25V grid.
        //
        // A level already sitting on the grid moves a whole step; one between grid lines snaps
        // outwards to the next line in the direction of travel, so the first keypress from an
        // arbitrary scope-reported level always lands on the grid rather than jumping past it.
        // ###########################################################################################
        public static double GetNextSnappedTriggerLevelVolts(double currentTriggerLevelVolts, int direction)
        {
            const double triggerLevelStepVolts = 0.25;
            const double stepTolerance = 1e-6;

            double scaledValue = currentTriggerLevelVolts / triggerLevelStepVolts;
            double nearestWholeStep = Math.Round(scaledValue, MidpointRounding.AwayFromZero);
            bool isNearWholeStep = Math.Abs(scaledValue - nearestWholeStep) <= stepTolerance;

            if (direction > 0)
            {
                double targetStep = isNearWholeStep
                    ? nearestWholeStep + 1.0
                    : Math.Ceiling(scaledValue);

                return targetStep * triggerLevelStepVolts;
            }

            double downTargetStep = isNearWholeStep
                ? nearestWholeStep - 1.0
                : Math.Floor(scaledValue);

            return downTargetStep * triggerLevelStepVolts;
        }

        // The two windows that carry a connection suffix - the main window and the component info
        // popup - both build and strip it from here, so the strings cannot drift apart.
        private const string ConnectedSuffix = " (oscilloscope connected)";
        private const string DisconnectedSuffix = " (oscilloscope disconnected)";

        // ###########################################################################################
        // Strips any oscilloscope connection suffix from the main window title so the caller can
        // re-append the current one without the suffixes stacking up.
        // ###########################################################################################
        public static string GetMainWindowTitleBase(string windowTitle)
        {
            if (windowTitle.EndsWith(ConnectedSuffix, StringComparison.Ordinal))
            {
                return windowTitle[..^ConnectedSuffix.Length];
            }

            if (windowTitle.EndsWith(DisconnectedSuffix, StringComparison.Ordinal))
            {
                return windowTitle[..^DisconnectedSuffix.Length];
            }

            return windowTitle;
        }

        // ###########################################################################################
        // Builds a window title carrying the oscilloscope connection suffix. This is the inverse of
        // GetMainWindowTitleBase above, and the pair share the suffix constants.
        //
        // isOscilloscopeTabEnabled wins over everything else: with the oscilloscope tab switched off
        // there is no auto-connect running and no session to report, so the caller gets its base
        // title back untouched even if a session was live when the tab was hidden.
        //
        // shouldReportSessionState is the caller's own rule for when a state is worth reporting. The
        // main window passes "a session has existed, OR auto-connect is enabled and still trying", so
        // a user waiting for the scope to come up on its own can see that it has not yet. The popup
        // passes only the former - it has nothing useful to say about a connection that has never
        // happened.
        // ###########################################################################################
        public static string BuildOscilloscopeWindowTitle(
            string baseTitle,
            bool isOscilloscopeTabEnabled,
            bool shouldReportSessionState,
            bool hasEstablishedSession)
        {
            if (!isOscilloscopeTabEnabled || !shouldReportSessionState)
            {
                return baseTitle;
            }

            return baseTitle + (hasEstablishedSession ? ConnectedSuffix : DisconnectedSuffix);
        }

        // ###########################################################################################
        // Makes one part of a captured oscilloscope image file name safe for the filesystem.
        // Blank input becomes "Unknown" so the resulting name never collapses to nothing.
        // ###########################################################################################
        public static string SanitizeCapturedOscilloscopeImageFileNamePart(string value)
        {
            string sanitized = string.IsNullOrWhiteSpace(value) ? "Unknown" : value.Trim();

            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                sanitized = sanitized.Replace(invalidChar, '_');
            }

            return sanitized;
        }

        // ###########################################################################################
        // Masks a scope serial number so logs and the output panel do not expose the real value.
        // The length is preserved so a truncated response still looks obviously different.
        // ###########################################################################################
        public static string MaskScopeSerial(string serial)
        {
            return string.IsNullOrEmpty(serial)
                ? string.Empty
                : new string('*', serial.Length);
        }

        // ###########################################################################################
        // Masks only the serial field inside a standard *IDN? response while leaving the other parts
        // unchanged for debugging and display purposes.
        // ###########################################################################################
        public static string MaskIdentifyResponseSerial(string response)
        {
            var parts = (response ?? string.Empty).Split(',');

            if (parts.Length > 2)
            {
                string trimmedSerial = parts[2].Trim();
                parts[2] = MaskScopeSerial(trimmedSerial);
            }

            return string.Join(",", parts);
        }
        // ###########################################################################################
        // Normalizes oscilloscope overlay values so the final unit character is uppercase and the
        // preceding unit character, when alphabetic, is lowercase.
        // Examples: 1US -> 1uS, 5MV -> 5mV, 1.5v -> 1.5V
        // ###########################################################################################
        public static string NormalizeScopeOverlayValue(string? value)
        {
            string trimmed = value?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return string.Empty;
            }

            char[] chars = trimmed.ToCharArray();
            int lastIndex = chars.Length - 1;

            if (char.IsLetter(chars[lastIndex]))
            {
                chars[lastIndex] = char.ToUpperInvariant(chars[lastIndex]);
            }

            int secondLastIndex = lastIndex - 1;
            if (secondLastIndex >= 0 && char.IsLetter(chars[secondLastIndex]))
            {
                chars[secondLastIndex] = char.ToLowerInvariant(chars[secondLastIndex]);
            }

            return new string(chars);
        }
    }
}
