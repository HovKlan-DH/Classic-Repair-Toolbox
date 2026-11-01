using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE A STORED LINK GOES WHEN IT IS OPENED - a link row is a DECLARED destination (a worklog's
    // link, a board's link on the Contribute tab), stored as typed, so "www.lemon64.com" is as likely
    // as "https://www.lemon64.com". Moved out of WorkbookPdfExporter (2026-10-29) when the Contribute
    // tab's Links part got "Open", so the PDF and the screen open a link by ONE rule:
    //
    //   - http:// or https:// as it stands, trimmed - and so is mailto:, the other scheme the launcher
    //     opens (code review, 2026-10-30: "mailto:anna@example.org" got https:// in front, which a
    //     browser reads as user "mailto:anna" on example.org, and opened that web site);
    //   - an address of ANY OTHER scheme ("ftp://...", "file://...") is no link CRT opens - null: with
    //     https:// in front it went somewhere else entirely;
    //   - anything else with no space in it gets https:// (a PDF hyperlink with no scheme is ignored
    //     by every reader, and the browser needs one too);
    //   - nothing, or text with a space in it, is no link at all - null.
    //
    // The launcher still decides what it will open (ExternalTargetLauncher.TryOpen).
    // ###########################################################################################
    public static class LinkTarget
    {
        public static string? For(string? url)
        {
            string trimmed = (url ?? string.Empty).Trim();

            if (trimmed.Length == 0 || trimmed.Contains(' '))
                return null;

            if (trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return trimmed;

            if (trimmed.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                return trimmed.Length > "mailto:".Length ? trimmed : null;

            // Told by "://" - "www.example.com:8080" is a host with a port, not a scheme.
            if (trimmed.Contains("://", StringComparison.Ordinal))
                return null;

            return $"https://{trimmed}";
        }
    }
}
