using System.Text;

namespace CRT.Server.Handlers.Health
{
    // ###########################################################################################
    // THE SERVER'S OWN CHANGELOG (owner request, 2026-11-24: 'In the "Maintainer" tab and under "Account", then
    // please have a "Changelog" like there is in "About" ... so I can check what has changed in server also').
    //
    // It is RELEASE-NOTES.md - the one-liners per server version that are also each version's release text in
    // HovKlan-DH/CRT-Server - EMBEDDED by CRT.Server.csproj, so the program answering carries the notes of its own
    // build: the test instance the notes of main, the live service those of the release it runs. Nothing is read
    // from disk, and nothing can be out of step with what is deployed.
    //
    // What is sent is the file from its first version heading on: the HTML comment at its top is the rules for
    // whoever writes it, and CRT's Markdown reader shows raw HTML as typed. Read once - it cannot change while the
    // process runs. Sent with GET /api/review/server-status (CRT.Data's ServerStatusAnswer.Changelog), never with
    // the public /api/health.
    // ###########################################################################################
    public static class ServerReleaseNotes
    {
        // The name CRT.Server.csproj gives the embedded RELEASE-NOTES.md.
        internal const string ResourceName = "CRT.Server.RELEASE-NOTES.md";

        private static readonly Lazy<string?> Embedded = new(() => ServerReleaseNotes.ShownPart(ServerReleaseNotes.ReadEmbedded()));

        // The changelog this build carries, as CRT shows it - null when it carries none (never expected).
        public static string? Text => ServerReleaseNotes.Embedded.Value;

        // ###########################################################################################
        // The notes without the comment that may open the file: everything after its "-->", trimmed. A file with
        // no such comment is sent as it stands; nothing left is null, so CRT says nothing rather than show an empty
        // changelog. An opened comment that never closes leaves nothing - the rest of it is the comment.
        // ###########################################################################################
        internal static string? ShownPart(string? markdown)
        {
            string text = (markdown ?? string.Empty).TrimStart('﻿').Replace("\r\n", "\n", StringComparison.Ordinal).Trim();

            if (text.StartsWith("<!--", StringComparison.Ordinal))
            {
                int end = text.IndexOf("-->", 4, StringComparison.Ordinal);
                text = end < 0 ? string.Empty : text[(end + 3)..].Trim();
            }

            return text.Length == 0 ? null : text;
        }

        private static string? ReadEmbedded()
        {
            using Stream? stream = typeof(ServerReleaseNotes).Assembly.GetManifestResourceStream(ServerReleaseNotes.ResourceName);

            if (stream is null)
                return null;

            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return reader.ReadToEnd();
        }
    }
}
