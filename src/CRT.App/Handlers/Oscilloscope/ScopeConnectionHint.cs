using System;
using System.Net.Sockets;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // The line the Oscilloscope tab adds under a failed connect when the error alone does not say
    // what to do about it.
    //
    // macOS answers a connection into the local network with "No route to host" (EHOSTUNREACH)
    // while CRT is not allowed on it under Privacy & Security > Local Network - also on the very
    // first attempt, while macOS is still asking - and fails to resolve a ".local" name the same
    // way. Both errors also have ordinary causes (a scope on another subnet, behind a VPN, or
    // switched off), so the hint asks for the check that tells them apart: the scope answering
    // ping from Terminal. Every other platform, and every other error, gets no hint: "Connection
    // refused" means the scope answered, so the permission is not it.
    // ###########################################################################################
    public static class ScopeConnectionHint
    {
        public const string MacLocalNetwork =
            "On macOS, if the scope answers ping from Terminal, this means CRT is not allowed on the local network. " +
            "Allow it when macOS asks, or switch on Classic Repair Toolbox in System Settings > Privacy & Security > " +
            "Local Network (if it is on already, switch it off and on again), and then connect again";

        // A CRT not started from its app bundle - a build from source - runs under the Local
        // Network permission of whatever started it, and is not in that list itself.
        public const string MacLocalNetworkOutsideAppBundle =
            "On macOS, if the scope answers ping from Terminal, this means the app CRT was started from (Terminal, " +
            "or your code editor) is not allowed on the local network. Switch that app on in System Settings > " +
            "Privacy & Security > Local Network, and then connect again";

        // ###########################################################################################
        // Returns the hint for this connect failure, or null when the error message is all there is.
        // ###########################################################################################
        public static string? ForFailure(Exception exception, string host, bool isMacOS, bool runsFromAppBundle)
        {
            if (!isMacOS || exception is not SocketException socketException)
            {
                return null;
            }

            bool blockedConnect = socketException.SocketErrorCode == SocketError.HostUnreachable;
            bool blockedLocalName = socketException.SocketErrorCode == SocketError.HostNotFound && IsLocalName(host);

            if (!blockedConnect && !blockedLocalName)
            {
                return null;
            }

            return runsFromAppBundle ? MacLocalNetwork : MacLocalNetworkOutsideAppBundle;
        }

        // ###########################################################################################
        // True when the process runs from inside a macOS app bundle (".app/Contents/MacOS/").
        // ###########################################################################################
        public static bool IsInsideAppBundle(string? processPath)
        {
            return processPath != null &&
                   processPath.Contains(".app/Contents/MacOS/", StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Whether a failed connect writes its error (and hint) to the output. A connect the user
        // pressed always does. An automatic attempt retries silently, except that the first one
        // with a hint says so - otherwise a macOS user with Auto-connect ticked would only ever
        // see "Automatically connecting ...". The caller resets automaticHintWritten when a
        // session is established.
        // ###########################################################################################
        public static bool ShouldWrite(bool isAutomaticAttempt, bool hasHint, bool automaticHintWritten)
        {
            return !isAutomaticAttempt || (hasHint && !automaticHintWritten);
        }

        // ###########################################################################################
        // A multicast DNS name ("scope.local", with or without the trailing dot) - resolved on the
        // local network, so subject to the same permission as the connection.
        // ###########################################################################################
        private static bool IsLocalName(string host)
        {
            string trimmed = host.Trim().TrimEnd('.');
            return trimmed.EndsWith(".local", StringComparison.OrdinalIgnoreCase);
        }
    }
}
