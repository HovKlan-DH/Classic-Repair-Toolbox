using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The logging seam for CRT.Data. This library is shared by the desktop app, the review app
    // and (from Phase 3) the server, and none of those hosts share a single logging mechanism -
    // the app's own Handlers.DataHandling.Logger is a static singleton that writes to a file path
    // resolved from Velopack's AppData folder, which makes no sense for a server process.
    //
    // CRT.Data therefore never calls a concrete logger directly. It calls the static gateway
    // below, which defaults to writing nothing (Sink is null until a host installs one). Each
    // host installs its own ICrtLog implementation once at startup:
    //   - CRT.App:    CrtLog.Sink = new AppLoggerAdapter();  // forwards to Handlers.DataHandling.Logger
    //   - CRT.Server: CrtLog.Sink = new ServerLogAdapter();  // forwards to whatever the server uses
    //
    // Per CLAUDE.md's test rules, no test may call Logger.Initialize() - and by construction, no
    // test needs to: CrtLog.Sink stays null (the no-op default) throughout the whole CRT.Data.Tests
    // and CRT.App.Tests suites, so every Logger.* call inside a moved class (BoardDataReader,
    // BoardDataWriter, BoardComponentHighlightStorage, DataValidator) is inert during tests,
    // exactly as it always has been.
    // ###########################################################################################
    public interface ICrtLog
    {
        void Debug(string message);
        void Info(string message);
        void Warning(string message);
        void Critical(string message);
    }

    // ###########################################################################################
    // Static gateway so call sites inside CRT.Data read almost exactly as they did when they
    // called Handlers.DataHandling.Logger directly (CrtLog.Warning(...) vs Logger.Warning(...)) -
    // deliberately, so moving these classes here was a path-and-namespace change, not a rewrite of
    // every call site's shape. Sink is null (no-op) until a host installs one.
    // ###########################################################################################
    public static class CrtLog
    {
        public static ICrtLog? Sink { get; set; }

        public static void Debug(string message) => Sink?.Debug(message);
        public static void Info(string message) => Sink?.Info(message);
        public static void Warning(string message) => Sink?.Warning(message);
        public static void Critical(string message) => Sink?.Critical(message);
    }
}
