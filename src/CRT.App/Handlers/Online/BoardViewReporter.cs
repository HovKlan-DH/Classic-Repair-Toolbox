using CRT;
using Handlers.DataHandling;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // SENDS BOARD VIEWS HOME (owner request, 2026-09-27) - see CRT.Data's BoardViewContract for what
    // a view is, and the Wiki's "Information collected" page for what users are told. Mandatory,
    // like the launch check-in: there is no setting to turn it off (owner decision).
    //
    // The file half and the network half of BoardViewOutbox:
    //
    //   Record          - a view counted (Main.BoardViews), written to the file at once.
    //   SendWaitingAsync - sends what waits, one report after another, until an answer says "not
    //                     now" or nothing is left. Main calls it a minute after a view is counted
    //                     (so several boards in a row go as one report) and once at launch (for
    //                     whatever an earlier run could not send). Only one send runs at a time.
    //
    // *** NOTHING HERE WAITS ON THE NETWORK FOR THE USER. *** Every failure is soft: no network, a
    // slow one, a server fault - the views stay in the file and go with the next send. A file that
    // cannot be written costs views, never the application.
    //
    // The Load/LoadFrom split is the test seam SubmissionReceiptStore and the other stores use:
    // NEVER call Load() from a test - it resolves the user's real AppData folder.
    // ###########################################################################################
    public static class BoardViewReporter
    {
        public const string FileName = "board-views.json";

        // The most reports one send delivers - a month offline is a few of them; the rest go next time.
        internal const int MaxReportsPerSend = 10;

        private static readonly object Gate = new();

        private static string _path = string.Empty;
        private static BoardViewOutbox _outbox = new();
        private static int _sending;

        // Resolves the real AppData location and loads. Called once at startup.
        public static void Load()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string directory = Path.Combine(appData, AppConfig.AppFolderName);

                Directory.CreateDirectory(directory);

                BoardViewReporter.LoadFrom(Path.Combine(directory, BoardViewReporter.FileName));
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load board views waiting to be sent: [{ex.Message}] - starting empty");
            }
        }

        // Loads from an explicit path and makes that path the save target - the test seam.
        internal static void LoadFrom(string path)
        {
            lock (BoardViewReporter.Gate)
            {
                _path = path ?? string.Empty;
                _outbox = new BoardViewOutbox();

                if (string.IsNullOrWhiteSpace(_path) || !File.Exists(_path))
                    return;

                try
                {
                    _outbox = BoardViewOutbox.FromJson(File.ReadAllText(_path));
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    Logger.Warning($"Failed to read board views waiting to be sent [{_path}]: [{ex.Message}] - starting empty");
                }
            }
        }

        // Whether any view waits to be sent - what Main's retry timer asks before sending.
        public static bool HasWaiting
        {
            get
            {
                lock (BoardViewReporter.Gate)
                    return _outbox.HasAnything;
            }
        }

        // A copy of what waits - for tests.
        internal static BoardViewOutbox OutboxForTests
        {
            get
            {
                lock (BoardViewReporter.Gate)
                    return BoardViewOutbox.FromJson(_outbox.ToJson());
            }
        }

        // A board counted as viewed, and when.
        public static void Record(string systemId, bool fromBeta, DateTimeOffset viewedUtc)
        {
            if (string.IsNullOrWhiteSpace(systemId))
                return;

            lock (BoardViewReporter.Gate)
            {
                _outbox.Add(new BoardView(systemId, viewedUtc, fromBeta));
                BoardViewReporter.Save();
            }
        }

        // ###########################################################################################
        // Sends what waits, to the real server. Never throws: it runs fire-and-forget from Main and
        // at launch, where nothing would observe an exception.
        // ###########################################################################################
        public static async Task SendWaitingAsync()
        {
            try
            {
                (string os, string osVersion, string cpu) = OnlineServices.DescribeMachine();

                await BoardViewReporter.SendWaitingAsync(
                    BoardViewReporter.PostAsync,
                    new BoardViewMachine(OnlineServices.VersionForServer, os, osVersion, cpu),
                    DateTimeOffset.UtcNow,
                    CancellationToken.None);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Sending board views failed: [{ex.Message}] - they wait for the next send");
            }
        }

        // ###########################################################################################
        // The sequencing, with the sending handed in - the test seam. Returns how many reports were
        // finished. A report is finished only by an answer that finishes it
        // (BoardViewContract.DeliveryFor); "not now" leaves it, whole, to be sent again next time.
        // ###########################################################################################
        internal static async Task<int> SendWaitingAsync(
            Func<BoardViewReport, CancellationToken, Task<BoardViewDelivery>> send,
            BoardViewMachine machine,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(send);

            if (Interlocked.Exchange(ref _sending, 1) == 1)
                return 0;

            int finished = 0;

            try
            {
                for (int round = 0; round < BoardViewReporter.MaxReportsPerSend; round++)
                {
                    BoardViewReport? report;

                    // Saved BEFORE it is sent: the report being sent is in the file, so a CRT closed
                    // mid-send sends the same batch next time rather than a new one.
                    lock (BoardViewReporter.Gate)
                    {
                        report = _outbox.NextReport(now, machine);
                        BoardViewReporter.Save();
                    }

                    if (report is null)
                        break;

                    if (await send(report, cancellationToken) == BoardViewDelivery.TryLater)
                        break;

                    lock (BoardViewReporter.Gate)
                    {
                        _outbox.Finished(report.BatchId);
                        BoardViewReporter.Save();
                    }

                    finished++;
                }
            }
            finally
            {
                Volatile.Write(ref _sending, 0);
            }

            return finished;
        }

        // One report to CRT.Server. Any failure to get an answer is "try later".
        private static async Task<BoardViewDelivery> PostAsync(BoardViewReport report, CancellationToken cancellationToken)
        {
            try
            {
                using var http = new HttpClient { Timeout = AppConfig.BoardViewTimeout };
                http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", OnlineServices.VersionForServer);

                using var content = new StringContent(BoardViewContract.ToJson(report), Encoding.UTF8, "application/json");
                using HttpResponseMessage response = await http.PostAsync(
                    $"{AppConfig.CrtServerBaseUrl}/{BoardViewContract.PathUnderApi}", content, cancellationToken);

                BoardViewDelivery delivery = BoardViewContract.DeliveryFor((int)response.StatusCode);

                Logger.Info($"Board views sent: [{report.Views?.Count ?? 0}] view(s), HTTP [{(int)response.StatusCode}] - {(delivery == BoardViewDelivery.Done ? "done" : "will try again later")}");

                return delivery;
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or IOException)
            {
                Logger.Info($"Board views not sent now: [{ex.Message}] - they wait for the next send");
                return BoardViewDelivery.TryLater;
            }
        }

        // Written beside itself and moved into place, so a crash mid-write leaves the old file whole.
        private static void Save()
        {
            if (string.IsNullOrWhiteSpace(_path))
                return;

            try
            {
                string temporary = _path + ".tmp";
                File.WriteAllText(temporary, _outbox.ToJson());
                File.Move(temporary, _path, overwrite: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Failed to save board views waiting to be sent [{_path}]: [{ex.Message}]");
            }
        }
    }
}
