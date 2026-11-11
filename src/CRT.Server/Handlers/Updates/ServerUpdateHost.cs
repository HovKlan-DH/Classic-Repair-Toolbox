using System.Diagnostics;
using System.Text.Json;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // THE UPDATER'S REAL HANDS (2026-11-09) - the thin part ServerUpdater decides for. Linux only: it
    // runs on the box, as the crt-deploy user, under crt-server-update.service.
    //
    //   - GitHub over HTTPS, without a token: a channel's pointer every minute, and the two files of a build.
    //   - `gh attestation verify` with the bundle the release carries - offline from GitHub, so the box
    //     needs no GitHub login - pinned to this repository, the channel's workflow, main and a GitHub-hosted
    //     runner (ServerReleaseFeed). gh runs with an environment of its own: no token, nothing of ours.
    //   - systemctl for the restart, which the polkit rule allows this user for the two services alone.
    //   - "releases/current" switched by renaming a new link over it (`mv -T`), so the service never
    //     finds no link there.
    // ###########################################################################################
    internal sealed class ServerUpdateHost : IServerUpdateHost, IDisposable
    {
        private const string Systemctl = "/usr/bin/systemctl";
        private const string Move = "/usr/bin/mv";

        private static readonly string[] InstalledFolders = ["/etc/systemd/system", "/etc/polkit-1/rules.d"];

        private readonly HttpClient thisHttp;
        private readonly string thisGh;
        private readonly string thisGhHome;
        private readonly Uri thisHealth;

        public ServerUpdateHost(string ghPath, string ghHome, Uri health, string userAgent)
        {
            this.thisGh = ghPath;
            this.thisGhHome = ghHome;
            this.thisHealth = health;

            // Redirects followed - a release file is served from GitHub's storage - but only ever HTTPS.
            this.thisHttp = new HttpClient(new HttpClientHandler { AllowAutoRedirect = true, MaxAutomaticRedirections = 5 })
            {
                Timeout = TimeSpan.FromMinutes(10),
            };

            this.thisHttp.DefaultRequestHeaders.UserAgent.ParseAdd(userAgent);
        }

        public DateTimeOffset UtcNow => DateTimeOffset.UtcNow;

        public void Dispose() => this.thisHttp.Dispose();

        // ###########################################################################################
        // The channel's pointer, a few lines - a release download (github.com, then GitHub's file storage), which
        // GitHub's API limit does not count, so every minute costs nothing there. Null when GitHub has none: the
        // newest release carries no server build yet, or main's release is not there yet.
        // ###########################################################################################
        public async Task<string?> ReadPointerAsync(string url, CancellationToken cancellationToken)
        {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(30));

            using HttpResponseMessage response = await this.thisHttp.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, timeout.Token);

            if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                return null;

            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException($"GitHub answered {(int)response.StatusCode} {response.ReasonPhrase}");

            if (response.RequestMessage?.RequestUri is { } landed && landed.Scheme != Uri.UriSchemeHttps)
                throw new HttpRequestException($"[{url}] was sent somewhere that is not HTTPS.");

            if (response.Content.Headers.ContentLength > ServerReleaseFeed.PointerMaxBytes)
                throw new HttpRequestException($"[{url}] is far larger than a pointer.");

            // Read no further than a pointer can be - an answer with no length said is cut off there, never read whole
            // first (code review, 2026-11-11).
            await using Stream body = await response.Content.ReadAsStreamAsync(timeout.Token);
            using var text = new MemoryStream();
            byte[] buffer = new byte[8192];
            int read;

            while ((read = await body.ReadAsync(buffer, timeout.Token)) > 0)
            {
                if (text.Length + read > ServerReleaseFeed.PointerMaxBytes)
                    throw new HttpRequestException($"[{url}] is far larger than a pointer.");

                text.Write(buffer, 0, read);
            }

            return System.Text.Encoding.UTF8.GetString(text.GetBuffer(), 0, (int)text.Length);
        }

        public async Task DownloadAsync(string url, string toPath, long maxBytes, CancellationToken cancellationToken)
        {
            using HttpResponseMessage response = await this.thisHttp.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException($"[{url}] answered {(int)response.StatusCode} {response.ReasonPhrase}");

            if (response.RequestMessage?.RequestUri is { } landed && landed.Scheme != Uri.UriSchemeHttps)
                throw new HttpRequestException($"[{url}] was sent somewhere that is not HTTPS.");

            if (response.Content.Headers.ContentLength > maxBytes)
                throw new HttpRequestException($"[{url}] is larger than {maxBytes / 1024 / 1024} MB.");

            await using Stream body = await response.Content.ReadAsStreamAsync(cancellationToken);
            await using FileStream file = new(toPath, FileMode.CreateNew, FileAccess.Write);

            byte[] buffer = new byte[81920];
            long total = 0;
            int read;

            while ((read = await body.ReadAsync(buffer, cancellationToken)) > 0)
            {
                total += read;

                if (total > maxBytes)
                    throw new HttpRequestException($"[{url}] is larger than {maxBytes / 1024 / 1024} MB.");

                await file.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
            }
        }

        public async Task<AttestationCheck> VerifyAsync(string archivePath, string bundlePath, string channel, CancellationToken cancellationToken)
        {
            Directory.CreateDirectory(this.thisGhHome);

            var start = new ProcessStartInfo(this.thisGh)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
            };

            foreach (string argument in ServerUpdateHost.VerifyArguments(archivePath, bundlePath, channel))
                start.ArgumentList.Add(argument);

            // An environment of its own: no GitHub token reaches it (none is needed - the bundle is on disk), its
            // caches in the updater's folder, nothing interactive.
            start.Environment.Clear();
            start.Environment["PATH"] = "/usr/bin:/bin";
            start.Environment["HOME"] = this.thisGhHome;
            start.Environment["GH_CONFIG_DIR"] = Path.Combine(this.thisGhHome, "config");
            start.Environment["XDG_CONFIG_HOME"] = Path.Combine(this.thisGhHome, "config");
            start.Environment["XDG_CACHE_HOME"] = Path.Combine(this.thisGhHome, "cache");
            start.Environment["XDG_DATA_HOME"] = Path.Combine(this.thisGhHome, "data");
            start.Environment["XDG_STATE_HOME"] = Path.Combine(this.thisGhHome, "state");
            start.Environment["GH_NO_UPDATE_NOTIFIER"] = "1";
            start.Environment["GH_PROMPT_DISABLED"] = "1";
            start.Environment["NO_COLOR"] = "1";

            (int exit, string output, string error) = await ServerUpdateHost.RunAsync(start, TimeSpan.FromMinutes(5), cancellationToken);

            if (exit != 0)
                return new AttestationCheck(false, null, ServerUpdateHost.FirstLine(error, output, $"gh ended with {exit}"));

            return new AttestationCheck(true, ServerUpdateHost.CommitOf(output), string.Empty);
        }

        // ###########################################################################################
        // gh's arguments: the build and its bundle, held to THIS repository, the channel's workflow (the release
        // workflow for the live service, the push workflow for the test instance), main and a GitHub-hosted
        // runner - so a build made anywhere else, from another branch, by the other workflow or on a runner
        // somebody controls, is no build of ours.
        // ###########################################################################################
        public static IReadOnlyList<string> VerifyArguments(string archivePath, string bundlePath, string channel) =>
        [
            "attestation", "verify", archivePath,
            "--bundle", bundlePath,
            "--repo", ServerReleaseFeed.Repository,
            "--signer-workflow", ServerReleaseFeed.SignerWorkflowFor(channel),
            "--source-ref", ServerReleaseFeed.SourceRef,
            "--deny-self-hosted-runners",
            "--format", "json",
        ];

        // The commit gh says the build came from - its certificate's source repository digest.
        public static string? CommitOf(string verifyJson)
        {
            try
            {
                using JsonDocument document = JsonDocument.Parse(verifyJson);

                return document.RootElement.ValueKind == JsonValueKind.Array &&
                       document.RootElement.GetArrayLength() > 0 &&
                       document.RootElement[0].TryGetProperty("verificationResult", out JsonElement result) &&
                       result.TryGetProperty("signature", out JsonElement signature) &&
                       signature.TryGetProperty("certificate", out JsonElement certificate) &&
                       certificate.TryGetProperty("sourceRepositoryDigest", out JsonElement digest) &&
                       digest.ValueKind == JsonValueKind.String
                    ? digest.GetString()
                    : null;
            }
            catch (JsonException)
            {
                return null;
            }
        }

        public string? PointedAt(string releasesFolder)
        {
            var link = new FileInfo(Path.Combine(releasesFolder, ServerUpdater.CurrentLinkName));

            return link.LinkTarget is string target
                ? Path.GetFileName(Path.TrimEndingDirectorySeparator(target))
                : null;
        }

        public void PointAt(string releasesFolder, string version)
        {
            string current = Path.Combine(releasesFolder, ServerUpdater.CurrentLinkName);
            string next = current + ".next";

            if (File.Exists(next) || new FileInfo(next).LinkTarget != null)
                File.Delete(next);

            File.CreateSymbolicLink(next, version);

            // rename(2): the link is replaced in one step - `mv -T` renames onto it rather than into the folder it names.
            var start = new ProcessStartInfo(ServerUpdateHost.Move) { UseShellExecute = false, RedirectStandardError = true, RedirectStandardOutput = true };
            start.ArgumentList.Add("-T");
            start.ArgumentList.Add(next);
            start.ArgumentList.Add(current);

            (int exit, string output, string error) = ServerUpdateHost.RunAsync(start, TimeSpan.FromSeconds(30), CancellationToken.None).GetAwaiter().GetResult();

            if (exit != 0)
                throw new IOException($"[{current}] could not be pointed at {version} - {ServerUpdateHost.FirstLine(error, output, $"mv ended with {exit}")}");
        }

        public void LinkSettings(string buildFolder, string settingsPath)
        {
            string link = Path.Combine(buildFolder, ServerUpdater.SettingsFileName);

            if (File.Exists(link) || new FileInfo(link).LinkTarget != null)
                File.Delete(link);

            File.CreateSymbolicLink(link, settingsPath);
        }

        public async Task<string?> RestartAsync(string service, CancellationToken cancellationToken)
        {
            var start = new ProcessStartInfo(ServerUpdateHost.Systemctl) { UseShellExecute = false, RedirectStandardError = true, RedirectStandardOutput = true };
            start.ArgumentList.Add("--no-ask-password");
            start.ArgumentList.Add("restart");
            start.ArgumentList.Add(service);

            (int exit, string output, string error) = await ServerUpdateHost.RunAsync(start, TimeSpan.FromMinutes(3), cancellationToken);

            return exit == 0 ? null : ServerUpdateHost.FirstLine(error, output, $"systemctl ended with {exit}");
        }

        public async Task<string?> HealthVersionAsync(CancellationToken cancellationToken)
        {
            try
            {
                using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                timeout.CancelAfter(TimeSpan.FromSeconds(5));

                string json = await this.thisHttp.GetStringAsync(this.thisHealth, timeout.Token);
                using JsonDocument document = JsonDocument.Parse(json);

                return document.RootElement.TryGetProperty("version", out JsonElement version) && version.ValueKind == JsonValueKind.String
                    ? version.GetString()
                    : null;
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
            {
                return null;
            }
        }

        public async Task<string> ServiceStateAsync(string service, CancellationToken cancellationToken)
        {
            var start = new ProcessStartInfo(ServerUpdateHost.Systemctl) { UseShellExecute = false, RedirectStandardError = true, RedirectStandardOutput = true };
            start.ArgumentList.Add("show");
            start.ArgumentList.Add("--property=ActiveState,Result,ExecMainStatus");
            start.ArgumentList.Add(service);

            (_, string output, _) = await ServerUpdateHost.RunAsync(start, TimeSpan.FromSeconds(30), cancellationToken);

            return ServerUpdateHost.StateOf(output);
        }

        // ###########################################################################################
        // `systemctl show`'s lines in words: "active", "failed (exit 78)", "activating (exit 69)".
        // ###########################################################################################
        public static string StateOf(string showOutput)
        {
            var values = showOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(line => line.Split('=', 2))
                .Where(parts => parts.Length == 2)
                .GroupBy(parts => parts[0])
                .ToDictionary(group => group.Key, group => group.First()[1], StringComparer.Ordinal);

            string state = values.TryGetValue("ActiveState", out string? active) && active.Length > 0 ? active : "unknown";

            return values.TryGetValue("ExecMainStatus", out string? exit) && exit is { Length: > 0 } && exit != "0"
                ? $"{state} (exit {exit})"
                : state;
        }

        public Task DelayAsync(TimeSpan delay, CancellationToken cancellationToken) => Task.Delay(delay, cancellationToken);

        public byte[]? ReadInstalledCopy(string fileName)
        {
            foreach (string folder in ServerUpdateHost.InstalledFolders)
            {
                try
                {
                    string path = Path.Combine(folder, fileName);

                    if (File.Exists(path))
                        return File.ReadAllBytes(path);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // Not readable to this user (polkit's folder is root's): not known to be installed.
                }
            }

            return null;
        }

        private static async Task<(int Exit, string Output, string Error)> RunAsync(ProcessStartInfo start, TimeSpan atMost, CancellationToken cancellationToken)
        {
            using var process = Process.Start(start) ?? throw new IOException($"[{start.FileName}] could not be started.");
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(atMost);

            Task<string> output = process.StandardOutput.ReadToEndAsync(timeout.Token);
            Task<string> error = process.StandardError.ReadToEndAsync(timeout.Token);

            try
            {
                await process.WaitForExitAsync(timeout.Token);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                process.Kill(entireProcessTree: true);
                return (-1, string.Empty, $"[{Path.GetFileName(start.FileName)}] took longer than {atMost.TotalSeconds:0} seconds");
            }

            return (process.ExitCode, await output, await error);
        }

        private static string FirstLine(string error, string output, string otherwise)
        {
            string? line = (error + "\n" + output)
                .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .FirstOrDefault();

            return line == null ? otherwise : line.Length > 300 ? line[..300] : line;
        }
    }
}
