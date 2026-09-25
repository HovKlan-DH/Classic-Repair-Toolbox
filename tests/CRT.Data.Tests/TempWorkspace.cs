namespace ClassicRepairToolbox.Tests;

// A throwaway directory for tests that touch the filesystem. Everything the suite writes goes
// under one of these and is deleted afterwards, so no test ever reaches the real data root or
// the user's AppData folder.
//
// Note on logging: Handlers.Logger writes nothing until Logger.Initialize() is called, and no
// test calls it. That is why the Logger.Warning(...) calls inside the classes under test are
// inert here - do NOT call Logger.Initialize() from a test, or the suite starts writing to the
// user's real log file.
public sealed class TempWorkspace : IDisposable
{
    public string Root { get; }

    public TempWorkspace()
    {
        this.Root = Path.Combine(Path.GetTempPath(), "crt-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(this.Root);
    }

    /// <summary>Absolute path inside the workspace. Parent directories are created.</summary>
    public string Path_(params string[] parts)
    {
        string full = System.IO.Path.Combine(new[] { this.Root }.Concat(parts).ToArray());
        string? directory = System.IO.Path.GetDirectoryName(full);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return full;
    }

    /// <summary>Writes a file inside the workspace and returns its absolute path.</summary>
    public string WriteFile(string relativePath, string content)
    {
        string full = this.Path_(relativePath);
        File.WriteAllText(full, content);
        return full;
    }

    public string ReadFile(string relativePath) => File.ReadAllText(this.Path_(relativePath));

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(this.Root))
            {
                Directory.Delete(this.Root, recursive: true);
            }
        }
        catch
        {
            // A leftover temp folder is not worth failing a test run over.
        }
    }
}
