using CRT;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// HandlerGuard - every async event handler of the Maintainer tab runs its work through it (code
// review, 2026-10-13). An exception thrown after an await in an `async void` handler reaches the
// dispatcher, where App lets it end CRT on purpose; through the guard it is logged and said where
// the screen says its errors, and the handler returns. Logger writes nothing in a test (it is
// never initialised here), so only what is SAID is asserted.
// ###########################################################################################
public sealed class HandlerGuardTests
{
    [Fact]
    public async Task A_handler_that_throws_after_an_await_is_said_and_does_not_throw_on()
    {
        string? said = null;

        Task Work() => ThrowAfterAwaitAsync("The list was not where it was expected.");

        await HandlerGuard.RunAsync(Work, message => said = message);

        Assert.Equal(MaintainerWaitWording.Unexpected("The list was not where it was expected."), said);
    }

    // Work that finishes says nothing - the guard is invisible until something goes wrong.
    [Fact]
    public async Task A_handler_that_finishes_says_nothing()
    {
        string? said = null;
        bool ran = false;

        await HandlerGuard.RunAsync(async () => { await Task.Yield(); ran = true; }, message => said = message);

        Assert.True(ran);
        Assert.Null(said);
    }

    // Saying it must not be the next thing to end CRT: a `show` that throws too is swallowed (the
    // failure is in the log), and no `show` at all is allowed.
    [Fact]
    public async Task Neither_a_failing_show_nor_none_at_all_lets_the_exception_out()
    {
        await HandlerGuard.RunAsync(() => ThrowAfterAwaitAsync("first"), _ => throw new InvalidOperationException("second"));
        await HandlerGuard.RunAsync(() => ThrowAfterAwaitAsync("first"), show: null);
    }

    // ###########################################################################################
    // *** WORK POSTED TO THE DISPATCHER IS GUARDED TOO (code review, 2026-10-14). *** An async lambda
    // posted with Dispatcher.UIThread.Post is an "async void" with no caller, exactly like an event
    // handler: opening a two-approval submission posted the Files view's read unguarded, while its
    // sibling three methods further down went through HandlerGuard. Every such post in the tab's
    // code must be a HandlerGuard.RunAsync.
    // ###########################################################################################
    [Fact]
    public void Every_async_lambda_the_tab_posts_to_the_dispatcher_runs_through_the_guard()
    {
        string folder = Path.Combine(HandlerGuardTests.AppFolder(), "Tabs", "Maintainer");

        var unguarded = Directory.GetFiles(folder, "*.cs", SearchOption.AllDirectories)
            .SelectMany(file => System.Text.RegularExpressions.Regex
                .Matches(File.ReadAllText(file), @"\.Post\(\s*async\b[^;]*;")
                .Where(post => !post.Value.Contains("HandlerGuard.RunAsync", StringComparison.Ordinal))
                .Select(post => $"{Path.GetFileName(file)}: {post.Value}"))
            .ToList();

        Assert.True(unguarded.Count == 0, string.Join(Environment.NewLine, unguarded));
    }

    private static string AppFolder()
    {
        string relative = Path.Combine("src", "CRT.App");
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, relative, "Tabs")))
            directory = directory.Parent;

        Assert.NotNull(directory);

        return Path.Combine(directory.FullName, relative);
    }

    private static async Task ThrowAfterAwaitAsync(string message)
    {
        await Task.Yield();
        throw new InvalidOperationException(message);
    }
}
