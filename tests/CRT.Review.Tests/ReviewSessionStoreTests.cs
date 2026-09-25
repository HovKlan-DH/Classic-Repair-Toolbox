using System.Text.Json;
using CRT.Review.Handlers;

namespace CRT.Review.Tests;

// ###########################################################################################
// Covers ReviewSessionStore - staying signed in between launches (maintainer request,
// 2026-09-22), and the three things that must still clear it: signing out, the account being
// locked or invalidated, and expiry.
//
// *** THESE TESTS ARE WINDOWS-ONLY BY NATURE, AND SAY SO RATHER THAN PRETENDING OTHERWISE. ***
// The token is encrypted with DPAPI, which does not exist on Linux, so on the CI runner
// (ubuntu-latest per CLAUDE.md) nothing can be stored at all. Every test that needs a stored
// session therefore asserts the REFUSAL off Windows and the round trip on it - which is the real
// contract, not a workaround: "stores nothing rather than storing plaintext" is the security
// property, and it deserves to be pinned on the platform where it applies.
//
// Uses its own temp folder per test rather than the real AppData path. ReviewSessionStore.
// Initialise() resolves the user's actual folder, so no test may call it - InitialiseAt is the
// seam, exactly as UserSettings.LoadFrom is in CRT.App.
// ###########################################################################################
// *** THE COLLECTION IS NOT OPTIONAL. *** ReviewSessionStore holds its path in a STATIC field, so
// two of these tests running concurrently would point it at each other's temp folder and fail
// intermittently while passing in isolation. CLAUDE.md records that exact diagnosis being reached
// the slow way once already (the tell is a test that never fails alone), so this class is pinned
// to its own collection from the start. Any future test class touching this store must join it.
[Collection("ReviewSessionStore")]
public sealed class ReviewSessionStoreTests : IDisposable
{
    private readonly string thisFolder;

    private static readonly DateTimeOffset Now = new(2026, 9, 22, 12, 0, 0, TimeSpan.Zero);

    public ReviewSessionStoreTests()
    {
        this.thisFolder = Path.Combine(Path.GetTempPath(), "crt-review-session-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(this.thisFolder);

        ReviewSessionStore.InitialiseAt(Path.Combine(this.thisFolder, ReviewSessionStore.SessionFileName));
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(this.thisFolder, recursive: true);
        }
        catch
        {
            // A locked temp file must not fail a test run.
        }
    }

    private string FilePath => Path.Combine(this.thisFolder, ReviewSessionStore.SessionFileName);

    private static ReviewSession Session(DateTimeOffset? expires = null) =>
        new(
            "a-very-secret-bearer-token",
            expires ?? ReviewSessionStoreTests.Now.AddDays(30),
            AccountId: 7,
            Email: "dennis@example.com",
            DisplayName: "Dennis");

    [Fact]
    public void A_remembered_session_comes_back_on_the_next_launch()
    {
        // The whole point of the feature: sign in once, and the next launch is already signed in.
        ReviewSessionStore.Remember(ReviewSessionStoreTests.Session());

        ReviewSession? recalled = ReviewSessionStore.Recall(ReviewSessionStoreTests.Now);

        if (!ReviewSessionStore.IsSupported)
        {
            // Off Windows nothing is stored, deliberately. See the class header.
            Assert.Null(recalled);
            return;
        }

        Assert.NotNull(recalled);
        Assert.Equal("a-very-secret-bearer-token", recalled!.BearerToken);
        Assert.Equal("dennis@example.com", recalled.Email);
        Assert.Equal("Dennis", recalled.DisplayName);
        Assert.Equal(7, recalled.AccountId);
    }

    [Fact]
    public void The_TOKEN_is_never_written_in_the_clear()
    {
        // *** THE SECURITY PROPERTY, ASSERTED AGAINST THE ACTUAL BYTES ON DISK. *** A regression
        // that serialised the plain token would still round-trip perfectly and pass every other
        // test in this file. This one reads the file as text and insists the secret is not in it.
        ReviewSessionStore.Remember(ReviewSessionStoreTests.Session());

        if (!ReviewSessionStore.IsSupported)
        {
            Assert.False(File.Exists(this.FilePath));
            return;
        }

        string onDisk = File.ReadAllText(this.FilePath);

        Assert.DoesNotContain("a-very-secret-bearer-token", onDisk, StringComparison.Ordinal);

        // *** AND THE STORED FIELD IS DECODED BEFORE BEING CHECKED. *** Searching the raw file
        // text alone is NOT enough, and this was proved by sabotage: replacing the encryption with
        // a plaintext fallback still passed, because base64 of the token does not contain the
        // token's own characters. A regression that "protected" the secret with nothing but
        // base64 would be invisible to the assertion above.
        using JsonDocument document = JsonDocument.Parse(onDisk);

        byte[] stored = Convert.FromBase64String(
            document.RootElement.GetProperty("protectedToken").GetString()!);

        Assert.DoesNotContain(
            "a-very-secret-bearer-token",
            System.Text.Encoding.UTF8.GetString(stored),
            StringComparison.Ordinal);

        // The non-secret fields ARE plain, on purpose - so a person debugging the file can read
        // what it is about without tooling.
        Assert.Contains("dennis@example.com", onDisk, StringComparison.Ordinal);
    }

    [Fact]
    public void SIGNING_OUT_forgets_it()
    {
        // Named by the maintainer as a thing that must clear the login.
        ReviewSessionStore.Remember(ReviewSessionStoreTests.Session());

        ReviewSessionStore.Forget();

        Assert.Null(ReviewSessionStore.Recall(ReviewSessionStoreTests.Now));

        // *** DELETED, NOT BLANKED. *** An empty file is indistinguishable from a truncated write,
        // and the next launch would have to guess which it was.
        Assert.False(File.Exists(this.FilePath));
    }

    [Fact]
    public void An_EXPIRED_stored_session_is_refused_AND_deleted()
    {
        // Returning it would make the app open on the queue, fire a request, take a 401 and drop
        // back to sign-in - which reads as a bug rather than as an expired session. Deleting it
        // stops the same dead token being retried on every future launch.
        ReviewSessionStore.Remember(
            ReviewSessionStoreTests.Session(ReviewSessionStoreTests.Now.AddDays(-1)));

        Assert.Null(ReviewSessionStore.Recall(ReviewSessionStoreTests.Now));
        Assert.False(File.Exists(this.FilePath));
    }

    [Fact]
    public void A_session_expiring_WITHIN_THE_MARGIN_is_refused()
    {
        // ReviewSession.IsUsableAt keeps a one-minute margin so a token cannot expire while a
        // request is in flight - which the user would see as the app logging them out at random.
        // Recall must honour it rather than applying its own rule.
        ReviewSessionStore.Remember(
            ReviewSessionStoreTests.Session(ReviewSessionStoreTests.Now.AddSeconds(30)));

        Assert.Null(ReviewSessionStore.Recall(ReviewSessionStoreTests.Now));
    }

    [Fact]
    public void A_CORRUPT_file_yields_nothing_and_is_cleared_rather_than_throwing()
    {
        // Hand-edited, truncated by a bad shutdown, or written by a future version. None of these
        // may crash the app on launch or block a normal sign-in.
        File.WriteAllText(this.FilePath, "this is not json at all");

        Assert.Null(ReviewSessionStore.Recall(ReviewSessionStoreTests.Now));
        Assert.False(File.Exists(this.FilePath));
    }

    [Fact]
    public void A_file_whose_CIPHERTEXT_is_not_ours_yields_nothing()
    {
        // What a file copied from ANOTHER MACHINE or another Windows account looks like: valid
        // JSON, valid base64, and undecryptable here. That is the property the encryption is for,
        // so it is asserted rather than assumed.
        File.WriteAllText(
            this.FilePath,
            JsonSerializer.Serialize(new
            {
                protectedToken = Convert.ToBase64String("not really dpapi output"u8.ToArray()),
                expiresUtc = ReviewSessionStoreTests.Now.AddDays(30),
                accountId = 7,
                email = "dennis@example.com",
                displayName = "Dennis"
            }));

        Assert.Null(ReviewSessionStore.Recall(ReviewSessionStoreTests.Now));
        Assert.False(File.Exists(this.FilePath));
    }

    [Fact]
    public void Recalling_with_NO_file_is_simply_nothing()
    {
        // The very first launch. Not an error, and nothing to clean up.
        Assert.Null(ReviewSessionStore.Recall(ReviewSessionStoreTests.Now));
    }

    [Fact]
    public void Forgetting_when_there_is_nothing_to_forget_is_harmless()
    {
        // Reached whenever a 401 arrives on a launch that never stored anything - which is the
        // ordinary case for a reviewer who has not signed in yet.
        ReviewSessionStore.Forget();
        ReviewSessionStore.Forget();

        Assert.False(File.Exists(this.FilePath));
    }

    [Fact]
    public void Remembering_TWICE_replaces_rather_than_appending()
    {
        // Signing in as a different account must not leave the previous one recoverable.
        ReviewSessionStore.Remember(ReviewSessionStoreTests.Session());

        ReviewSessionStore.Remember(new ReviewSession(
            "a-second-token",
            ReviewSessionStoreTests.Now.AddDays(30),
            AccountId: 9,
            Email: "someone.else@example.com",
            DisplayName: "Someone Else"));

        ReviewSession? recalled = ReviewSessionStore.Recall(ReviewSessionStoreTests.Now);

        if (!ReviewSessionStore.IsSupported)
        {
            Assert.Null(recalled);
            return;
        }

        Assert.Equal("a-second-token", recalled!.BearerToken);
        Assert.Equal(9, recalled.AccountId);
        Assert.DoesNotContain(
            "dennis@example.com", File.ReadAllText(this.FilePath), StringComparison.Ordinal);
    }

    [Fact]
    public void With_NO_store_configured_nothing_is_written_and_nothing_throws()
    {
        // Initialise() swallows a failure to create the AppData folder by leaving the path empty.
        // Every later call must then be inert rather than throwing on a null path.
        ReviewSessionStore.InitialiseAt(string.Empty);

        ReviewSessionStore.Remember(ReviewSessionStoreTests.Session());

        Assert.Null(ReviewSessionStore.Recall(ReviewSessionStoreTests.Now));

        ReviewSessionStore.Forget();
    }
}
