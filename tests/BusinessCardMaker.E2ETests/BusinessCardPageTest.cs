using System.Diagnostics;
using System.Net.Sockets;
using Microsoft.Playwright;
using Microsoft.Playwright.NUnit;
using NUnit.Framework;

namespace BusinessCardMaker.E2ETests;

/// <summary>
/// Base class for CardMaker E2E tests
/// Starts the web server once and provides helper methods for waiting on interactive pages
/// </summary>
public abstract class BusinessCardPageTest : PageTest
{
    private static Process? _serverProcess;
    private static string? _baseUrl;
    private static readonly List<string> _serverLogs = new();

    protected static string BaseUrl => _baseUrl ?? "http://localhost:5255";

    [OneTimeSetUp]
    public async Task StartServerIfNeeded()
    {
        if (_serverProcess != null)
        {
            return;
        }

        var port = GetFreeTcpPort();
        _baseUrl = $"http://localhost:{port}";

        var webProjectPath = Path.GetFullPath(Path.Combine(
            TestContext.CurrentContext.TestDirectory,
            "..", "..", "..", "..", "..", "src", "BusinessCardMaker.Web"));

        var startInfo = new ProcessStartInfo
        {
            FileName = "dotnet",
            Arguments = $"run --project \"{webProjectPath}\" --no-build --urls {_baseUrl}",
            WorkingDirectory = webProjectPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        startInfo.EnvironmentVariables["ASPNETCORE_ENVIRONMENT"] = "Development";
        startInfo.EnvironmentVariables["DISABLE_RATE_LIMITER"] = "true";
        startInfo.EnvironmentVariables["E2E_FAKE_IMPORT"] = "true";

        _serverProcess = Process.Start(startInfo);
        if (_serverProcess == null)
        {
            throw new InvalidOperationException("Failed to start CardMaker server for E2E tests.");
        }

        _serverProcess.OutputDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
            {
                _serverLogs.Add(args.Data);
            }
        };
        _serverProcess.ErrorDataReceived += (_, args) =>
        {
            if (!string.IsNullOrWhiteSpace(args.Data))
            {
                _serverLogs.Add(args.Data);
            }
        };
        _serverProcess.BeginOutputReadLine();
        _serverProcess.BeginErrorReadLine();

        // Wait for the server to respond
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
        var started = false;
        for (var i = 0; i < 120 && !started; i++)
        {
            if (_serverProcess.HasExited)
            {
                var log = string.Join(Environment.NewLine, _serverLogs.TakeLast(40));
                throw new InvalidOperationException(
                    $"CardMaker server exited early with code {_serverProcess.ExitCode}. " +
                    $"Recent log:{Environment.NewLine}{log}");
            }

            try
            {
                var response = await client.GetAsync(BaseUrl);
                if ((int)response.StatusCode < 500)
                {
                    var html = await response.Content.ReadAsStringAsync();
                    started = response.IsSuccessStatusCode || (int)response.StatusCode < 400 || html.Length > 0;
                    if (started) break;
                }
            }
            catch
            {
                // ignore and retry
            }
            await Task.Delay(1000);
        }

        if (!started)
        {
            StopServer();
            var log = string.Join(Environment.NewLine, _serverLogs.TakeLast(40));
            throw new InvalidOperationException(
                $"CardMaker server did not become ready for E2E tests. " +
                $"Recent log:{Environment.NewLine}{log}");
        }
    }

    [OneTimeTearDown]
    public void StopServer()
    {
        if (_serverProcess == null)
        {
            return;
        }

        try
        {
            if (!_serverProcess.HasExited)
            {
                _serverProcess.Kill(entireProcessTree: true);
                _serverProcess.WaitForExit(TimeSpan.FromSeconds(10));
            }
        }
        finally
        {
            _serverProcess.Dispose();
            _serverProcess = null;
        }
    }

    [SetUp]
    public void SetUpDefaults()
    {
        // Give Blazor time to hydrate
        Page.SetDefaultTimeout(30000);
    }

    private static int GetFreeTcpPort()
    {
        var listener = new TcpListener(System.Net.IPAddress.Loopback, 0);
        listener.Start();
        var port = ((System.Net.IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }

    /// <summary>
    /// Wait for the home page to become interactive (Blazor hydration complete)
    /// </summary>
    protected async Task WaitForHomeInteractiveAsync()
    {
        var response = await Page.GotoAsync(BaseUrl, new() { WaitUntil = WaitUntilState.NetworkIdle, Timeout = 60000 });
        if (response != null && response.Status >= 400)
        {
            throw new InvalidOperationException($"Home page returned status code {response.Status}");
        }

        await Page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 60000 });

        // Wait for main heading
        var heading = Page.Locator("h1");
        await Expect(heading).ToBeVisibleAsync(new() { Timeout = 60000 });
        await Expect(heading).ToContainTextAsync("CardMaker", new() { Timeout = 60000 });

        // Wait for template download button to indicate Blazor is interactive
        var downloadButton = Page.Locator("button:has-text('Excel Template')");
        await downloadButton.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 60000 });
    }
}
