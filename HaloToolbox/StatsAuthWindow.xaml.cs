using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;

namespace HaloToolbox
{
    /// <summary>
    /// WebView2 popup that navigates to Halo Waypoint and intercepts the
    /// x-343-authorization-spartan header emitted by mccapi calls.
    /// Used by the Stats tab to obtain a live token for the HW API.
    /// </summary>
    public partial class StatsAuthWindow : Window
    {
        // Persistent WebView2 user-data so the Microsoft session is saved across runs.
        internal static readonly string WebViewDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "HaloMCCToolbox", "stats_webview2");

        public string? CapturedToken { get; private set; }

        private readonly bool _silent;
        private readonly string _gamertag;
        private bool _captured = false;
        private bool _navigatedToProfile = false;
        private readonly CancellationTokenSource _timeoutCts = new();
        private bool _closing;

        public StatsAuthWindow(string gamertag) : this(gamertag, silent: false) { }

        public StatsAuthWindow(string gamertag, bool silent)
        {
            _gamertag = gamertag;
            _silent = silent;
            InitializeComponent();

            if (silent)
            {
                WindowState   = WindowState.Minimized;
                ShowInTaskbar = false;
                ShowActivated = false;
            }

            Loaded += OnLoadedAsync;
        }

        private string ProfileUrl =>
            $"https://www.halowaypoint.com/halo-the-master-chief-collection/players/{Uri.EscapeDataString(_gamertag)}";

        private async void OnLoadedAsync(object sender, RoutedEventArgs e)
        {
            if (_silent)
            {
                _ = CloseAfterSilentTimeoutAsync(_timeoutCts.Token);
            }

            try
            {
                Directory.CreateDirectory(WebViewDataDir);
                var env = await CoreWebView2Environment.CreateAsync(null, WebViewDataDir);
                if (_closing)
                    return;

                await Browser.EnsureCoreWebView2Async(env);
                if (_closing)
                {
                    Browser.Dispose();
                    return;
                }

                Browser.CoreWebView2.AddWebResourceRequestedFilter(
                    "https://mccapi.svc.halowaypoint.com/*",
                    CoreWebView2WebResourceContext.All);

                Browser.CoreWebView2.WebResourceRequested += OnWebResourceRequested;
                Browser.CoreWebView2.NavigationCompleted  += OnNavigationCompleted;

                Browser.CoreWebView2.Navigate(ProfileUrl);
            }
            catch (Exception ex)
            {
                if (_closing)
                    return;

                if (!_silent)
                    StatusText.Text = $"WebView2 init failed: {ex.Message}";
                else
                {
                    DialogResult = false;
                    Close();
                }
            }
        }

        private async Task CloseAfterSilentTimeoutAsync(CancellationToken cancellationToken)
        {
            try
            {
                await Task.Delay(15_000, cancellationToken);
                if (!_captured && !_closing)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        if (_closing)
                            return;
                        DialogResult = false;
                        Close();
                    });
                }
            }
            catch (OperationCanceledException)
            {
                // Normal window shutdown cancels the pending timeout.
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            _closing = true;
            _timeoutCts.Cancel();
            Loaded -= OnLoadedAsync;

            try
            {
                if (Browser.CoreWebView2 is not null)
                {
                    Browser.CoreWebView2.WebResourceRequested -= OnWebResourceRequested;
                    Browser.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
                }
                Browser.Dispose();
            }
            catch
            {
                // The control may be closing while asynchronous initialization unwinds.
            }
            finally
            {
                _timeoutCts.Dispose();
            }

            base.OnClosed(e);
        }

        private void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (_captured || _navigatedToProfile) return;

            string url = Browser.CoreWebView2.Source;
            bool onHaloWaypoint = url.Contains("halowaypoint.com", StringComparison.OrdinalIgnoreCase);
            bool alreadyOnProfile = url.Contains("/players/", StringComparison.OrdinalIgnoreCase);

            if (onHaloWaypoint && !alreadyOnProfile)
            {
                _navigatedToProfile = true;
                if (!_silent)
                    StatusText.Text = "Logged in - opening player profile...";
                Browser.CoreWebView2.Navigate(ProfileUrl);
            }
        }

        private void OnWebResourceRequested(object? sender, CoreWebView2WebResourceRequestedEventArgs e)
        {
            if (_captured) return;

            if (e.Request.Headers.Contains("x-343-authorization-spartan"))
            {
                string token = e.Request.Headers.GetHeader("x-343-authorization-spartan");
                if (!string.IsNullOrWhiteSpace(token))
                {
                    _captured = true;
                    CapturedToken = token;

                    Dispatcher.InvokeAsync(async () =>
                    {
                        if (!_silent)
                        {
                            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0xC8, 0xFF));
                            StatusText.Text = "Token captured. Closing...";
                            await Task.Delay(1200);
                        }
                        DialogResult = true;
                        Close();
                    });
                }
            }
        }
    }
}
