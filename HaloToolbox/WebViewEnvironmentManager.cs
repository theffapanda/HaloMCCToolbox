using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;

namespace HaloToolbox
{
    /// <summary>
    /// Process-wide singleton for the WebView2 CoreWebView2Environment.
    ///
    /// WHY THIS EXISTS:
    ///   WebView2 requires that every WebView2 control in the same process that
    ///   uses the same user-data folder share the exact same CoreWebView2Environment
    ///   *object instance*.  If you call CoreWebView2Environment.CreateAsync() more
    ///   than once — even with identical arguments — you get two different objects,
    ///   and passing the second one to EnsureCoreWebView2Async throws:
    ///     "WebView2 was already initialized with a different CoreWebView2Environment."
    ///
    ///   By caching the instance here and always handing out the same object, we
    ///   avoid that exception no matter how many times HaloReportWindow is opened.
    ///
    /// PERSISTENCE:
    ///   The user-data folder lives at:
    ///     %LocalAppData%\HaloMCCToolbox\WebView2
    ///   Cookies, local storage, and login sessions written there survive across
    ///   window close/reopen AND across full app restarts, so the user only has to
    ///   sign in to Halo Waypoint once.
    /// </summary>
    internal static class WebViewEnvironmentManager
    {
        // Stable per-user location:  C:\Users\<user>\AppData\Local\HaloMCCToolbox\WebView2
        public static readonly string UserDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "HaloMCCToolbox",
            "WebView2");

        private static CoreWebView2Environment? _env;
        private static readonly SemaphoreSlim   _lock = new(1, 1);

        /// <summary>
        /// Returns the shared environment, creating it on first call.
        /// Safe to call concurrently — creation is serialised by a semaphore.
        /// </summary>
        public static async Task<CoreWebView2Environment> GetOrCreateAsync()
        {
            if (_env != null)
                return _env;

            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_env == null)
                {
                    var options = new CoreWebView2EnvironmentOptions
                    {
                        // Allow Windows / Microsoft-account SSO so the user may be
                        // signed in automatically on machines where their Windows login
                        // is linked to their Xbox / Microsoft account.
                        AllowSingleSignOnUsingOSPrimaryAccount = true,
                    };

                    _env = await CoreWebView2Environment.CreateAsync(
                        browserExecutableFolder: null,   // use system WebView2 runtime
                        userDataFolder:          UserDataFolder,
                        options:                 options
                    ).ConfigureAwait(false);
                }

                return _env;
            }
            finally
            {
                _lock.Release();
            }
        }
    }
}
