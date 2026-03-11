using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace HaloToolbox
{
    public partial class HaloReportWindow : Window
    {
        public string SuspectGamertag { get; set; } = "";
        public string SuspectXboxId   { get; set; } = "";
        public string CheatType       { get; set; } = "";
        public string GameTitle       { get; set; } = "";  // e.g. "Halo 3", "Halo Reach"
        public string MapName         { get; set; } = "";
        public string GameType        { get; set; } = "";
        public string GameDate        { get; set; } = "";
        public string Notes           { get; set; } = "";
        public string Scoreboard      { get; set; } = "";
        public string ZipPath         { get; set; } = "";  // full path to report ZIP

        // URL pre-selects MCC via query param
        private const string FormUrl =
            "https://support.halowaypoint.com/hc/en-us/requests/new" +
            "?ticket_form_id=360003459232" +
            "&tf_360048983931=safety__game_title__halo_mcc";

        public HaloReportWindow()
        {
            InitializeComponent();
            Loaded += async (_, _) => await InitWebViewAsync();
        }

        private async Task InitWebViewAsync()
        {
            // Use the process-wide persistent environment so cookies / login sessions
            // survive both window close/reopen and full app restarts.
            // EnsureCoreWebView2Async must be called with the environment BEFORE
            // setting WebView.Source; setting Source first triggers an implicit
            // default-environment init and would cause an ArgumentException here.
            var env = await WebViewEnvironmentManager.GetOrCreateAsync();
            await WebView.EnsureCoreWebView2Async(env);

            WebView.CoreWebView2.NavigationStarting  += (_, e) =>
                Dispatcher.Invoke(() => TxtNavStatus.Text = e.Uri);
            WebView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
            WebView.CoreWebView2.DOMContentLoaded    += OnDomContentLoaded;

            TxtSuspectPill.Text = SuspectGamertag;

            // Show the ZIP filename in the reminder bar
            if (!string.IsNullOrEmpty(ZipPath) && File.Exists(ZipPath))
            {
                var fname = Path.GetFileName(ZipPath);
                TxtAttachReminder.Text =
                    "ATTACH:  " + fname + "   (Explorer window opened with it selected -- drag it into the form below, check the box, then click Submit)";
            }
            else
            {
                TxtAttachReminder.Text =
                    "REMINDER: Build a Report ZIP first, then attach it using the file area below, check the Media Attachments box, and click Submit.";
            }

            WebView.Source = new Uri(FormUrl);
        }

        private void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                TxtNavStatus.Text = e.IsSuccess
                    ? (WebView.Source?.ToString() ?? "")
                    : "Navigation failed: " + e.WebErrorStatus;
            });
        }

        private async void OnDomContentLoaded(object? sender, CoreWebView2DOMContentLoadedEventArgs e)
        {
            // Wait for Zendesk's Nesty JS widgets to finish initialising
            await Task.Delay(2000);
            await Dispatcher.InvokeAsync(async () =>
            {
                TxtFillStatus.Text       = "Filling fields...";
                TxtFillStatus.Foreground = System.Windows.Media.Brushes.Gray;
                await RunFillAsync();
            });
        }

        private async void BtnFill_Click(object sender, RoutedEventArgs e)
        {
            TxtFillStatus.Text       = "Re-filling...";
            TxtFillStatus.Foreground = System.Windows.Media.Brushes.Gray;
            await RunFillAsync();
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(ZipPath) && File.Exists(ZipPath))
                System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + ZipPath + "\"");
        }

        // Build plain-text description
        private string BuildDescription()
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(Notes))
            {
                sb.AppendLine(Notes);
                sb.AppendLine();
            }
            sb.AppendLine("--- Game Details ---");
            sb.AppendLine("Game:      " + GameTitle);
            sb.AppendLine("Map:       " + MapName);
            sb.AppendLine("Game type: " + GameType);
            sb.AppendLine("Date/time: " + GameDate);
            sb.AppendLine();
            sb.AppendLine("--- Reported Player ---");
            sb.AppendLine("Gamertag:     " + SuspectGamertag);
            sb.AppendLine("Xbox User ID: " + SuspectXboxId);
            sb.AppendLine("(Xbox User ID is persistent and cannot be changed by renaming)");
            sb.AppendLine();
            sb.AppendLine("--- Cheat Type ---");
            sb.AppendLine(CheatType);
            sb.AppendLine();
            sb.AppendLine("--- Full Scoreboard ---");
            sb.AppendLine(Scoreboard);
            sb.AppendLine();
            sb.AppendLine("Evidence: carnage report XML and theater .mov files attached as ZIP.");
            return sb.ToString();
        }

        // JS injection using exact Zendesk field IDs from the live form HTML
        private async Task RunFillAsync()
        {
            var gamertag = EscJs(SuspectGamertag);
            var title    = EscJs("Cheating - " + CheatType + " [" + GameTitle + " - " + MapName + "]");

            // Plain text for the textarea fallback (actual \n newline characters in the JS string)
            var descText = EscJs(BuildDescription());

            // HTML for the rich-text editor: build <br>-separated HTML in C# first,
            // then EscJs it for safe embedding in a JS single-quoted string.
            // This avoids the fragile double-escape trick that produced literal \n in the editor.
            var descHtml = EscJs(
                BuildDescription()
                    .Replace("&",  "&amp;")
                    .Replace("<",  "&lt;")
                    .Replace(">",  "&gt;")
                    .Replace("\r\n", "<br>")
                    .Replace("\n",   "<br>")
                    .Replace("\r",   "<br>")
            );

            var js =
                "(function() {" +
                "  var log = [];" +

                // Helper: set a Zendesk Nesty hidden input + update visible anchor text
                "  function setNesty(id, value, label) {" +
                "    var el = document.getElementById(id);" +
                "    if (!el) { log.push(id + ': NOT FOUND'); return; }" +
                "    el.value = value;" +
                "    el.dispatchEvent(new Event('change', {bubbles:true}));" +
                "    var wrap = el.closest('.form-field') || el.parentElement;" +
                "    if (wrap) { var a = wrap.querySelector('a.nesty-input'); if (a) a.textContent = label; }" +
                "    log.push(id + ': OK');" +
                "  }" +

                // Helper: set a plain input/textarea using native setter (React-safe)
                "  function setInput(id, value) {" +
                "    var el = document.getElementById(id);" +
                "    if (!el) { log.push(id + ': NOT FOUND'); return; }" +
                "    var proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;" +
                "    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, value);" +
                "    el.dispatchEvent(new Event('input',  {bubbles:true}));" +
                "    el.dispatchEvent(new Event('change', {bubbles:true}));" +
                "    log.push(id + ': OK');" +
                "  }" +

                // 1. Select a game = Halo: The Master Chief Collection
                "  setNesty('request_custom_fields_360048983931', 'safety__game_title__halo_mcc', 'Halo: The Master Chief Collection');" +

                // 2. Behavior being reported = Cheating
                "  setNesty('request_custom_fields_360048984131', 'safety__reported_behavior__cheating', 'Cheating');" +

                // 3. Who are you reporting (plain text input)
                "  setInput('request_custom_fields_360048984151', '" + gamertag + "');" +

                // 4. Ticket title (plain text input)
                "  setInput('request_subject', '" + title + "');" +

                // 5. Description — try every TinyMCE surface in priority order
                "  (function() {" +
                "    var html  = '" + descHtml + "';" +   // proper <br>-separated HTML, already JS-escaped
                "    var plain = '" + descText + "';" +   // plain text with actual \n newlines
                "    var filled = false;" +
                "    try {" +
                // 5a. TinyMCE JS API — works for TinyMCE 4, 5, and 6
                "      if (typeof tinymce !== 'undefined') {" +
                "        var ed = tinymce.get('request_description') || tinymce.activeEditor;" +
                "        if (ed && ed.setContent) {" +
                "          ed.setContent(html);" +
                "          filled = true;" +
                "          log.push('desc: tinymce-api OK');" +
                "        }" +
                "      }" +
                // 5b. TinyMCE 4 iframe (id = request_description_ifr)
                "      if (!filled) {" +
                "        var ifr4 = document.getElementById('request_description_ifr');" +
                "        if (ifr4 && ifr4.contentDocument && ifr4.contentDocument.body) {" +
                "          ifr4.contentDocument.body.innerHTML = html;" +
                "          filled = true;" +
                "          log.push('desc: tinymce4-iframe OK');" +
                "        }" +
                "      }" +
                // 5c. TinyMCE 5+ .tox iframe
                "      if (!filled) {" +
                "        var toxIfr = document.querySelector('.tox-edit-area__iframe');" +
                "        if (toxIfr && toxIfr.contentDocument && toxIfr.contentDocument.body) {" +
                "          toxIfr.contentDocument.body.innerHTML = html;" +
                "          filled = true;" +
                "          log.push('desc: tinymce5-tox OK');" +
                "        }" +
                "      }" +
                // 5d. Textarea fallback — also always write so the hidden field has the value for form POST
                "      var ta = document.getElementById('request_description');" +
                "      if (ta) { ta.value = plain; ta.dispatchEvent(new Event('input', {bubbles:true})); }" +
                "      if (!filled) log.push('desc: textarea-only');" +
                "    } catch(ex) { log.push('desc error: ' + ex.message); }" +
                "  })();" +

                // 6. Media Attachments checkbox
                "  (function() {" +
                "    var cb = document.getElementById('request_custom_fields_1900002621047');" +
                "    if (!cb) { log.push('checkbox: NOT FOUND'); return; }" +
                "    if (!cb.checked) { cb.checked = true; cb.dispatchEvent(new Event('change', {bubbles:true})); }" +
                "    log.push('checkbox: OK');" +
                "  })();" +

                "  console.log('[HaloToolbox] ' + log.join(' | '));" +
                "  return log.join(' | ');" +
                "})();";

            try
            {
                var raw    = await WebView.CoreWebView2.ExecuteScriptAsync(js);
                var result = raw.Trim('"').Replace("\\n", "\n");
                var allOk  = !result.Contains("NOT FOUND") && !result.Contains("error");
                Dispatcher.Invoke(() =>
                {
                    TxtFillStatus.Text       = allOk
                        ? "Fields filled -- attach your ZIP and click Submit."
                        : "Partial fill (some fields may need manual entry). " + result;
                    TxtFillStatus.Foreground = allOk
                        ? System.Windows.Media.Brushes.LimeGreen
                        : System.Windows.Media.Brushes.Orange;
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    TxtFillStatus.Text       = "Fill error: " + ex.Message;
                    TxtFillStatus.Foreground = System.Windows.Media.Brushes.OrangeRed;
                });
            }
        }

        // Escape for embedding in a JS single-quoted string
        private static string EscJs(string s) =>
            s.Replace("\\", "\\\\")
             .Replace("'",  "\\'")
             .Replace("\r\n", "\\n")
             .Replace("\n",   "\\n")
             .Replace("\r",   "\\n");
    }
}
